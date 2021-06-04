'''
Windows 已知 asycio 問題:
1. 不支援 stdin, stdout 的 StreamReader, StreamWriter
2. 不支援 add_signal_handler, 不能處理 SIGINT, SIGTERM

SKCOM.dll 已知狀況:
1. SKCenterLib_Login() 只能在 main thread 執行
2. SKCenterLib_Login() 執行第二次時, 會回傳錯誤 2003, 表示已登入
3. signal.signal() 只能在 main thread 執行
4. SKQuoteLib_EnterMonitor() 可以在 child thread 執行
5. SKQuoteLib_EnterMonitor() 失敗後執行第二次時, 看起來沒有作用, 而且不會結束
6. SKQuoteLib_EnterMonitor() 成功後執行第二次時, 有作用 (self.debug_retry())
7. SKQuoteLib_EnterMonitor() 本身具備斷線重連機制, 不需要重新登入和連線

Retry 待測試項目:
1. LOGIN_FAILED - ok!
2. MONITOR_FAILED
3. MISSING_CONN
'''

import asyncio
import enum
import os
import os.path
import re
import signal
import sys
import threading

import pythoncom
from comtypes import COMError
import comtypes.client
import comtypes.gen.SKCOMLib as sk

from skcom.helper import load_config
from skcom.exception import ConfigException

class ReceiverState(enum.Enum):
    IDLE = enum.auto()
    LOGIN = enum.auto()
    LOGIN_DONE = enum.auto()
    LOGIN_FAILED = enum.auto()
    MONITOR = enum.auto()
    MONITOR_DONE = enum.auto()
    MONITOR_FAILED = enum.auto()
    MISSING_CONN = enum.auto()
    RETRY = enum.auto()
    STOP = enum.auto()
    STOP_DONE = enum.auto()

class AsyncQuoteReceiver():

    def __init__(self):
        # 延遲參數, 測試狀態變化時微調用
        self.RETRY_LIMIT = 3
        self.DELAY_LOGIN_DONE = 0
        self.DELAY_RETRY = 3
        self.DELAY_PUMP = 0.5

        self.skc = None
        self.skq = None
        self.skr = None
        self.log_path = os.path.expanduser(r'~\.skcom\logs\capital')
        self.connect_result_event = None
        self.state = ReceiverState.IDLE
        self.retry_count = 0

        # 產生 log 目錄
        if not os.path.isdir(self.log_path):
            os.makedirs(self.log_path)

        try:
            self.config = load_config()
        except ConfigException as ex:
            if not ex.loaded:
                print(ex)
            sys.exit(1)

    def ctrl_c(self, sig, frm):
        if self.state not in [ReceiverState.STOP, ReceiverState.STOP_DONE]:
            print('偵測到 Ctrl+C, 結束監聽', flush=True)
            print('目前工作階段', self.state, flush=True)
            self.stop()

    def start(self):
        signal.signal(signal.SIGINT, self.ctrl_c)

        # 啟動非同步迴路
        asyncio.run(self.root_task())

    async def root_task(self):
        # 載入 COM 元件
        try:
            self.skr = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
            skh0 = comtypes.client.GetEvents(self.skr, self)
            self.skc = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
            self.skc.SKCenterLib_SetLogPath(self.log_path)
            self.skq = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
            skh1 = comtypes.client.GetEvents(self.skq, self)
        except COMError as ex:
            self.logger.error('init() 發生不預期狀況')
            self.logger.error(ex)

        while self.state in [ReceiverState.IDLE, ReceiverState.RETRY]:
            if self.state is ReceiverState.RETRY:
                print('root_task() retry_count=%d' % self.retry_count)
                self.state = ReceiverState.IDLE
            
            await asyncio.gather(
                self.pump(),    # 更新 COM 事件
                self.request(), # 處理請求
                self.connect(), # 連線
            )
        print('root_task() done.', flush=True)

    async def pump(self):
        while self.state not in [ReceiverState.RETRY, ReceiverState.STOP]:
            pythoncom.PumpWaitingMessages()
            await asyncio.sleep(self.DELAY_PUMP)
        print('pump() done.', flush=True)

    async def connect(self):
        print('Login', flush=True)
        self.state = ReceiverState.LOGIN
        n_code = self.skc.SKCenterLib_Login(self.config['account'], self.config['password'])
        # 注意! 錯誤碼 2003 表示已登入, 這種情況也要放行
        if n_code not in [0, 2003]:
            self.state = ReceiverState.LOGIN_FAILED
            self.handle_sk_error('Login()', n_code)
            await self.retry()
            print('connect() cancelled.', flush=True)
            return
        print('Login - ok!', flush=True)
        self.state = ReceiverState.LOGIN_DONE

        # 刻意停頓 n 秒, 用來測試切斷 Wifi 之後, EnterMonitor() 的邏輯
        await asyncio.sleep(self.DELAY_LOGIN_DONE)

        def target():
            print('Enter monitor', flush=True)
            self.state = ReceiverState.MONITOR
            n_code = self.skq.SKQuoteLib_EnterMonitor()
            if n_code != 0:
                self.state = ReceiverState.MONITOR_FAILED
                self.handle_sk_error('EnterMonitor()', n_code)
                async def do_retry():
                    await self.retry()
                asyncio.run(do_retry())
                return
            print('Enter monitor - ok!', flush=True)

        if self.state in [ReceiverState.RETRY, ReceiverState.STOP]:
            print('connect() cancelled.', flush=True)
        
        threading.Thread(target=target).start()
        print('connect() done.', flush=True)

    async def request(self):
        # 等待 EnterMonitor() 結果
        self.connect_result_event = asyncio.Event()
        await self.connect_result_event.wait()

        # EnterMonitor() 失敗, 取消聽牌
        if self.state != ReceiverState.MONITOR_DONE:
            print('request() cancelled.', flush=True)
            return

        (page_no, n_code) = self.skq.SKQuoteLib_RequestTicks(-1, "5530")
        if n_code != 0:
            self.handle_sk_error('RequestTicks()', n_code)
        
        # TODO
        while self.state not in [ReceiverState.RETRY, ReceiverState.STOP]:
            await asyncio.sleep(1)
        print('request() done.', flush=True)

    async def retry(self, inc = True):
        if inc:
            self.retry_count += 1

        if self.retry_count > self.RETRY_LIMIT:
            self.stop()
        else:
            print('%d 秒後重試' % self.DELAY_RETRY, flush=True)
            await asyncio.sleep(self.DELAY_RETRY)
            self.state = ReceiverState.RETRY
            self.connect_result_event.set()

    def debug_retry(self):
        n_code = self.skq.SKQuoteLib_LeaveMonitor()
        if n_code != 0:
            self.handle_sk_error('LeaveMonitor()', n_code)
        self.state = ReceiverState.RETRY
        self.connect_result_event.set()

    def stop(self):
        if self.state == ReceiverState.MONITOR:
            print('等待 EnterMonitor() 完成')
            # 如果在這個階段執行 self.skq.SKQuoteLib_LeaveMonitor(), 會觸發例外
            #   OSError: exception: access violation reading 0x0000000000000008
            # 可以看看 API 有沒有提供取消功能

        if self.state == ReceiverState.MONITOR_DONE:
            print('Leave monitor', flush=True)
            n_code = self.skq.SKQuoteLib_LeaveMonitor()
            if n_code != 0:
                self.handle_sk_error('LeaveMonitor()', n_code)

        self.state = ReceiverState.STOP
        self.connect_result_event.set()

    def handle_sk_error(self, action, n_code):
        skmsg = self.skc.SKCenterLib_GetReturnCodeMessage(n_code)
        print('執行動作 [%s] 時發生錯誤, 詳細原因: #%d %s' % (action, n_code, skmsg))

    def OnReplyMessage(self, bstrUserID, bstrMessage):
        return 0xffff
    
    def OnConnection(self, nKind, nCode):
        if nCode != 0:
            # 這裡的 nCode 沒有對應的文字訊息
            action = '狀態變更 %d' % nKind
            self.handle_sk_error(action, nCode)

        # 參考文件: 6. 代碼定義表 (p.170)
        # 3001 已連線
        # 3002 正常斷線
        # 3003 已就緒
        # 3021 異常斷線
        msg = '無法識別'
        if nKind == 3001:
            msg = '連線成功'
        if nKind == 3003:
            self.state = ReceiverState.MONITOR_DONE
            self.retry_count = 0
            self.connect_result_event.set()
            msg = '連線就緒'
        if nKind == 3002:
            self.connect_result_event.set()
            msg = '結束連線'
        if nKind == 3021:
            msg = '異常斷線'
        
        print('{}, nKind={}, nCode={}'.format(msg, nKind, nCode), flush=True)

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        time_str = "{}".format(nTimehms)
        time_pretty = re.sub(r'(\d{2})(\d{2})(\d{2})', r'\1:\2:\3', time_str)
        print("[{}] {:2.2f}".format(time_pretty, nClose / 100))
        self.debug_retry()

def main():
    AsyncQuoteReceiver().start()

if __name__ == "__main__":
    main()
