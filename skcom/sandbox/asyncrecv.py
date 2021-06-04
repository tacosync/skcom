'''
非同步聽牌機實驗程式

SKCOM.dll 已知狀況:
1. SKCenterLib_Login() 只能在 main thread 執行
2. SKCenterLib_Login() 執行第二次時, 會回傳錯誤 2003, 表示已登入
3. signal.signal() 只能在 main thread 執行
4. SKQuoteLib_EnterMonitor() 可以在 child thread 執行
5. SKQuoteLib_EnterMonitor() 失敗後執行第二次時, skcom.dll 需要 2.13.29.0 才有作用
6. SKQuoteLib_EnterMonitor() 成功後執行第二次時, 有作用 (self.debug_retry())
7. SKQuoteLib_EnterMonitor() 本身具備斷線重連機制, 不需要重新登入和連線

Retry 待測試項目:
1. LOGIN_FAILED - ok!
2. MONITOR_FAILED - ok!
3. MISSING_CONN (需要自幹連線檢查程式)

Windows 已知 asycio 問題:
1. 不支援 stdin, stdout 的 StreamReader, StreamWriter
2. 不支援 add_signal_handler, 不能處理 SIGINT, SIGTERM
'''

import asyncio
import enum
import logging
import os
import os.path
import re
import signal
import sys
import threading
import time

import pythoncom
from comtypes import COMError
import comtypes.client
import comtypes.gen.SKCOMLib as sk

from skcom.helper import load_config
from skcom.exception import ConfigException

logger = logging.getLogger('skcom')

class ReceiverState(enum.Enum):
    """ 非同步聽牌機生命週期狀態 """
    IDLE = enum.auto()
    LOGIN = enum.auto()
    LOGIN_DONE = enum.auto()
    LOGIN_FAILED = enum.auto()
    MONITOR = enum.auto()
    MONITOR_DONE = enum.auto()
    MONITOR_FAILED = enum.auto()
    MISSING_CONN = enum.auto()   # 目前未使用
    RETRY = enum.auto()
    STOP = enum.auto()
    STOP_DONE = enum.auto()

class AsyncQuoteReceiver():
    """ 非同步聽牌機 """

    def __init__(self, debug=False):
        """ 非同步聽牌機初始配置 """

        if debug:
            logger.setLevel('DEBUG')

        # 延遲參數, 測試狀態變化時微調用
        self.RETRY_LIMIT = 3
        self.REPLAY_LIMIT = 1
        self.FLUSH_INTERVAL = 3
        self.DELAY_LOGIN_DONE = 5 # 想在 LOGIN_DONE <-> MONITOR 之間 Ctrl+C, 延長這個時間
        self.DELAY_RETRY = 3
        self.DELAY_PUMP = 0.5
        self.STOCK_ID = '2345'
        self.STATE_LOG_LEVEL = 'INFO'

        # 群益 API COM 元件
        self.skc = None # 登入 API
        self.skq = None # 報價 API
        self.skr = None # 回報 API

        # skcom.dll log 路徑
        self.log_path = os.path.expanduser(r'~\.skcom\logs\capital')

        # 供 request() 等待連線就緒或失敗的是件
        self.monitor_event = None

        # 生命週期狀態
        self.state = ReceiverState.IDLE

        # 連線重試次數
        self.retry_count = 0
        self.replay_count = 0

        # 產生 log 目錄
        if not os.path.isdir(self.log_path):
            os.makedirs(self.log_path)

        # 載入 yaml 設定
        try:
            self.config = load_config()
        except ConfigException as ex:
            if not ex.loaded:
                logger.info(ex)
            sys.exit(1)

    def ctrl_c(self, sig, frm):
        """ Ctrl+C 處理 """
        if self.state not in [ReceiverState.STOP, ReceiverState.STOP_DONE]:
            logger.info('偵測到 Ctrl+C, 結束監聽')
            logger.info('目前工作階段 %s', self.state)
            self.stop()

    def start(self):
        """ 使用 asyncio 啟動非同步聽牌作業 """
        asyncio.run(self.root_task())

    async def root_task(self):
        """ 非同步聽牌作業進入點 """
        logger.debug('root_task(): begin')

        # 接收 Ctrl+C
        signal.signal(signal.SIGINT, self.ctrl_c)

        # 載入 COM 元件
        try:
            self.skr = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
            skh0 = comtypes.client.GetEvents(self.skr, self)
            self.skc = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
            self.skc.SKCenterLib_SetLogPath(self.log_path)
            self.skq = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
            skh1 = comtypes.client.GetEvents(self.skq, self)
        except COMError as ex:
            logger.error('init() 發生不預期狀況')
            logger.error(ex)

        # 啟動與重試聽牌作業
        while self.state in [ReceiverState.IDLE, ReceiverState.RETRY]:
            if self.state is ReceiverState.RETRY:
                logger.debug('root_task(): retry_count=%d', self.retry_count)
                self.change_state(ReceiverState.IDLE)
            
            await asyncio.gather(
                self.pump(),    # 更新 COM 事件
                self.request(), # 處理請求
                self.connect(), # 連線
            )

        self.change_state(ReceiverState.STOP_DONE)
        logger.debug('root_task(): done')

    async def pump(self):
        """ 推送 COM 元件事件 """
        logger.debug('pump(): begin')

        prev = time.time()
        while self.state not in [ReceiverState.RETRY, ReceiverState.STOP]:
            pythoncom.PumpWaitingMessages()
            await asyncio.sleep(self.DELAY_PUMP)
            interval = time.time() - prev
            if interval > self.FLUSH_INTERVAL:
                logger.debug('pump(): flush stdout')
                sys.stdout.flush()
                prev = time.time()

        logger.debug('pump(): done')

    async def connect(self):
        """ 建立連線 """
        logger.debug('connect(): begin')

        # 登入
        # 注意! 錯誤碼 2003 表示已登入, 這種情況也要放行
        self.change_state(ReceiverState.LOGIN)
        n_code = self.skc.SKCenterLib_Login(self.config['account'], self.config['password'])
        if n_code not in [0, 2003]:
            self.change_state(ReceiverState.LOGIN_FAILED)
            self.handle_sk_error('Login()', n_code)
            await self.retry()
            return
        self.change_state(ReceiverState.LOGIN_DONE)

        # 刻意停頓 n 秒, 用來測試切斷 Wifi 之後, EnterMonitor() 的邏輯
        await asyncio.sleep(self.DELAY_LOGIN_DONE)

        # 啟動監聽器的作業細節
        def target():
            self.change_state(ReceiverState.MONITOR)
            n_code = self.skq.SKQuoteLib_EnterMonitor()
            if n_code != 0:
                self.change_state(ReceiverState.MONITOR_FAILED)
                self.monitor_event.set()
                self.handle_sk_error('EnterMonitor()', n_code)
                async def do_retry():
                    await self.retry()
                asyncio.run(do_retry())
                return
            logger.debug('connect(): SKQuoteLib_EnterMonitor() done')

        # 登入失敗或 Ctrl+C, 放棄監聽作業
        if self.state in [ReceiverState.RETRY, ReceiverState.STOP]:
            logger.debug('connect(): cancelled')
        
        # 以 child thread 啟動監聽器
        threading.Thread(target=target).start()
        logger.debug('connect(): done')

    async def request(self):
        """ 設定監聽項目 """
        logger.debug('request(): begin')

        # 等待 EnterMonitor() 結果
        self.monitor_event = asyncio.Event()
        await self.monitor_event.wait()

        # EnterMonitor() 失敗, 取消聽牌
        if self.state != ReceiverState.MONITOR_DONE:
            logger.info('request() cancelled.')
            return

        (page_no, n_code) = self.skq.SKQuoteLib_RequestTicks(-1, self.STOCK_ID)
        if n_code != 0:
            self.handle_sk_error('RequestTicks()', n_code)

        logger.debug('request(): done')

    async def retry(self, inc = True):
        """ 重試聽牌流程 """
        if inc:
            self.retry_count += 1

        if self.retry_count > self.RETRY_LIMIT:
            self.stop()
        else:
            logger.debug('retry(): %d 秒後重試' % self.DELAY_RETRY)
            await asyncio.sleep(self.DELAY_RETRY)
            self.change_state(ReceiverState.RETRY)

    def replay(self):
        """ 監聽成功時, 強制重試聽牌流程 """
        n_code = self.skq.SKQuoteLib_LeaveMonitor()
        if n_code != 0:
            self.handle_sk_error('LeaveMonitor()', n_code)
        self.change_state(ReceiverState.RETRY)

    def stop(self):
        """ 結束聽牌作業 """
        if self.state == ReceiverState.MONITOR:
            logger.info('stop(): 等待 EnterMonitor() 完成')
            # 如果在這個階段執行 self.skq.SKQuoteLib_LeaveMonitor(), 會觸發例外
            #   OSError: exception: access violation reading 0x0000000000000008
            # 可以看看 API 有沒有提供取消功能

        if self.state == ReceiverState.MONITOR_DONE:
            logger.debug('stop(): Leave monitor')
            n_code = self.skq.SKQuoteLib_LeaveMonitor()
            if n_code != 0:
                self.handle_sk_error('LeaveMonitor()', n_code)

        self.monitor_event.set()
        self.change_state(ReceiverState.STOP)

    def change_state(self, newState):
        """ 變更生命週期狀態 """
        if self.state is newState:
            logger.warning('state: %s not changed', newState.name)
            return
        
        if self.STATE_LOG_LEVEL == 'INFO':
            logger.info('生命週期: %s -> %s', self.state.name, newState.name)
        else:
            logger.debug('生命週期: %s -> %s', self.state.name, newState.name)

        self.state = newState

    def handle_sk_error(self, action, n_code):
        """ 顯示錯誤訊息 """
        skmsg = self.skc.SKCenterLib_GetReturnCodeMessage(n_code)
        logger.info('執行動作 [%s] 時發生錯誤, 詳細原因: #%d %s', action, n_code, skmsg)

    def OnReplyMessage(self, bstrUserID, bstrMessage):
        """ 自動答應 Login 之後的問題 """
        return 0xffff
    
    def OnConnection(self, nKind, nCode):
        """ EnterMonitor() 之後的連線事件處理 """
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
        if nKind == 3002:
            msg = '結束連線'
        if nKind == 3003:
            msg = '連線就緒'
            self.change_state(ReceiverState.MONITOR_DONE)
            self.retry_count = 0
            self.monitor_event.set()
        if nKind == 3021:
            msg = '異常斷線'
        
        logger.info('%s: nKind=%d, nCode=%d', msg, nKind, nCode)

        if nCode != 0:
            self.retry()

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """ 接收 Ticks """
        time_str = "{}".format(nTimehms)
        time_pretty = re.sub(r'(\d{2})(\d{2})(\d{2})', r'\1:\2:\3', time_str)
        logger.info("    成交: %2.2f - %s",  nClose / 100, time_pretty)

        if self.replay_count < self.REPLAY_LIMIT:
            self.replay()
            self.replay_count += 1
        else:
            self.stop()

def main():
    AsyncQuoteReceiver().start()

if __name__ == "__main__":
    main()
