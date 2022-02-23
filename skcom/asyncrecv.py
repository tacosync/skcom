'''
非同步聽牌機
'''

import asyncio
import enum
import json
import logging
import math
import os
import os.path
import signal
import sys
import threading
import time
import types
from datetime import datetime, timedelta

import pythoncom
import comtypes.client
import comtypes.gen.SKCOMLib as sk
from comtypes import COMError

from skcom.helper import load_config
from skcom.exception import ConfigException

logger = logging.getLogger('skcom')

def fix_encoding(thestr):
    # TODO: 股票名稱的編碼可能會被 Python 的參數影響, 需要測試一下
    # 換了 Python 版本以後, 這樣才能取得正確中文字
    newstr = bytes(map(ord, thestr)).decode('cp950')
    # 原本這樣就可以
    # newstr = thestr
    return newstr

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
        self.FLUSH_INTERVAL = 3
        self.DELAY_LOGIN_DONE = 5 # 想在 LOGIN_DONE <-> MONITOR 之間 Ctrl+C, 延長這個時間
        self.DELAY_RETRY = 3
        self.DELAY_PUMP = 0.5
        self.STATE_LOG_LEVEL = 'DEBUG'

        # 群益 API COM 元件
        self.skc = None # 登入 API
        self.skq = None # 報價 API
        self.skr = None # 回報 API

        # 接收器設定屬性
        self.dst_conf = os.path.expanduser(r'~\.skcom\skcom.yaml')
        self.log_path = os.path.expanduser(r'~\.skcom\logs\capital')

        # 供 request() 等待連線就緒或失敗的是件
        self.monitor_event = None

        # 生命週期狀態
        self.state = ReceiverState.IDLE

        # 連線重試次數
        self.retry_count = 0

        # Ticks 處理用屬性
        self.ticks_hook = None
        self.ticks_total = {}
        self.ticks_include_history = False

        # 日 K 處理用屬性
        self.kline_hook = None
        self.stock_name = {}
        self.daily_kline = {}
        self.kline_days_limit = 20
        self.kline_last_mtime = 0

        # 最佳五檔處理用屬性
        self.best5_hook = None

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

    def start(self):
        """ 使用 asyncio 啟動非同步聽牌作業 """
        asyncio.run(self.root_task())

    def set_kline_hook(self, hook, days_limit=20):
        """ 設定日 K 回傳函數 """
        self.kline_days_limit = days_limit
        self.kline_hook = hook

    def set_ticks_hook(self, hook, include_history=False):
        """ 設定撮合回傳函數 """
        self.ticks_hook = hook
        self.ticks_include_history = include_history
    
    def set_best5_hook(self, hook):
        """ 設定最佳五檔回傳函數 """
        self.best5_hook = hook

    def ctrl_c(self, sig, frm):
        """ Ctrl+C 處理 """
        if self.state not in [ReceiverState.STOP, ReceiverState.STOP_DONE]:
            logger.info('偵測到 Ctrl+C, 結束監聽')
            self.stop()

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
        sys.stdout.flush()

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
            n_code = self.skq.SKQuoteLib_EnterMonitorLONG()
            if n_code != 0:
                self.change_state(ReceiverState.MONITOR_FAILED)
                self.monitor_event.set()
                self.handle_sk_error('EnterMonitor()', n_code)
                async def do_retry():
                    await self.retry()
                asyncio.run(do_retry())
                return
            logger.debug('connect(): SKQuoteLib_EnterMonitorLONG() done')

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

        self.request_ticks()
        self.request_kline()
        await self.handle_kline()

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
    
    def request_ticks(self):
        """ 請求 Ticks """
        if len(self.config['products']) > 50:
            # 發生這個問題不阻斷使用, 讓其他功能維持正常運作
            logger.warning('Ticks 最多只能監聽 50 檔')
        else:
            for stock_no in self.config['products']:
                # 參考文件: 4-4-6 (p.186)
                # 1. 這裡的回傳值是個 list [pageNo, nCode], 與官方文件不符
                # 2. 參數 psPageNo 在官方文件上表示一個 pn 只能對應一檔股票, 但實測發現可以一對多,
                #    因為這樣, 實際上可能可以突破只能聽 50 檔的限制, 不過暫時先照文件友善使用 API
                # 3. 參數 psPageNo 指定 -1 會自動分配 page, page 介於 0-49, 與 stock page 不同
                # 4. 參數 psPageNo 指定 50 會取消報價
                (page_no, n_code) = self.skq.SKQuoteLib_RequestTicks(-1, stock_no) # pylint: disable=unused-variable
                if n_code != 0:
                    self.handle_sk_error('RequestTicks()', n_code)
    
    def request_kline(self):
        """ 取得股名 & 請求日 K 資料 """
        # 取樣截止日
        # 生成開始日與結束日參數, 配合 SKQuoteLib_RequestKLineAMByDate 使用 YYYYMMDD 格式
        # end_date 15:00 以前取樣到昨日
        # end_date 15:00 以後取樣到當日
        # start_date 逆推 kline_days_limit 之後再逆推 14 天, 補足例假日或國定假日缺口
        now = datetime.today()
        human_min = now.hour * 100 + now.minute
        end_date_offset = 0
        if human_min < 1500:
            end_date_offset = 1
        start_date_offset = end_date_offset + self.kline_days_limit + 14
        end_date = (now - timedelta(days=end_date_offset)).strftime('%Y%m%d')
        start_date = (now - timedelta(days=start_date_offset)).strftime('%Y%m%d')
        # logger.info('request_kline() %s ~ %s', start_date, end_date)

        # 載入股票代碼/名稱對應
        for stock_no in self.config['products']:
            # 取得個股名稱
            # 參考文件: 4-4-32 (p.201)
            # 1. 參數 pSKStock 可以省略
            # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
            (p_stock, n_code) = self.skq.SKQuoteLib_GetStockByNoLONG(stock_no)
            if n_code != 0:
                if n_code == 9999:
                    logger.warning('商品 %s 資料無法取得, 請確認是否已下市', stock_no)
                else:
                    self.handle_sk_error('GetStockByNoLONG()', n_code)
                continue
            self.daily_kline[p_stock.bstrStockNo] = {
                'id': p_stock.bstrStockNo,
                'name': fix_encoding(p_stock.bstrStockName),
                'quotes': []
            }
        logger.info('股票名稱載入完成')

        # 請求日 K
        for stock_no in self.config['products']:
            # 股票資訊載入失敗的項目就跳過
            # 可以用下市產品模擬這段, 如: 00677U
            if stock_no not in self.daily_kline:
                continue

            # 參考文件: 4-4-29 (p.198)
            # 1. 使用方式與文件相符
            # 2. 台股日 K 使用全盤與 AM 盤效果相同
            logger.info('請求 %s 的日 K 資料', stock_no)
            kline_type = 4           # 0:分線 / 4:日線 / 5:週線 / 6:月線
            out_type = 1             # 0:舊版 / 1:新版
            trade_session = 1        # 0:全盤 / 1:AM盤
            min_number = 0           # 分K線的分鐘間隔 kline_type = 0 才有用
            n_code = self.skq.SKQuoteLib_RequestKLineAMByDate(
                stock_no, kline_type, out_type, trade_session,
                start_date, end_date, min_number
            )
            if n_code != 0:
                self.handle_sk_error('RequestKLine()', n_code)
                continue
        logger.info('日 K 請求完成')

    async def handle_kline(self):
        """ 整理日 K 資料, 發送事件給 hook """
        while self.state not in [ReceiverState.RETRY, ReceiverState.STOP]:
            await asyncio.sleep(0.5)

            # 最後一次收到日 K 的時間差不夠久跳過
            passed = time.time() - self.kline_last_mtime
            if passed < 0.15:
                continue

            # 生成日 K 事件
            for stock_id in self.daily_kline:
                # 觸發事件
                # pylint: disable=line-too-long
                if self.kline_hook is not None:
                    # 報價數量只留下最後 kline_days_limit 筆, 其餘捨棄
                    resp = self.daily_kline[stock_id]
                    resp['quotes'] = resp['quotes'][-self.kline_days_limit:]
                    # 觸發事件
                    retv = self.kline_hook(resp)
                    self.await_coroutine(retv)
                else:
                    logger.info('    日 K: %s 已接收', stock_id)

            # 清除緩衝資料
            # TODO: 這個做法會
            self.daily_kline = None
            break

    def handle_ticks(self, stock_id, name, timestr, bid, ask, close, qty, vol): # pylint: disable=too-many-arguments
        """ 處理當天回補 ticks 或即時 ticks """
        entry = {
            'id': stock_id,
            'name': name,
            'time': timestr,
            'bid': bid,
            'ask': ask,
            'close': close,
            'qty': qty,
            'vol': vol
        }
        if self.ticks_hook is not None:
            retv = self.ticks_hook(entry)
            self.await_coroutine(retv)
        else:
            logger.info(
                '    成交: %6s %s %.2f - %s',
                entry['id'],
                entry['name'],
                entry['close'],
                entry['time'],
            )

    def handle_sk_error(self, action, n_code):
        """ 顯示錯誤訊息 """
        skmsg = self.skc.SKCenterLib_GetReturnCodeMessage(n_code)
        logger.info('執行動作 [%s] 時發生錯誤, 詳細原因: #%d %s', action, n_code, skmsg)

    def await_coroutine(self, retv):
        """ 如果 function 回傳值是 coroutine, 放進 event loop """
        if isinstance(retv, types.CoroutineType):
            loop = asyncio.get_running_loop()
            loop.create_task(retv)

    def OnReplyMessage(self, bstrUserID, bstrMessage):
        """ 處理登入時的公告訊息 4-3-e (p.167) """
        # pylint: disable=invalid-name, unused-argument, no-self-use
        # 檢查是否有設定自動回應已讀公告
        reply_read = False
        if 'reply_read' in self.config:
            reply_read = self.config['reply_read']

        logger.info('系統公告: %s', bstrMessage)
        if not reply_read:
            answer = input('回答已讀嗎? [y/n]: ')
        else:
            answer = 'y'

        if answer == 'y':
            return 0xffff
        return 0
    
    def OnConnection(self, nKind, nCode):
        """ EnterMonitor() 之後的連線事件處理 4-4-a (p.205) """
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

    def OnNotifyTicksLONG(self, sMarketNo, nStockIndex, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """ 接收即時撮合 4-4-t (p.215) """
        # pylint: disable=invalid-name, unused-argument, too-many-arguments
        # pylint: enable=invalid-name
        # pylint: disable=too-many-locals

        # 忽略試撮回報
        # 盤中最後一筆與零股交易, 即使收盤也不會觸發歷史 Ticks, 這兩筆會在這裡觸發
        # [2330 台積電] 時間:13:24:59.463 買:238.00 賣:238.50 成:238.50 單量:43 總量:31348
        # [2330 台積電] 時間:13:30:00.000 買:238.00 賣:238.50 成:238.00 單量:3221 總量:34569
        # [2330 台積電] 時間:14:30:00.000 買:0.00 賣:0.00 成:238.00 單量:18 總量:34587
        if nTimehms < 90000 or (132500 <= nTimehms < 133000):
            return

        # 參考文件: 4-4-31 (p.200)
        # 1. pSKStock 參數可忽略
        # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
        # 3. 如果沒有 RequestStocks(), 這裡得到的總量 pStock.nTQty 恆為 0
        (p_stock, n_code) = self.skq.SKQuoteLib_GetStockByIndexLONG(sMarketNo, nStockIndex)
        if n_code != 0:
            self.handle_sk_error('GetStockByIndexLONG()', n_code)
            return

        # 累加總量
        if p_stock.bstrStockNo not in self.ticks_total:
            self.ticks_total[p_stock.bstrStockNo] = nQty
        else:
            self.ticks_total[p_stock.bstrStockNo] += nQty

        # 時間字串化
        ssdec = nTimehms % 100
        nTimehms /= 100
        mmdec = nTimehms % 100
        nTimehms /= 100
        hhdec = nTimehms
        timestr = '%02d:%02d:%02d.%03d' % (hhdec, mmdec, ssdec, nTimemillis//1000)

        # 格式轉換
        ppow = math.pow(10, p_stock.sDecimal)
        self.handle_ticks(
            p_stock.bstrStockNo,
            fix_encoding(p_stock.bstrStockName),
            timestr,
            nBid / ppow,
            nAsk / ppow,
            nClose / ppow,
            nQty,
            self.ticks_total[p_stock.bstrStockNo]
        )

    def OnNotifyHistoryTicksLONG(self, sMarketNo, nStockIndex, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """ 接收當日回補 4-4-s (p.214) """
        # pylint: disable=invalid-name, unused-argument, too-many-arguments
        # pylint: enable=invalid-name
        # pylint: disable=too-many-locals

        if self.ticks_include_history:
            self.OnNotifyTicksLONG(sMarketNo, nStockIndex, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate)

    def OnNotifyKLineData(self, bstrStockNo, bstrData):
        """ 接收 K 線資料 (文件 4-4-f p.206) """
        # pylint: disable=invalid-name
        # pylint: enable=invalid-name

        # 新版 K 線資料格式
        # 日期        開           高          低          收          量
        # 2019/05/21, 233.500000, 236.000000, 232.500000, 234.000000, 79971
        cols = bstrData.split(', ')
        this_date = cols[0].replace('/', '-')
        self.kline_last_mtime = time.time()
        # logger.info('OnNotifyKLineData() %s %.5f' % (this_date, self.kline_last_mtime))

        if self.daily_kline[bstrStockNo] is not None:
            # 寫入緩衝區與交易日數限制處理
            quote = {
                'date': this_date,
                'open': float(cols[1]),
                'high': float(cols[2]),
                'low': float(cols[3]),
                'close': float(cols[4]),
                'volume': int(cols[5])
            }
            buffer = self.daily_kline[bstrStockNo]['quotes']
            buffer.append(quote)

    def OnNotifyBest5LONG(self, sMarketNo, nStockIndex, \
            nBestBid1, nBestBidQty1, \
            nBestBid2, nBestBidQty2, \
            nBestBid3, nBestBidQty3, \
            nBestBid4, nBestBidQty4, \
            nBestBid5, nBestBidQty5, \
            nExtendBid, nExtendBidQty, \
            nBestAsk1, nBestAskQty1, \
            nBestAsk2, nBestAskQty2, \
            nBestAsk3, nBestAskQty3, \
            nBestAsk4, nBestAskQty4, \
            nBestAsk5, nBestAskQty5, \
            nExtendAsk, nExtendAskQty, \
            nSimulate):
        """ 接收最佳五檔資料 (文件 4-4-u p.216) """

        # 參考文件: 4-4-31 (p.200)
        # 1. pSKStock 參數可忽略
        # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
        # 3. 如果沒有 RequestStocks(), 這裡得到的總量 pStock.nTQty 恆為 0
        (p_stock, n_code) = self.skq.SKQuoteLib_GetStockByIndexLONG(sMarketNo, nStockIndex)
        if n_code != 0:
            self.handle_sk_error('GetStockByIndex()', n_code)
            return

        best5_entry = {
            'id': p_stock.bstrStockNo,
            'name': fix_encoding(p_stock.bstrStockName),
            'best': [
                { 'bid': nBestBid1/100, 'bidQty': nBestBidQty1, 'ask': nBestAsk1/100, 'askQty': nBestAskQty1 },
                { 'bid': nBestBid2/100, 'bidQty': nBestBidQty2, 'ask': nBestAsk2/100, 'askQty': nBestAskQty2 },
                { 'bid': nBestBid3/100, 'bidQty': nBestBidQty3, 'ask': nBestAsk3/100, 'askQty': nBestAskQty3 },
                { 'bid': nBestBid4/100, 'bidQty': nBestBidQty4, 'ask': nBestAsk4/100, 'askQty': nBestAskQty4 },
                { 'bid': nBestBid5/100, 'bidQty': nBestBidQty5, 'ask': nBestAsk5/100, 'askQty': nBestAskQty5 },
                # 這條用途不明, 暫不使用
                # { 'bid': nExtendBid, 'bidQty': nExtendBidQty, 'ask': nExtendAsk, 'askQty': nExtendAskQty },
            ]
        }
        if self.best5_hook is not None:
            retv = self.best5_hook(best5_entry)
            self.await_coroutine(retv)
