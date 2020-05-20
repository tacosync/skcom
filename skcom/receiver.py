"""
skcom.receiver
"""

from datetime import datetime, timedelta
import json
import logging
import math
import os
import os.path
import signal
import sys
import time

import pythoncom
from comtypes import COMError
import comtypes.client
import comtypes.gen.SKCOMLib as sk

from skcom.helper import load_config
from skcom.exception import ConfigException

class QuoteReceiver():
    """
    群益 API 報價接收器

    參考文件 v2.13.16:
      * 4-1 SKCenterLib (p.18)
      * 4-4 SKQuoteLib (p.93)
    """
    # pylint: disable=too-many-instance-attributes

    def __init__(self, gui_mode=False):
        # 狀態屬性
        self.done = False
        self.ready = False
        self.stopping = False

        # 接收器設定屬性
        self.gui_mode = gui_mode
        self.cache_path = os.path.expanduser(r'~\.skcom\cache')
        self.log_path = os.path.expanduser(r'~\.skcom\logs\capital')
        self.dst_conf = os.path.expanduser(r'~\.skcom\skcom.yaml')

        # Ticks 處理用屬性
        self.ticks_hook = None
        self.ticks_total = {}
        self.ticks_include_history = False

        # 日 K 處理用屬性
        self.kline_hook = None
        self.stock_name = {}
        self.daily_kline = {}
        self.end_date = ''
        self.kline_days_limit = 20
        self.kline_last_mtime = 0

        self.skc = None
        self.skq = None
        self.skr = None
        self.logger = logging.getLogger('skcom')

        # 產生 log 目錄
        if not os.path.isdir(self.log_path):
            os.makedirs(self.log_path)

        # 產生 cache 目錄
        if not os.path.isdir(self.cache_path):
            os.makedirs(self.cache_path)

        try:
            self.config = load_config()
        except ConfigException as ex:
            if not ex.loaded:
                print(ex)
            sys.exit(1)

    def ctrl_c(self, sig, frm):
        """
        Ctrl+C 事件觸發處理
        """
        # pylint: disable=unused-argument
        if not self.done and not self.stopping:
            self.logger.info('偵測到 Ctrl+C, 結束監聽')
            self.stop()

    def set_kline_hook(self, hook, days_limit=20):
        """
        設定日 K 回傳函數
        """
        self.kline_days_limit = days_limit
        self.kline_hook = hook

    def set_ticks_hook(self, hook, include_history=False):
        """
        設定撮合回傳函數
        """
        self.ticks_hook = hook
        self.ticks_include_history = include_history

    def start(self):
        """
        開始接收報價
        """
        # pylint: disable=too-many-branches,too-many-nested-blocks,too-many-statements,too-many-locals

        if self.ticks_hook is None and self.kline_hook is None:
            self.logger.info('沒有設定監聽項目')
            return

        try:
            signal.signal(signal.SIGINT, self.ctrl_c)

            # 登入
            self.skr = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib) # pylint: disable=invalid-name, no-member
            skh0 = comtypes.client.GetEvents(self.skr, self) # pylint: disable=unused-variable
            self.skc = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib) # pylint: disable=invalid-name, no-member
            self.skc.SKCenterLib_SetLogPath(self.log_path)
            n_code = self.skc.SKCenterLib_Login(self.config['account'], self.config['password'])
            if n_code != 0:
                # 沒插網路線會回傳 1001, 不會觸發 OnConnection
                self.handle_sk_error('Login()', n_code)
                return
            self.logger.info('登入成功')

            # 建立報價連線
            # 注意: comtypes.client.GetEvents() 有雷
            # * 一定要收回傳值, 即使這個回傳值沒用到, 如果不收回傳值會導致事件收不到
            # * 指定給 self.skH 會在程式結束時產生例外
            # * 指定給 global skH 會在程式結束時產生例外
            self.skq = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib) # pylint: disable=invalid-name, no-member
            skh1 = comtypes.client.GetEvents(self.skq, self) # pylint: disable=unused-variable
            n_code = self.skq.SKQuoteLib_EnterMonitor()
            if n_code != 0:
                # 這裡拔網路線會得到 3022, 查表沒有對應訊息
                self.handle_sk_error('EnterMonitor()', n_code)
                return
            self.logger.info('連線成功')

            # 等待連線就緒
            while not self.ready and not self.done:
                time.sleep(1)
                if not self.gui_mode:
                    pythoncom.PumpWaitingMessages() # pylint: disable=no-member

            if self.done:
                return
            self.logger.info('連線就緒')

            # 這時候拔網路線疑似會陷入無窮迴圈
            # time.sleep(3)

            # 接收 Ticks
            if self.ticks_hook is not None:
                if len(self.config['products']) > 50:
                    # 發生這個問題不阻斷使用, 讓其他功能維持正常運作
                    self.logger.warning('Ticks 最多只能監聽 50 檔')
                else:
                    for stock_no in self.config['products']:
                        # 參考文件: 4-4-6 (p.97)
                        # 1. 這裡的回傳值是個 list [pageNo, nCode], 與官方文件不符
                        # 2. 參數 psPageNo 在官方文件上表示一個 pn 只能對應一檔股票, 但實測發現可以一對多,
                        #    因為這樣, 實際上可能可以突破只能聽 50 檔的限制, 不過暫時先照文件友善使用 API
                        # 3. 參數 psPageNo 指定 -1 會自動分配 page, page 介於 0-49, 與 stock page 不同
                        # 4. 參數 psPageNo 指定 50 會取消報價
                        (page_no, n_code) = self.skq.SKQuoteLib_RequestTicks(-1, stock_no) # pylint: disable=unused-variable
                        # self.logger.info('tick page=%d' % pageNo)
                        if n_code != 0:
                            self.handle_sk_error('RequestTicks()', n_code)
                    # self.logger.info('Ticks 請求完成')

            # 接收日 K
            if self.kline_hook is not None:
                # 取樣截止日
                # 15:00 以前取樣到昨日
                # 15:00 以後取樣到當日
                now = datetime.today()
                iso_today = now.strftime('%Y-%m-%d')
                human_min = now.hour * 100 + now.minute
                day_offset = 0
                if human_min < 1500:
                    day_offset = 1

                # TODO: end_date 正確的值應該取最後交易日, 這個日期不一定是今天或昨天
                #       錯誤的日期值會導致整批回報沒觸發
                self.end_date = (datetime.today() - timedelta(days=day_offset)).strftime('%Y-%m-%d')

                # 載入股票代碼/名稱對應
                for stock_no in self.config['products']:
                    # 參考文件: 4-4-5 (p.97)
                    # 1. 參數 pSKStock 可以省略
                    # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
                    cache_filename = r'{}\{}\kline\{}.json'.format(
                        self.cache_path,
                        iso_today,
                        stock_no
                    )
                    if os.path.isfile(cache_filename):
                        with open(cache_filename, 'r') as cache_file:
                            self.daily_kline[stock_no] = json.load(cache_file)
                            self.logger.info('載入 %s 的日 K 快取', stock_no)
                        continue
                    (p_stock, n_code) = self.skq.SKQuoteLib_GetStockByNo(stock_no)
                    if n_code != 0:
                        self.handle_sk_error('GetStockByNo()', n_code)
                        return
                    self.daily_kline[p_stock.bstrStockNo] = {
                        'id': p_stock.bstrStockNo,
                        'name': p_stock.bstrStockName,
                        'quotes': []
                    }
                self.logger.info('股票名稱載入完成')

                # 請求日 K
                for stock_no in self.config['products']:
                    # 參考文件: 4-4-9 (p.99), 4-4-21 (p.105)
                    # 1. 使用方式與文件相符
                    # 2. 台股日 K 使用全盤與 AM 盤效果相同
                    cache_filename = r'{}\{}\kline\{}.json'.format(
                        self.cache_path,
                        iso_today,
                        stock_no
                    )
                    if os.path.isfile(cache_filename):
                        continue
                    self.logger.info('請求 %s 的日 K 資料', stock_no)
                    n_code = self.skq.SKQuoteLib_RequestKLine(stock_no, 4, 1)
                    # n_code = self.skq.SKQuoteLib_RequestKLineAM(stock_no, 4, 1, 1)
                    if n_code != 0:
                        self.handle_sk_error('RequestKLine()', n_code)
                self.logger.info('日 K 請求完成')

            # 命令模式下等待 Ctrl+C
            if not self.gui_mode:
                while not self.done:
                    if self.daily_kline is not None:
                        passed = time.time() - self.kline_last_mtime
                        if passed > 0.15:
                            # 生成快取目錄
                            iso_today = datetime.now().strftime('%Y-%m-%d')
                            cache_base = r'{}\{}\kline'.format(self.cache_path, iso_today)
                            if not os.path.isdir(cache_base):
                                os.makedirs(cache_base)

                            for stock_id in self.daily_kline:
                                cache_filename = r'{}\{}.json'.format(cache_base, stock_id)

                                # 寫入快取檔
                                if not os.path.isfile(cache_filename):
                                    self.daily_kline[stock_id]['quotes'].reverse()
                                    with open(cache_filename, 'w') as cache_file:
                                        # TODO: 實際存檔時發現是 BIG5 編碼, 需要看有沒有辦法改 utf-8
                                        json.dump(
                                            self.daily_kline[stock_id],
                                            cache_file,
                                            indent=2,
                                            ensure_ascii=False
                                        )

                                # 觸發事件
                                # pylint: disable=line-too-long
                                self.daily_kline[stock_id]['quotes'] = self.daily_kline[stock_id]['quotes'][0:self.kline_days_limit]
                                self.kline_hook(self.daily_kline[stock_id])

                            self.daily_kline = None
                    pythoncom.PumpWaitingMessages() # pylint: disable=no-member
                    time.sleep(0.5)

            self.logger.info('監聽結束')
        except COMError as ex:
            self.logger.error('init() 發生不預期狀況')
            self.logger.error(ex)
        # TODO: 需要再補充其他類型的例外

    def stop(self):
        """
        停止接收報價
        """
        if self.skq is not None:
            self.stopping = True
            n_code = self.skq.SKQuoteLib_LeaveMonitor()
            if n_code != 0:
                self.handle_sk_error('EnterMonitor', n_code)
        else:
            self.done = True

    def handle_sk_error(self, action, n_code):
        """
        處理群益 API 元件錯誤
        """
        # 參考文件: 4-1-3 (p.19)
        skmsg = self.skc.SKCenterLib_GetReturnCodeMessage(n_code)
        self.logger.error('執行動作 [%s] 時發生錯誤, 詳細原因: %s', action, skmsg)

    def handle_ticks(self, stock_id, name, timestr, bid, ask, close, qty, vol): # pylint: disable=too-many-arguments
        """
        處理當天回補 ticks 或即時 ticks
        """
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
        self.ticks_hook(entry)

    def OnConnection(self, nKind, nCode):
        """
        接收連線狀態變更 4-4-a (p.107)
        """
        # pylint: disable=invalid-name
        # pylint: enable=invalid-name

        if nCode != 0:
            # 這裡的 nCode 沒有對應的文字訊息
            action = '狀態變更 %d' % nKind
            self.handle_sk_error(action, nCode)

        # 參考文件: 6. 代碼定義表 (p.170)
        # 3001 已連線
        # 3002 正常斷線
        # 3003 已就緒
        # 3021 異常斷線
        if nKind == 3003:
            self.ready = True
        if nKind == 3002:
            self.done = True
            self.logger.info('結束連線')
        if nKind == 3021:
            self.done = True
            self.logger.error('異常斷線')

    def OnReplyMessage(self, bstrUserID, bstrMessage):
        """
        處理登入時的公告訊息
        """
        # pylint: disable=invalid-name, unused-argument, no-self-use
        # 檢查是否有設定自動回應已讀公告
        reply_read = False
        if 'reply_read' in self.config:
            reply_read = self.config['reply_read']

        self.logger.info('系統公告: %s', bstrMessage)
        if not reply_read:
            answer = input('回答已讀嗎? [y/n]: ')
        else:
            answer = 'y'

        if answer == 'y':
            return 0xffff
        return 0

    def OnNotifyHistoryTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """
        接收當天回補撮合 Ticks 4-4-c (p.108)
        """
        # pylint: disable=invalid-name, unused-argument, too-many-arguments
        # pylint: enable=invalid-name
        # pylint: disable=too-many-locals

        # 忽略試撮回報
        # 13:30:00 的最後一筆撮合, 即使收盤後也是透過一般 Ticks 觸發, 不會出現在回補資料中
        # [2330 台積電] 時間:13:24:59.463 買:238.00 賣:238.50 成:238.50 單量:43 總量:31348
        if nTimehms < 90000 or nTimehms >= 132500:
            return

        # 參考文件: 4-4-4 (p.96)
        # 1. pSKStock 參數可忽略
        # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
        # 3. 如果沒有 RequestStocks(), 這裡得到的總量恆為 0
        (p_stock, n_code) = self.skq.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx)
        if n_code != 0:
            self.handle_sk_error('GetStockByIndex()', n_code)
            return

        # 累加總量
        # 總量採用歷史與即時撮合累加最理想, 如果用 pStock.nTQty 會讓回補撮合的總量顯示錯誤
        if p_stock.bstrStockNo not in self.ticks_total:
            self.ticks_total[p_stock.bstrStockNo] = nQty
        else:
            self.ticks_total[p_stock.bstrStockNo] += nQty

        if self.ticks_include_history:
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
                p_stock.bstrStockName,
                timestr,
                nBid / ppow,
                nAsk / ppow,
                nClose / ppow,
                nQty,
                self.ticks_total[p_stock.bstrStockNo] # p_stock.nTQty
            )

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """
        接收即時撮合 4-4-d (p.109)
        """
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

        # 參考文件: 4-4-4 (p.96)
        # 1. pSKStock 參數可忽略
        # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
        # 3. 如果沒有 RequestStocks(), 這裡得到的總量 pStock.nTQty 恆為 0
        (p_stock, n_code) = self.skq.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx)
        if n_code != 0:
            self.handle_sk_error('GetStockByIndex()', n_code)
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
            p_stock.bstrStockName,
            timestr,
            nBid / ppow,
            nAsk / ppow,
            nClose / ppow,
            nQty,
            self.ticks_total[p_stock.bstrStockNo]
        )

    def OnNotifyKLineData(self, bstrStockNo, bstrData):
        """
        接收 K 線資料 (文件 4-4-f p.112)
        """
        # pylint: disable=invalid-name
        # pylint: enable=invalid-name

        # 新版 K 線資料格式
        # 日期        開           高          低          收          量
        # 2019/05/21, 233.500000, 236.000000, 232.500000, 234.000000, 79971
        cols = bstrData.split(', ')
        this_date = cols[0].replace('/', '-')
        self.kline_last_mtime = time.time()
        # self.logger.info('%s %.5f' % (this_date, self.kline_last_mtime))

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

            # 這部分改由事件觸發前處理
            #if self.kline_days_limit > 0 and len(buffer) > self.kline_days_limit:
            #    buffer.pop(0)

            # 取得最後一筆後觸發 hook, 並且清除緩衝區
            #if this_date == self.end_date:
            #    self.kline_hook(self.daily_kline[bstrStockNo])
            #    self.daily_kline[bstrStockNo] = None

        # 除錯用, 確認當日資料產生時機後就刪除
        # 14:00, kline 會收到昨天
        # 14:30, 待確認
        # 15:00, kline 會收到當天
        #if this_date > self.end_date:
        #    self.logger.info('當日資料已產生', this_date)
