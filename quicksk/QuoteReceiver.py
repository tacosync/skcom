import os
import json
import math
import time
import shutil
import signal
import os.path
from datetime import datetime, timedelta

import pythoncom
import comtypes.client
import comtypes.gen.SKCOMLib as sk

class QuoteReceiver():
    """
    群益 API 報價接收器

    參考文件 v2.13.16:
      * 4-1 SKCenterLib (p.18)
      * 4-4 SKQuoteLib (p.93)
    """

    def __init__(self, gui_mode=False):
        super().__init__()
        self.done = False
        self.ready = False
        self.stopping = False
        self.valid_config = False
        self.gui_mode = gui_mode
        self.log_path = os.path.expanduser('~\\.skcom\\logs')
        self.dst_conf = os.path.expanduser('~\\.skcom\\quicksk.json')
        self.tpl_conf = os.path.dirname(os.path.realpath(__file__)) + '\\conf\\quicksk.json'
        self.ticks_hook = None
        self.kline_hook = None
        self.stock_name = {}
        self.daily_kline = {}

        if not os.path.isfile(self.dst_conf):
            # 產生 log 目錄
            if not os.path.isdir(self.log_path):
                os.makedirs(self.log_path)
            # 複製設定檔範本
            shutil.copy(self.tpl_conf, self.dst_conf)
        else:
            # 載入設定檔
            with open(self.dst_conf, 'r') as cfgfile:
                self.config = json.load(cfgfile)
                if self.config['account'] != 'A123456789':
                    self.valid_config = True

        if not self.valid_config:
            self.prompt()

    def prompt(self):
        # 提示
        print('請開啟設定檔，將帳號密碼改為您的證券帳號')
        print('設定檔路徑: ' + self.dst_conf)
        exit(0)

    def ctrl_c(self, sig, frm):
        if not self.done and not self.stopping:
            print('偵測到 Ctrl+C, 結束監聽')
            self.stop()

    def set_kline_hook(self, hook, days_limit=20):
        self.kline_days_limit = days_limit
        self.kline_hook = hook

    def set_ticks_hook(self, hook, include_history=False):
        self.ticks_hook = hook
        self.ticks_include_history = include_history

    def start(self):
        """
        開始接收報價
        """
        if self.ticks_hook is None and self.kline_hook is None:
            print('沒有設定監聽項目')
            return

        try:
            signal.signal(signal.SIGINT, self.ctrl_c)

            # 登入
            self.skC = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
            self.skC.SKCenterLib_SetLogPath(self.log_path)
            nCode = self.skC.SKCenterLib_Login(self.config['account'], self.config['password'])
            if nCode != 0:
                # 沒插網路線會回傳 1001, 不會觸發 OnConnection
                self.handleSkError('Login()', nCode)
                return
            print('登入成功')

            # 建立報價連線
            # 注意: comtypes.client.GetEvents() 有雷
            # * 一定要收回傳值, 即使這個回傳值沒用到, 如果不收回傳值會導致事件收不到
            # * 指定給 self.skH 會在程式結束時產生例外
            # * 指定給 global skH 會在程式結束時產生例外
            self.skQ = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
            skH = comtypes.client.GetEvents(self.skQ, self)
            nCode = self.skQ.SKQuoteLib_EnterMonitor()
            if nCode != 0:
                # 這裡拔網路線會得到 3022, 查表沒有對應訊息
                self.handleSkError('EnterMonitor()', nCode)
                return
            print('連線成功')

            # 等待連線就緒
            while not self.ready and not self.done:
                time.sleep(1)
                if not self.gui_mode:
                    pythoncom.PumpWaitingMessages()

            if self.done: return
            print('連線就緒')

            # 這時候拔網路線疑似會陷入無窮迴圈
            # time.sleep(3)

            # 接收 Ticks
            if self.ticks_hook is not None:
                if len(self.config['products']) > 50:
                    # 發生這個問題不阻斷使用, 讓其他功能維持正常運作
                    print('Ticks 最多只能監聽 50 檔')
                else:
                    # 參考文件: 4-4-3 (p.95)
                    # 1. 這裡的回傳值是個 list [pageNo, nCode], 與官方文件不符
                    # 2. 官方文件說 pn 值應介於 1-49, 每個 page 最多 100 檔
                    # 3. 參數 psPageNo 指定 -1 會自動分配 page
                    # 4. 參數 psPageNo 指定 50 會取消報價
                    # 5. 如果沒請求, 總量值在 Ticks 事件觸發時也無法取得
                    stock_list = ','.join(self.config['products'])
                    (pageNo, nCode) = self.skQ.SKQuoteLib_RequestStocks(-1, stock_list)
                    # print('stock page=%d' % pageNo)
                    if nCode != 0:
                        self.handleSkError('RequestStocks()', nCode)

                    for stock_no in self.config['products']:
                        # 參考文件: 4-4-6 (p.97)
                        # 1. 這裡的回傳值是個 list [pageNo, nCode], 與官方文件不符
                        # 2. 參數 psPageNo 在官方文件上表示一個 pn 只能對應一檔股票, 但實測發現可以一對多,
                        #    因為這樣, 實際上可能可以突破只能聽 50 檔的限制, 不過暫時先照文件友善使用 API
                        # 3. 參數 psPageNo 指定 -1 會自動分配 page, page 介於 0-49, 與 stock page 不同
                        # 4. 參數 psPageNo 指定 50 會取消報價
                        # 5. 如果只請求 ticks 而不請求 stocks, 總量會恆為 0
                        # 6. Ticks page 與 Stock page 各為獨立的 page, 彼此沒有關係
                        (pageNo, nCode) = self.skQ.SKQuoteLib_RequestTicks(-1, stock_no)
                        # print('tick page=%d' % pageNo)
                        if nCode != 0:
                            self.handleSkError('RequestTicks()', nCode)
                    # print('Ticks 請求完成')

            # 接收日 K
            if self.kline_hook is not None:
                # 日期範圍字串
                self.end_date = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')

                # 載入股票代碼/名稱對應
                for stock_no in self.config['products']:
                    # 參考文件: 4-4-5 (p.97)
                    # 1. 參數 pSKStock 可以省略
                    # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
                    (pStock, nCode) = self.skQ.SKQuoteLib_GetStockByNo(stock_no)
                    if nCode != 0:
                        self.handleSkError('GetStockByNo()', nCode)
                        return
                    self.daily_kline[pStock.bstrStockNo] = {
                        'id': pStock.bstrStockNo,
                        'name': pStock.bstrStockName,
                        'quotes': []
                    }
                # print('股票名稱載入完成')

                # 請求日 K
                for stock_no in self.config['products']:
                    # 參考文件: 4-4-9 (p.99), 4-4-21 (p.105)
                    # 1. 使用方式與文件相符
                    # 2. 台股日 K 使用全盤與 AM 盤效果相同
                    nCode = self.skQ.SKQuoteLib_RequestKLine(stock_no, 4, 1)
                    # nCode = self.skQ.SKQuoteLib_RequestKLineAM(stock_no, 4, 1, 1)
                    if nCode != 0:
                        self.handleSkError('RequestKLine()', nCode)
                # print('日 K 請求完成')

            # 命令模式下等待 Ctrl+C
            if not self.gui_mode:
                while not self.done:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.5)

            print('監聽結束')
        except Exception as ex:
            print('init() 發生不預期狀況', flush=True)
            print(ex)

    def stop(self):
        """
        停止接收報價
        """
        if self.skQ is not None:
            self.stopping = True
            nCode = self.skQ.SKQuoteLib_LeaveMonitor()
            if nCode != 0:
                self.handleSkError('EnterMonitor', nCode)
        else:
            self.done = True

    def handleSkError(self, action, nCode):
        """
        處理群益 API 元件錯誤
        """
        # 參考文件: 4-1-3 (p.19)
        skmsg = self.skC.SKCenterLib_GetReturnCodeMessage(nCode)
        msg = '執行動作 [%s] 時發生錯誤, 詳細原因: %s' % (action, skmsg)
        print(msg)

    def OnConnection(self, nKind, nCode):
        """
        接收連線狀態變更 4-4-a (p.107)
        """
        if nCode != 0:
            # 這裡的 nCode 沒有對應的文字訊息
            action = '狀態變更 %d' % nKind
            self.handleSkError(action, nCode)

        # 參考文件: 6. 代碼定義表 (p.170)
        # 3001 已連線
        # 3002 正常斷線
        # 3003 已就緒
        # 3021 異常斷線
        if nKind == 3003:
            self.ready = True
        if nKind == 3002 or nKind == 3021:
            self.done = True
            print('斷線')

    def OnNotifyQuote(self, sMarketNo, sStockidx):
        """
        接收報價更新 4-4-b (p.108)
        僅用來確保總量值正常, 不處理這個事件
        """
        pass

    def OnNotifyHistoryTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """
        接收當天回補的 Ticks 4-4-c (p.108)
        """
        if self.ticks_include_history:
            self.OnNotifyTicks(
                sMarketNo, sStockidx, nPtr,
                nDate, nTimehms, nTimemillis,
                nBid, nAsk, nClose, nQty, nSimulate
            )

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """
        接收 Ticks 4-4-d (p.109)
        """
        # 參考文件: 4-4-4 (p.96)
        # 1. pSKStock 參數可忽略
        # 2. 回傳值是 list [SKSTOCKS*, nCode], 與官方文件不符
        # 3. 如果沒有 RequestStocks(), 這裡得到的總量恆為 0
        (pStock, nCode) = self.skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx)
        ppow = math.pow(10, pStock.sDecimal)

        # 14:30:00 的紀錄不處理
        if nTimehms > 133000:
            return

        # 時間字串化
        s = nTimehms % 100
        nTimehms /= 100
        m = nTimehms % 100
        nTimehms /= 100
        h = nTimehms
        timestr = '%02d:%02d:%02d.%03d' % (h, m, s, nTimemillis//1000)

        entry = {
            'id': pStock.bstrStockNo,
            'name': pStock.bstrStockName,
            'time': timestr,
            'bid': nBid / ppow,
            'ask': nAsk / ppow,
            'close': nClose / ppow,
            'qty': nQty,
            'volume': pStock.nTQty
        }
        self.ticks_hook(entry)

    def OnNotifyKLineData(self, bstrStockNo, bstrData):
        """
        接收 K 線資料 (文件 4-4-f p.112)
        """
        # 盤中時, 日 K 最後一筆 date 是昨天
        # 收盤時, 日 K 最後一筆 date 還是昨天
        # 14:00, 待確認
        # 15:00, 待確認
        # 16:00, 待確認
        # 17:00, 待確認

        # 寫入緩衝區與日數限制
        cols = bstrData.split(', ')
        this_date = cols[0].replace('/', '-')
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
        if self.kline_days_limit > 0 and len(buffer) > self.kline_days_limit:
            buffer.pop(0)

        # 取得最後一筆後觸發 hook, 並且清除緩衝區
        if this_date == self.end_date:
            self.kline_hook(self.daily_kline[bstrStockNo])
            self.daily_kline[bstrStockNo] = None
        else:
            #if this_date > self.end_date:
            #    print(this_date)
            pass
