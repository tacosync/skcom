import os
import json
import math
import time
import shutil
import signal
import os.path

import pythoncom
import comtypes.client
import comtypes.gen.SKCOMLib as sk

class QuoteReceiver():

    def __init__(self, gui_mode=False):
        super().__init__()
        self.done = False
        self.ready = False
        self.stopping = False
        self.valid_config = False
        self.gui_mode = gui_mode
        self.stock_state = {}
        self.log_path = os.path.expanduser('~\\.skcom\\logs')
        self.dst_conf = os.path.expanduser('~\\.skcom\\quicksk.json')
        self.tpl_conf = os.path.dirname(os.path.realpath(__file__)) + '\\conf\\quicksk.json'
        self.ticks_hook = None
        self.kline_hook = None

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

    def set_kline_hook(self, hook):
        self.kline_hook = hook

    def set_ticks_hook(self, hook):
        self.ticks_hook = hook

    def start(self):
        if self.ticks_hook is None and self.kline_hook is None:
            print('沒有設定監聽項目')
            return

        try:
            signal.signal(signal.SIGINT, self.ctrl_c)

            # 登入
            skC = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
            skC.SKCenterLib_SetLogPath(self.log_path)
            retc = skC.SKCenterLib_Login(self.config['account'], self.config['password'])
            if retc != 0: return
            print('登入成功')

            # 建立報價連線
            # 注意: comtypes.client.GetEvents() 有雷
            # * 一定要收回傳值, 即使這個回傳值沒用到, 如果不收回傳值會導致事件收不到
            # * 指定給 self.skH 會在程式結束時產生例外
            # * 指定給 global skH 會在程式結束時產生例外
            self.skQ = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
            skH = comtypes.client.GetEvents(self.skQ, self)
            retv = self.skQ.SKQuoteLib_EnterMonitor()
            if retv != 0: return
            print('連線成功')

            # 等待連線就緒
            while not self.ready and not self.done:
                time.sleep(1)
                if not self.gui_mode:
                    pythoncom.PumpWaitingMessages()

            if self.done: return
            print('連線就緒')

            # 取得產品資訊
            for stock in self.config['products']:
                self.skQ.SKQuoteLib_RequestStocks(1, stock)

            # 等待產品資訊蒐集完成
            loaded = 0
            total = len(self.config['products'])
            while loaded < total and not self.done:
                time.sleep(1)
                if not self.gui_mode:
                    pythoncom.PumpWaitingMessages()
                loaded = len(self.stock_state)

            if self.done: return
            print('產品資訊載入完成')

            # 接收訊息
            for stock in self.config['products']:
                if self.ticks_hook is not None:
                    self.skQ.SKQuoteLib_RequestTicks(1, stock)
                if self.kline_hook is not None:
                    self.skQ.SKQuoteLib_RequestKLine(stock, 4, 1)

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
        if self.skQ is not None:
            self.stopping = True
            self.skQ.SKQuoteLib_LeaveMonitor()
        else:
            self.done = True

    def OnConnection(self, nKind, nCode):
        if nKind == 3001:
            pass
        if nKind == 3003:
            self.ready = True
        if nKind == 3002 or nKind == 3021:
            print('斷線')
            self.done = True

    def OnNotifyQuote(self, sMarketNo, sStockidx):
        pStock = sk.SKSTOCK()
        self.skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)
        if pStock.bstrStockNo not in self.stock_state:
            self.stock_state[pStock.bstrStockNo] = pStock

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, nDate, nTimehms, nTimemillis, nBid, nAsk, nClose, nQty, nSimulate):
        pStock = sk.SKSTOCK()
        self.skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)
        self.stock_state[pStock.bstrStockNo] = pStock
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
            'ask': nAsk / ppow,
            'bid': nBid / ppow,
            'close': nClose / ppow,
            'qty': nQty,
            'volume': pStock.nTQty
        }
        self.ticks_hook(entry)

    def OnNotifyKLineData(self, bstrStockNo, bstrData):
        # 群益 CSV 回傳值轉換為 Python dict 型態
        pStock = self.stock_state[bstrStockNo]
        cols = bstrData.split(', ')
        entry = {
            'id': bstrStockNo,
            'name': pStock.bstrStockName,
            'date': cols[0].replace('/', '-'),
            'open': float(cols[1]),
            'high': float(cols[2]),
            'low': float(cols[3]),
            'close': float(cols[4]),
            'volume': int(cols[5])
        }
        self.kline_hook(entry)
