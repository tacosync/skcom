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

    def __init__(self, gui_mode=False):
        super().__init__()
        self.done = False
        self.ready = False
        self.stopping = False
        self.valid_config = False
        self.gui_mode = gui_mode
        # self.stock_state = {}
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

    def set_ticks_hook(self, hook):
        self.ticks_hook = hook

    def start(self):
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
                    pn = 0
                    for stock_no in self.config['products']:
                        # 1. 這裡的回傳值是個 list, 與官方文件不符
                        #    其中 retv[0], retv[1] 都是整數值, 不確定用途, 這裡先假設 retv[1] 作為狀態碼
                        #    實際回傳: [0, 0]
                        # 2. 參數 pn 在官方文件上表示一個 pn 只能對應一檔股票, 但實測發現可以一對多,
                        #    因為這樣, 實際上可能可以突破只能聽 50 檔的限制, 不過暫時先照文件友善使用 API
                        #    pn=50 實測確實有取消 ticks 監聽作用, 所以要防止 pn=50
                        (_, nCode) = self.skQ.SKQuoteLib_RequestTicks(pn, stock_no)
                        pn += 1
                        if nCode != 0:
                            self.handleSkError('RequestTicks()', nCode)
                    # print('Ticks 請求完成')

            # 接收日 K
            if self.kline_hook is not None:
                # 日期範圍字串
                self.end_date = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')

                # 載入股票代碼/名稱對應
                for stock_no in self.config['products']:
                    pStock = sk.SKSTOCK()
                    # 這裡的回傳值是個 list, 與官方文件不符
                    # 其中 retv[0] 是 pStock 物件, retv[1] 才是整數回傳值
                    # 實際回傳: [<comtypes.gen._75AAD71C_8F4F_4F1F_9AEE_3D41A8C9BA5E_0_1_0.SKSTOCK object at 0x0000026C91615F48>, 0]
                    (_, nCode) = self.skQ.SKQuoteLib_GetStockByNo(stock_no, pStock)
                    if nCode != 0:
                        self.handleSkError('GetStockByNo()', nCode)
                        return
                    # self.stock_name[pStock.bstrStockNo] = pStock.bstrStockName
                    self.daily_kline[pStock.bstrStockNo] = {
                        'id': pStock.bstrStockNo,
                        'name': pStock.bstrStockName,
                        'quotes': []
                    }
                # print('股票名稱載入完成')

                # 請求日 K
                for stock_no in self.config['products']:
                    nCode = self.skQ.SKQuoteLib_RequestKLine(stock_no, 4, 1)
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
        if self.skQ is not None:
            self.stopping = True
            nCode = self.skQ.SKQuoteLib_LeaveMonitor()
            if nCode != 0:
                self.handleSkError('EnterMonitor', nCode)
        else:
            self.done = True

    def handleSkError(self, action, nCode):
        msg = '執行動作 [%s] 時發生錯誤, 詳細原因: %s' % (action, self.skC.SKCenterLib_GetReturnCodeMessage(nCode))
        print(msg)

    def OnConnection(self, nKind, nCode):
        if nCode != 0:
            # 這裡的 nCode 沒有對應的文字訊息
            action = '狀態變更 %d' % nKind
            self.handleSkError(action, nCode)

        # 3001 已連線
        # 3002 正常斷線
        # 3003 已就緒
        # 3021 異常斷線
        if nKind == 3003:
            self.ready = True
        if nKind == 3002 or nKind == 3021:
            self.done = True
            print('斷線')

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, nDate, nTimehms, nTimemillis, nBid, nAsk, nClose, nQty, nSimulate):
        pStock = sk.SKSTOCK()
        self.skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)
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

    ## 群益 CSV 回傳值轉換為 Python dict 型態
    def OnNotifyKLineData(self, bstrStockNo, bstrData):
        # 盤中時, 日 K 最後一筆 date 是昨天
        # 收盤時, 日 K 最後一筆 date 是??

        # 寫入緩衝區與日數限制
        cols = bstrData.split(', ')
        this_date = cols[0].replace('/', '-')
        entry = {
            'date': this_date,
            'open': float(cols[1]),
            'high': float(cols[2]),
            'low': float(cols[3]),
            'close': float(cols[4]),
            'volume': int(cols[5])
        }
        buffer = self.daily_kline[bstrStockNo]['quotes']
        buffer.append(entry)
        if self.kline_days_limit > 0 and len(buffer) > self.kline_days_limit:
            buffer.pop(0)

        # 取得最後一筆後觸發 hook, 並且清除緩衝區
        if this_date == self.end_date:
            self.kline_hook(self.daily_kline[bstrStockNo])
            self.daily_kline[bstrStockNo] = None
