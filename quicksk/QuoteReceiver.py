import os
import json
import math
import time
import shutil
import signal
import asyncio
import threading
import os.path

import pythoncom
import comtypes.client
import comtypes.gen.SKCOMLib as sk

# 文件 4-1 p16
skC = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)

# 文件 4-4 p77
skQ = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)

class QuoteReceiver(threading.Thread):

    def __init__(self):
        super().__init__()
        self.done = False
        self.ready = False
        self.valid_config = False
        self.stock_state = {}
        self.log_path = os.path.expanduser('~\\.skcom\\logs')
        self.dst_conf = os.path.expanduser('~\\.skcom\\quicksk.json')
        self.tpl_conf = os.path.dirname(os.path.realpath(__file__)) + '\\conf\\quicksk.json'

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

    async def wait_for_ready(self):
        if not self.done:
            while not self.ready:
                time.sleep(0.25)

    async def monitor_quote(self):
        global skC, skQ
        try:
            #skC.SKCenterLib_ResetServer('morder1.capital.com.tw')
            skC.SKCenterLib_SetLogPath(self.log_path)
            print('登入', flush=True)
            retq = -1
            retc = skC.SKCenterLib_Login(self.config['account'], self.config['password'])
            if retc == 0:
                print('啟動行情監控', flush=True)
                retry = 0
                while retq != 0 and retry < 3 and not self.done:
                    if retry > 0:
                        print('嘗試再啟動 {}'.format(retry))
                    # TODO: 第二次 call 沒有回應, 需要改進重新連線的寫法
                    retq = skQ.SKQuoteLib_EnterMonitor()
                    retry += 1
            else:
                msg = skC.SKCenterLib_GetReturnCodeMessage(retc)
                print('登入失敗: #{} {}'.format(retc, msg))

            if retq == 0:
                print('等待行情監控器啟動完成')
                await self.wait_for_ready()
                print('設定商品', flush=True)
                for (i, stock) in enumerate(self.config['products']):
                    pn = i + 1
                    self.stock_state[stock] = None
                    skQ.SKQuoteLib_RequestStocks(pn, stock)
                    skQ.SKQuoteLib_RequestTicks(pn, stock)
                    skQ.SKQuoteLib_RequestKLine(stock, 4, 1)
            else:
                print('無法監看報價: #{}'.format(retq))

            while not self.done:
                time.sleep(1)
        except Exception as ex:
            print('init() 發生不預期狀況', flush=True)
            print(ex)
            self.done = True

    def run(self):
        ehQ = comtypes.client.GetEvents(skQ, self)
        t = threading.Thread(target=self.pumper)
        t.start()
        asyncio \
            .new_event_loop() \
            .run_until_complete(self.monitor_quote())

    def pumper(self):
        while not self.done:
            # print('fuck')
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)

    def finish(self):
        skQ.SKQuoteLib_LeaveMonitor()
        self.done = True

    def OnConnection(self, nKind, nCode):
        print('OnConnection(): nKind={}, nCode={}'.format(nKind, nCode), flush=True)
        if nKind == 3003:
            self.ready = True

    def OnNotifyQuote(self, sMarketNo, sStockidx):
        pStock = sk.SKSTOCK()
        skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)
        self.stock_state[pStock.bstrStockNo] = pStock

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, nDate, nTimehms, nTimemillis, nBid, nAsk, nClose, nQty, nSimulate):
        pStock = sk.SKSTOCK()
        skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)
        latest_state = self.stock_state[pStock.bstrStockNo]
        if latest_state is None:
            return
        ppow = math.pow(10, latest_state.sDecimal)
        msg = '{} {} 買: {} 賣: {} 成: {} 單量: {} 總量: {} 現價: {}'.format(
            pStock.bstrStockNo,
            pStock.bstrStockName,
            nAsk / ppow,
            nBid / ppow,
            nClose / ppow,
            nQty,
            latest_state.nTQty,
            latest_state.nClose / ppow
        )
        print(msg, flush=True)

    def OnNotifyKLineData(self, bstrStockNo, bstrData):
        print(bstrStockNo)
        print(bstrData)
