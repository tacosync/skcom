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

class MiniQuoteReceiver():

    def __init__(self, gui_mode=False):
        self.skc = None
        self.skq = None
        self.skr = None
        self.ready = False
        self.done = False
        self.stopping = False
        self.log_path = os.path.expanduser(r'~\.skcom\logs\capital')

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
        if not self.done and not self.stopping:
            print('偵測到 Ctrl+C, 結束監聽')
            self.stop()

    def start(self):
        try:
            signal.signal(signal.SIGINT, self.ctrl_c)
            
            # 載入 COM 元件
            self.skr = comtypes.client.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)
            skh0 = comtypes.client.GetEvents(self.skr, self)
            self.skc = comtypes.client.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
            self.skc.SKCenterLib_SetLogPath(self.log_path)
            self.skq = comtypes.client.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
            skh1 = comtypes.client.GetEvents(self.skq, self)
            
            # 連線
            self.login()
            self.connect()
            
            # 保持連線
            while not self.done:
                print("保持連線")
                time.sleep(1)
                pythoncom.PumpWaitingMessages() # pylint: disable=no-member
        except COMError as ex:
            self.logger.error('init() 發生不預期狀況')
            self.logger.error(ex)
        # TODO: 需要再補充其他類型的例外

    def login(self):
        n_code = self.skc.SKCenterLib_Login(self.config['account'], self.config['password'])
        if n_code != 0:
            self.handle_sk_error('Login()', n_code)
            return
        print('登入成功')

    def connect(self):
        n_code = self.skq.SKQuoteLib_EnterMonitor()
        if n_code != 0:
            self.handle_sk_error('EnterMonitor()', n_code)
            return
        print('連線成功')

        while not self.ready and not self.done:
            time.sleep(1)
            pythoncom.PumpWaitingMessages()
        print('連線就緒')

    def stop(self):
        if self.skq is not None:
            self.stopping = True
            print("Call LeaveMonitor")
            n_code = self.skq.SKQuoteLib_LeaveMonitor()
            print(n_code)
            if n_code != 0:
                self.handle_sk_error('EnterMonitor', n_code)
        else:
            print("self.skq is None")
            self.done = True

    def handle_sk_error(self, action, n_code):
        skmsg = self.skc.SKCenterLib_GetReturnCodeMessage(n_code)
        print('執行動作 [%s] 時發生錯誤, 詳細原因: %s' % (action, skmsg))

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
        if nKind == 3003:
            self.ready = True
        if nKind == 3002:
            self.done = True
            print('結束連線')
        if nKind == 3021:
            # self.done = True
            print('異常斷線')
            self.connect()

def main():
    MiniQuoteReceiver().start()

if __name__ == "__main__":
    main()