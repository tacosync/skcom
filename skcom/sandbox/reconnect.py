from datetime import datetime, timedelta
import http.client as httplib
import json
import logging
import math
import os
import os.path
import re
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
        """
        Initialize config.

        Args:
            self: (todo): write your description
            gui_mode: (str): write your description
        """
        self.skc = None
        self.skq = None
        self.skr = None
        self.ready = False
        self.resume = False
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
        """
        Print the keyboard.

        Args:
            self: (todo): write your description
            sig: (todo): write your description
            frm: (array): write your description
        """
        if not self.done and not self.stopping:
            print('偵測到 Ctrl+C, 結束監聽')
            self.stop()

    def start(self):
        """
        Starts the client.

        Args:
            self: (todo): write your description
        """
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
            self.connect()

            # 保持連線
            while not self.done:
                print('Loop start() ...')
                time.sleep(1)
                pythoncom.PumpWaitingMessages()
        except COMError as ex:
            self.logger.error('init() 發生不預期狀況')
            self.logger.error(ex)

    def connect(self):
        """
        Connects the connection.

        Args:
            self: (todo): write your description
        """
        n_code = self.skc.SKCenterLib_Login(self.config['account'], self.config['password'])
        if n_code != 0:
            self.handle_sk_error('Login()', n_code)
            return

        n_code = self.skq.SKQuoteLib_EnterMonitor()
        if n_code != 0:
            self.handle_sk_error('EnterMonitor()', n_code)
            return

        while not self.ready and not self.done:
            time.sleep(1)
            pythoncom.PumpWaitingMessages()
        
        (page_no, n_code) = self.skq.SKQuoteLib_RequestTicks(-1, "2330")
        if n_code != 0:
            self.handle_sk_error('RequestKLine()', n_code)

    def reconnect(self):
        """
        Reconnects connection

        Args:
            self: (todo): write your description
        """
        print('reconnect() A')
        n_code = self.skq.SKQuoteLib_EnterMonitor()
        if n_code != 0:
            self.handle_sk_error('EnterMonitor()', n_code)
            return

        print('reconnect() B')
        while not self.resume:
            print('Loop reconnect ...')
            time.sleep(1)
            pythoncom.PumpWaitingMessages()

        print('reconnect() C')
        (page_no, n_code) = self.skq.SKQuoteLib_RequestTicks(-1, "2330")
        if n_code != 0:
            self.handle_sk_error('RequestKLine()', n_code)
        print('reconnect() D')

    def stop(self):
        """
        Stop the code.

        Args:
            self: (todo): write your description
        """
        if self.skq is not None:
            self.stopping = True
            n_code = self.skq.SKQuoteLib_LeaveMonitor()
            if n_code != 0:
                self.handle_sk_error('LeaveMonitor()', n_code)
        else:
            self.done = True

    def handle_sk_error(self, action, n_code):
        """
        Handle an error.

        Args:
            self: (todo): write your description
            action: (str): write your description
            n_code: (int): write your description
        """
        skmsg = self.skc.SKCenterLib_GetReturnCodeMessage(n_code)
        print('執行動作 [%s] 時發生錯誤, 詳細原因: %s' % (action, skmsg))

    def OnReplyMessage(self, bstrUserID, bstrMessage):
        """
        Returns the message

        Args:
            self: (todo): write your description
            bstrUserID: (todo): write your description
            bstrMessage: (str): write your description
        """
        return 0xffff
    
    def OnConnection(self, nKind, nCode):
        """
        Called when the nCode

        Args:
            self: (todo): write your description
            nKind: (str): write your description
            nCode: (str): write your description
        """
        if nCode != 0:
            # 這裡的 nCode 沒有對應的文字訊息
            action = '狀態變更 %d' % nKind
            self.handle_sk_error(action, nCode)

        print('nKind={}, nCode={}'.format(nKind, nCode))

        # 參考文件: 6. 代碼定義表 (p.170)
        # 3001 已連線
        # 3002 正常斷線
        # 3003 已就緒
        # 3021 異常斷線
        if nKind == 3001:
            self.resume = True
            print('連線成功')
        if nKind == 3003:
            self.ready = True
            print('連線就緒')
        if nKind == 3002:
            self.done = True
            print('結束連線')
        if nKind == 3021:
            self.resume = False
            print('異常斷線')
            self.reconnect()

    def OnNotifyTicks(self, sMarketNo, sStockidx, nPtr, \
                      nDate, nTimehms, nTimemillis, \
                      nBid, nAsk, nClose, nQty, nSimulate):
        """
        Prints the ticks ] time

        Args:
            self: (todo): write your description
            sMarketNo: (str): write your description
            sStockidx: (str): write your description
            nPtr: (todo): write your description
            nDate: (todo): write your description
            nTimehms: (todo): write your description
            nTimemillis: (todo): write your description
            nBid: (str): write your description
            nAsk: (str): write your description
            nClose: (bool): write your description
            nQty: (str): write your description
            nSimulate: (str): write your description
        """
        time_str = "{}".format(nTimehms)
        time_pretty = re.sub(r'(\d{2})(\d{2})(\d{2})', r'\1:\2:\3', time_str)
        print("[{}] {:2.2f}".format(time_pretty, nClose / 100))

def wait_internet_on():
    """
    Wait until a connection to be closed.

    Args:
    """
    connected = False
    while not connected:
        conn = httplib.HTTPConnection("216.58.192.142", timeout=3)
        try:
            conn.request("HEAD", "/")
            connected = True
        except:
            time.sleep(1)
        finally:
            conn.close()

def main():
    """
    Main function.

    Args:
    """
    # MiniQuoteReceiver().start()
    wait_internet_on()
    print('Done')

if __name__ == "__main__":
    main()