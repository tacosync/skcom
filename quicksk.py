import time
import signal
import threading

import pythoncom
import quicksk

qrcv = None
app_done = False

def ctrl_c(sig, frm):
    global app_done, qrcv
    print('Ctrl+C detected.')
    qrcv.finish()
    app_done = True

def main():
    global app_done, qrcv
    signal.signal(signal.SIGINT, ctrl_c)
    qrcv = quicksk.QuoteReceiver()
    # TODO: 設定報價監聽器
    qrcv.start()

    # 一定要這段才能觸發事件, 而且只有 MainThread 有效, 原因以後再查吧
    '''
    while not app_done:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)
    '''

if __name__ == '__main__':
    main()
