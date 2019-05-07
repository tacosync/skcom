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
    # TODO: 設定 K 線監聽器
    qrcv.start()

if __name__ == '__main__':
    main()
