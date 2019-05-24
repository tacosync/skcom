"""
日 K 範例程式
"""

import skcom

def on_receive_kline(kline):
    """
    處理日 K 資料
    """
    print('[%s %s] 的日K資料' % (kline['id'], kline['name']))
    for quote in kline['quotes']:
        print(
            '>> 日期:%s 開:%.2f 收:%.2f 高:%.2f 低:%.2f 量:%d' % (
                quote['date'],
                quote['open'],
                quote['close'],
                quote['high'],
                quote['low'],
                quote['volume']
            )
        )

def main():
    """
    main()
    """
    qrcv = skcom.QuoteReceiver()
    # 第二個參數是日數限制
    # * 0 不限制日數, 取得由史以來所有資料, 用於首次資料蒐集
    # * 預設值 20, 取得近月資料
    qrcv.set_kline_hook(on_receive_kline, 5)
    qrcv.start()

if __name__ == '__main__':
    main()
