"""
日 K 範例程式
"""

try:
    from skcom.receiver import QuoteReceiver
except ImportError as ex:
    print('尚未生成 SKCOMLib.py 請先執行一次 python -m skcom.tools.setup')
    print('例外訊息:', ex)
    exit(1)

class StockBot(QuoteReceiver):

    def __init__(self):
        super().__init__()
        self.avg_scale = [5, 10, 20, 60, 120, 240]
        self.avg_lines = [[], [], [], [], [], []]
        # self.set_kline_hook(self.on_receive_kline, 360)
        self.set_kline_hook(self.on_receive_kline, 30)

    def on_receive_kline(self, kline):
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
    #from datetime import datetime
    #print(datetime.now().strftime('%Y-%m-%d'))
    #exit(1)
    bot = StockBot()
    bot.start()

if __name__ == '__main__':
    main()
