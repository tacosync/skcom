"""
日 K 範例程式
"""
from functools import reduce
from operator import itemgetter

try:
    from skcom.receiver import QuoteReceiver
except ImportError as ex:
    print('尚未生成 SKCOMLib.py 請先執行一次 python -m skcom.tools.setup')
    print('例外訊息:', ex)
    exit(1)

class StockBot(QuoteReceiver):

    def __init__(self):
        super().__init__()
        self.steps = {}
        self.level = {}
        self.set_kline_hook(self.on_receive_kline, 240)
        self.set_ticks_hook(self.on_receive_ticks, True)

    def get_level(self, security_id, close):
        """
        檢查價位關卡
        * == -1 低於所有均線
        * >=  0 站上第 n 條均線
        """
        level = -1
        for (avg, days) in self.steps[security_id]:
            if close < avg:
                break
            level += 1
        return level

    def on_receive_ticks(self, tick):
        if self.steps:
            level = self.get_level(tick['id'], tick['close'])
            if level == self.level[tick['id']]:
                return

            # 位階發生變化, 進行通知
            if level > self.level[tick['id']]:
                lname = '%d日線' % self.steps[tick['id']][level][1]
                print('[%s] %s 現價 %.2f - 站上%s' % (tick['time'], tick['id'], tick['close'], lname))
            else:
                lname = '%d日線' % self.steps[tick['id']][level + 1][1]
                print('[%s] %s 現價 %.2f - 跌破%s' % (tick['time'], tick['id'], tick['close'], lname))

            # 記住目前位階
            self.level[tick['id']] = level
            # self.stop()

    def on_receive_kline(self, kline):
        """
        處理日 K 資料
        """
        security_id = kline['id']

        # 均線算式 (取收盤, 算總和, 算平均)
        fclose = lambda q: q['close']
        fsum   = lambda c1, c2: c1 + c2
        favg   = lambda quotes, n: reduce(fsum, map(fclose, quotes[0:n])) / n

        # 計算各條均線當日位置
        avg5   = favg(kline['quotes'], 5)
        avg10  = favg(kline['quotes'], 10)
        avg20  = favg(kline['quotes'], 20)
        avg60  = favg(kline['quotes'], 60)
        avg120 = favg(kline['quotes'], 120)
        avg240 = favg(kline['quotes'], 240)

        # 紀錄均線值, 依價位排序
        self.steps[security_id] = [
            (avg5, 5),
            (avg10, 10),
            (avg20, 20),
            (avg60, 60),
            (avg120, 120),
            (avg240, 240)
        ]
        self.steps[security_id].sort(key=itemgetter(0))

        print('均線位置:')
        for (close, days) in self.steps[security_id]:
            print('  %3d日線: %.2f' % (days, close))

        close = kline['quotes'][0]['close']
        level = self.get_level(security_id, close)
        self.level[security_id] = level
        if level == -1:
            lname = '所有均線之下'
        else:
            lname = '%d日線' % self.steps[security_id][level][1]
        print('昨收: %.2f, 位階: %s' % (close, lname))

def main():
    """
    main()
    """
    StockBot().start()

if __name__ == '__main__':
    main()
