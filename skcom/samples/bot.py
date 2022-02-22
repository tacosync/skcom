"""
Telegram 機器人自動通知範例程式
"""
from functools import reduce
from operator import itemgetter
import logging

try:
    from skcom.receiver import AsyncQuoteReceiver as QuoteReceiver
except ImportError as ex:
    print('尚未生成 SKCOMLib.py 請先執行一次 python -m skcom.tools.setup')
    print('例外訊息:', ex)
    exit(1)

class StockBot(QuoteReceiver):

    def __init__(self):
        super().__init__()

        # 狀態值初始化
        self.avgline_steps = {}   # 均線關卡
        self.avgline_curr = {}    # 均線目前位階
        self.volume_steps = {}    # 量能關卡
        self.volume_curr = {}     # 量能目前位階
        self.shaking_log = {}     # 均線震盪紀錄
        self.freq_threshold = {}  # 均線震盪通知頻率的時間管制值

        # 設定事件接收
        self.set_kline_hook(self.on_receive_kline, 240)
        self.set_ticks_hook(self.on_receive_ticks, True)

    def sub_minutes(self, t1, t2):
        """
        hh:mm:ss 時間值相減取分鐘數, 控制震盪通知頻率用
        """
        (hh1, mm1, ss1) = t1.split(':')
        (hh2, mm2, ss2) = t2.split(':')
        m1 = int(hh1) * 60 + int(mm1) + float(ss1) / 60
        m2 = int(hh2) * 60 + int(mm2) + float(ss2) / 60
        return m2 - m1

    def get_avgline_step(self, security_id, close):
        """
        取均線關卡
        * == -1 低於所有均線
        * >=  0 站上第 n 條均線
        """
        step = -1
        for (avg, days) in self.avgline_steps[security_id]:
            if close < avg:
                break
            step += 1
        return step

    def get_volume_step(self, security_id, volume):
        """
        取量能關卡
        * == -1 低於所有關卡
        * >=  0 站上第 n 個量能關卡
        """
        step = -1
        for (tvol, name) in self.volume_steps[security_id]:
            if volume < tvol:
                break
            step += 1
        return step

    async def on_receive_ticks(self, tick):
        if self.avgline_steps:
            logger = logging.getLogger('bot')

            security_id   = tick['id']
            security_name = tick['name']
            evt_time      = tick['time'][0:8]
            close         = tick['close']
            volume        = tick['vol']

            astep = self.get_avgline_step(security_id, close)
            if astep != self.avgline_curr[security_id]:
                # 檢查是否正在挑戰均線中
                shaking = False
                astep_vector = astep - self.avgline_curr[security_id]
                if astep_vector in [1, -1]:
                    shaking = True
                    if self.shaking_log[security_id] and self.shaking_log[security_id][-1][1] + astep_vector != 0:
                        shaking = False

                if shaking:
                    footprint = (evt_time, astep_vector)
                    self.shaking_log[security_id].append(footprint)
                else:
                    self.shaking_log[security_id].clear()
                    self.freq_threshold[security_id] = 10

                # 位階發生變化, 進行通知
                if len(self.shaking_log[security_id]) < 3:
                    # 非震盪狀況, 立即通知
                    if astep > self.avgline_curr[security_id]:
                        action = '站上'
                        days = self.avgline_steps[security_id][astep][1]
                    else:
                        action = '跌破'
                        days = self.avgline_steps[security_id][astep + 1][1]
                    logger.info('[%s] %s, %s %s 日線', security_id, security_name, action, days)
                    logger.info('... 現價 %.2f - %s', close, evt_time)
                else:
                    # 震盪狀況, 達到時間門檻才通知
                    min_passed = self.sub_minutes(self.shaking_log[security_id][0][0], self.shaking_log[security_id][-1][0])
                    if min_passed >= self.freq_threshold[security_id]:
                        if self.freq_threshold[security_id] == 10:
                            self.freq_threshold[security_id] = 30
                        else:
                            self.freq_threshold[security_id] += 30
                        if astep > self.avgline_curr[security_id]:
                            days = self.avgline_steps[security_id][astep][1]
                        else:
                            days = self.avgline_steps[security_id][astep + 1][1]
                        logger.info('[%s] %s, 在 %d 日線震盪', security_id, security_name, days)
                        logger.info('... %d 分鐘 - %s', min_passed, evt_time)

                # 記住目前位階
                self.avgline_curr[security_id] = astep

            vstep = self.get_volume_step(security_id, volume)
            if vstep != self.volume_curr[security_id]:
                self.volume_curr[security_id] = vstep
                vname  = self.volume_steps[security_id][vstep][1]
                logger.info('[%s] %s, 突破%s', security_id, security_name, vname)
                logger.info('... 總量 %d - %s', volume, evt_time)

            # self.stop()

    async def on_receive_kline(self, kline):
        """
        處理日 K 資料
        """
        security_id = kline['id']

        # 均線算式 (取收盤, 算總和, 算平均)
        fclose = lambda q: q['close']
        fsum   = lambda c1, c2: c1 + c2
        favg   = lambda quotes, n: reduce(fsum, map(fclose, quotes[0:n]), 0) / n

        # 計算各條均線當日位置
        quotes = kline['quotes']
        avg5   = favg(quotes, 5)
        avg10  = favg(quotes, 10)
        avg20  = favg(quotes, 20)
        avg60  = favg(quotes, 60)
        avg120 = favg(quotes, 120)
        avg240 = favg(quotes, 240)

        # 紀錄均線值, 依價位排序
        self.avgline_steps[security_id] = [
            (avg5, 5),
            (avg10, 10),
            (avg20, 20),
            (avg60, 60),
            (avg120, 120),
            (avg240, 240)
        ]
        self.avgline_steps[security_id].sort(key=itemgetter(0))

        # 取月均量與最大量, 作為出量參考值
        # TODO: 取 60, 20 均量 & 20 最大量, 做成量能階梯
        fvol  = lambda q: q['volume']
        fvavg = lambda quotes, n: reduce(fsum, map(fvol, quotes[0:n]), 0) / n
        fvmax = lambda quotes, n: max(map(fvol, quotes[0:n]))
        avg60 = fvavg(quotes, 60)
        avg20 = fvavg(quotes, 20)
        avgmax = fvmax(quotes, 20)
        self.volume_steps[security_id] = [
            (avg60, '季均量'),
            (avg20, '月均量'),
            (avgmax, '月最大量')
        ]
        self.volume_steps[security_id].sort(key=itemgetter(0))
        self.volume_curr[security_id] = -1
        self.shaking_log[security_id] = []
        self.freq_threshold[security_id] = 10

        close = quotes[0]['close']
        step = self.get_avgline_step(security_id, close)
        self.avgline_curr[security_id] = step
        if step == -1:
            lname = '所有均線之下'
        else:
            lname = '%d日線' % self.avgline_steps[security_id][step][1]

        logger = logging.getLogger('bot')
        logger.info('[%s] %s', kline['id'], kline['name'])
        logger.info('昨收: %.2f, 位階: %s', close, lname)

        logger.info('量能排列:')
        for (vol, name) in self.volume_steps[security_id]:
            if len(name) == 3:
                name = '\u3000' + name
            logger.info('  %s %d', name, vol)

        logger.info('均線排列:')
        for (close, days) in self.avgline_steps[security_id]:
            logger.info('  \u3000%3d日 %.2f', days, close)
        logger.info('$')

def main():
    """
    main()
    """
    logger = logging.getLogger('bot')
    StockBot().start()

if __name__ == '__main__':
    main()
