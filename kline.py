import quicksk

def on_receive_kline_entry(kline_entry):
    print(
        '[%s %s] 日期:%s 開:%.2f 收:%.2f 高:%.2f 低:%.2f 量:%d' % (
            kline_entry['id'],
            kline_entry['name'],
            kline_entry['date'],
            kline_entry['open'],
            kline_entry['close'],
            kline_entry['high'],
            kline_entry['low'],
            kline_entry['volume']
        )
    )

if __name__ == '__main__':
    qrcv = quicksk.QuoteReceiver()
    qrcv.set_kline_hook(on_receive_kline_entry)
    qrcv.start()
