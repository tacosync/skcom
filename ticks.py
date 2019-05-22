import quicksk

def on_receive_ticks_entry(ticks_entry):
    print(
        '[%s %s] 時間:%s 買:%.2f 賣:%.2f 成:%.2f 單量:%d 總量:%d' % (
            ticks_entry['id'],
            ticks_entry['name'],
            ticks_entry['time'],
            ticks_entry['bid'],
            ticks_entry['ask'],
            ticks_entry['close'],
            ticks_entry['qty'],
            ticks_entry['volume']
        )
    )

if __name__ == '__main__':
    qrcv = quicksk.QuoteReceiver()
    qrcv.set_ticks_hook(on_receive_ticks_entry)
    qrcv.start()
