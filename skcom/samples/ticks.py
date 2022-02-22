"""
撮合事件接收範例程式
"""

import asyncio

try:
    from skcom.receiver import AsyncQuoteReceiver as QuoteReceiver
except ImportError as ex:
    print('尚未生成 SKCOMLib.py 請先執行一次 python -m skcom.tools.setup')
    print('例外訊息:', ex)
    exit(1)

async def on_receive_ticks_entry(ticks_entry):
    """
    處理撮合事件
    """
    print(
        '[%s %s] 時間:%s 買:%.2f 賣:%.2f 成:%.2f 單量:%d 總量:%d' % (
            ticks_entry['id'],
            ticks_entry['name'],
            ticks_entry['time'],
            ticks_entry['bid'],
            ticks_entry['ask'],
            ticks_entry['close'],
            ticks_entry['qty'],
            ticks_entry['vol']
        )
    )

async def on_receive_best5(best5_entry):
    """
    處理最佳五檔事件
    """
    print('[%s %s] 最佳五檔' % (best5_entry['id'], best5_entry['name']))
    for i in range(0, 5):
        print('%5d %.2f | %.2f %5d' % (
            best5_entry['best'][i]['bidQty'],
            best5_entry['best'][i]['bid'],
            best5_entry['best'][i]['ask'],
            best5_entry['best'][i]['askQty'],
        ))

async def main():
    """
    main()
    """
    # 改為 False 可以關閉當日回補效果
    include_history = True
    qrcv = QuoteReceiver()
    qrcv.set_ticks_hook(on_receive_ticks_entry, include_history)
    qrcv.set_best5_hook(on_receive_best5)
    await qrcv.root_task()

if __name__ == '__main__':
    asyncio.run(main())
