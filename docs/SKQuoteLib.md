# 4-4 SKQuoteLib 導讀

2.13.31 解決了 nStockId > 32767 的問題, 由於參數型態改變需要改用新的方法與事件

## 方法清單

控制

名稱 | 用途
---- | ----
SKQuoteLib_EnterMonitorLONG |
SKQuoteLib_LeaveMonitor |
SKQuoteLib_IsConnected |

訂閱

名稱 | 用途
---- | ----
SKQuoteLib_RequestBoolTunnel | 訂閱布林通道
SKQuoteLib_RequestFutureTradeInfo | 訂閱期貨商品資訊
SKQuoteLib_RequestKLine | K 線查詢
SKQuoteLib_RequestKLineAM | 新 K 線查詢
SKQuoteLib_RequestKLineAMByDate | 指定區間 K 線查詢
SKQuoteLib_RequestLiveTick | 訂閱即時成交明細
SKQuoteLib_RequestMACD | 訂閱 MACD
SKQuoteLib_RequestServerTime | 主機時間查詢
SKQuoteLib_RequestStockList | 國內商品清單查詢
SKQuoteLib_RequestStocks | 訂閱報價
SKQuoteLib_RequestStocksByMarketNo | 取得行情報價
SKQuoteLib_RequestTicks | 訂閱 Ticks 及五檔
SKQuoteLib_RequestTicksWithMarketNo | 取得即時成交明細 Ticks & Best5
SKQuoteLib_CancelRequestStocks | 取消訂閱報價
SKQuoteLib_CancelRequestTicks | 取消訂閱 Ticks 及五檔

資訊取得

名稱 | 用途
---- | ----
SKQuoteLib_GetBoolTunnelLONG |
SKQuoteLib_GetMarketBuySell |
SKQuoteLib_GetMarketPriceTS |
SKQuoteLib_GetQuoteStatus |
SKQuoteLib_GetStockByIndexLONG |
SKQuoteLib_GetStockByMarketAndNO |
SKQuoteLib_GetStockBtNoLONG |
SKQuoteLib_GetStrikePrices |
SKQuoteLib_GetTickLONG |

風險參數

名稱 | 用途
---- | ----
SKQuoteLib_Delta |
SKQuoteLib_Gamma |
SKQuoteLib_Rho |
SKQuoteLib_Theta |
SKQuoteLib_Vega |

## 事件清單

控制

名稱 | 用途
---- | ----
OnConnection |

重點訂閱資訊

名稱 | 對應訂閱方法 | 用途
---- | ---- | ----
OnNotifyQuoteLONG | 不確定 | 報價更新通知
OnNotifyHistoryTicksLONG | OnNotifyHistoryTicks (即將下線) | 回補 Ticks
OnNotifyTicksLONG | SKQuoteLib_RequestTicks | 即時 Ticks
OnNotifyTicksBest5LONG | SKQuoteLib_RequestTicks | 即時五檔
OnNotifyKLineData | SKQuoteLib_RequestKLine<br>SKQuoteLib_RequestKLineAM<br>SKQuoteLib_RequestKLineAMByDate | K 線資料

其他訂閱資訊

名稱 | 用途
---- | ----
OnNotifyCommodityListData |
OnNotifyMarketBuySell |
OnNotifyMarketHighLow |
OnNotifyMarketHighLowWarrant |
OnNotifyMarketTot |
OnNotifyOddLotSpreadDeal |
OnNotifyServerTime |
OnNotifyStockListData |
OnNotifyStrikePrices |
OnNotifyTicksBoolTunnelLONG |
OnNotifyTicksFutureTradeLONG |
OnNotifyTicksMACDLONG |
