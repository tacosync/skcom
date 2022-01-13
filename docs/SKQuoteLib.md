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
SKQuoteLib_RequestBoolTunnel |
SKQuoteLib_RequestFutureTradeInfo |
SKQuoteLib_RequestKLine |
SKQuoteLib_RequestKLineAM |
SKQuoteLib_RequestKLineAMByDate |
SKQuoteLib_RequestLiveTick |
SKQuoteLib_RequestMACD |
SKQuoteLib_RequestServerTime |
SKQuoteLib_RequestStockList |
SKQuoteLib_RequestStocks |
SKQuoteLib_RequestStocksByMarketNo |
SKQuoteLib_RequestTicks |
SKQuoteLib_RequestTicksWithMarketNo |
SKQuoteLib_CancelRequestStocks |
SKQuoteLib_CancelRequestTicks |

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

其他

名稱 | 用途
---- | ----
SKQuoteLib_Delta |
SKQuoteLib_Gamma |
SKQuoteLib_Rho |
SKQuoteLib_Theta |
SKQuoteLib_Vega |

## 事件清單

