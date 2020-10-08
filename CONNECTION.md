# 斷線測試

## 停用網卡 / 啟用網卡

**不重新登入, 重新連線**

* OnConnection() 3021
* SKQuoteLib_EnterMonitor()
* OnConnection() 3001

```
登入成功
連線成功
連線就緒
執行動作 [狀態變更 3021] 時發生錯誤, 詳細原因: 3
異常斷線
連線成功
```

**重新登入, 重新連線**

* OnConnection() 3021
* SKCenterLib_Login()
* SK_WARNING_LOGIN_ALREADY
* SKQuoteLib_EnterMonitor()
* OnConnection() 3001

```
登入成功
連線成功
連線就緒
執行動作 [狀態變更 3021] 時發生錯誤, 詳細原因: 3
異常斷線
執行動作 [Login()] 時發生錯誤, 詳細原因: SK_WARNING_LOGIN_ALREADY
連線成功
```
