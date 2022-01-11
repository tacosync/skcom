## 基本指令

安裝群益 API 元件, 在具有管理員權限只需要這兩條指令

* 安裝: regsvr32 'C:\Users\Taco Sync\.skcom\lib\SKCOM.dll'
* 移除: regsvr32 /u 'C:\Users\Taco Sync\.skcom\lib\SKCOM.dll'

不過在沒有管理員權限時, cmd 會執行失敗, 只能透過 PowerShell 的 -Verb RunAs 參數實現權限要求

## PowerShell 內執行 regsvr32

取得管理員權限的執行方式, 其中 -ArgumentList 最外層需要有引號, 路徑部分則需要多一組引號

這樣 -ArgumentList 的參數才能變成 String[] 型態

https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-7.2

```powershell
# 安裝
Start-Process regsvr32 -ArgumentList "`"C:\Users\Taco Sync\.skcom\lib\SKCOM.dll`"" -Verb RunAs
# 移除
Start-Process regsvr32 -ArgumentList "/u `"C:\Users\Taco Sync\.skcom\lib\SKCOM.dll`"" -Verb RunAs
```

改用 Scripting 方式生成 String[] 值, 作法如下

```powershell
# 安裝
$arglist = @("`"C:\Users\Taco Sync\.skcom\lib\SKCOM.dll`"")
Start-Process regsvr32 -ArgumentList $arglist -Verb RunAs
# 移除
$arglist = @("\u", "`"C:\Users\Taco Sync\.skcom\lib\SKCOM.dll`"")
Start-Process regsvr32 -ArgumentList $arglist -Verb RunAs
```

錯誤示範

```powershell
Start-Process regsvr32 -ArgumentList 'C:\Users\Taco Sync\.skcom\lib\SKCOM.dll' -Verb RunAs

$args = @('C:\Users\Taco Sync\.skcom\lib\SKCOM.dll')
Start-Process regsvr32 -ArgumentList $args -Verb RunAs

$args = @("C:\Users\Taco Sync\.skcom\lib\SKCOM.dll")
Start-Process regsvr32 -ArgumentList $args -Verb RunAs

$args = @("'C:\Users\Taco Sync\.skcom\lib\SKCOM.dll'")
Start-Process regsvr32 -ArgumentList $args -Verb RunAs
```

上述做法在 7.0 與 5.1 皆適用

## cmd 環境執行 PowerShell 再執行 regsvr32

使用 -encodedCommand 參數, 才能避免字串在 shell 傳遞之間被跳脫字元而發生不預期結果

```powershell
# 安裝
powershell -encodedCommand JABhAHIAZwBsAGkAcwB0ACAAPQAgAEAAKAAiAGAAIgBDADoAXABVAHMAZQByAHMAXABUAGEAYwBvACAAUwB5AG4AYwBcAC4AcwBrAGMAbwBtAFwAbABpAGIAXABTAEsAQwBPAE0ALgBkAGwAbABgACIAIgApAAoAUwB0AGEAcgB0AC0AUAByAG8AYwBlAHMAcwAgAHIAZQBnAHMAdgByADMAMgAgAC0AQQByAGcAdQBtAGUAbgB0AEwAaQBzAHQAIAAkAGEAcgBnAGwAaQBzAHQAIAAtAFYAZQByAGIAIABSAHUAbgBBAHMA
# 移除
powershell -encodedCommand JABhAHIAZwBsAGkAcwB0ACAAPQAgAEAAKAAiAC8AdQAiACwAIAAiAGAAIgBDADoAXABVAHMAZQByAHMAXABUAGEAYwBvACAAUwB5AG4AYwBcAC4AcwBrAGMAbwBtAFwAbABpAGIAXABTAEsAQwBPAE0ALgBkAGwAbABgACIAIgApAAoAUwB0AGEAcgB0AC0AUAByAG8AYwBlAHMAcwAgAHIAZQBnAHMAdgByADMAMgAgAC0AQQByAGcAdQBtAGUAbgB0AEwAaQBzAHQAIAAkAGEAcgBnAGwAaQBzAHQAIAAtAFYAZQByAGIAIABSAHUAbgBBAHMA
# 安裝
powershell -encodedCommand JABhAHIAZwBsAGkAcwB0ACAAPQAgAEAAKAAiAGAAIgBDADoAXABVAHMAZQByAHMAXABUAGEAYwBvACAAUwB5AG4AYwBcAC4AcwBrAGMAbwBtAFwAbABpAGIAXABTAEsAQwBPAE0ALgBkAGwAbABgACIAIgApAAoAUwB0AGEAcgB0AC0AUAByAG8AYwBlAHMAcwAgAHIAZQBnAHMAdgByADMAMgAgAC0AQQByAGcAdQBtAGUAbgB0AEwAaQBzAHQAIAAkAGEAcgBnAGwAaQBzAHQAIAAtAFYAZQByAGIAIABSAHUAbgBBAHMA
# 移除
powershell -encodedCommand JABhAHIAZwBsAGkAcwB0ACAAPQAgAEAAKAAiAC8AdQAiACwAIAAiAGAAIgBDADoAXABVAHMAZQByAHMAXABUAGEAYwBvACAAUwB5AG4AYwBcAC4AcwBrAGMAbwBtAFwAbABpAGIAXABTAEsAQwBPAE0ALgBkAGwAbABgACIAIgApAAoAUwB0AGEAcgB0AC0AUAByAG8AYwBlAHMAcwAgAHIAZQBnAHMAdgByADMAMgAgAC0AQQByAGcAdQBtAGUAbgB0AEwAaQBzAHQAIAAkAGEAcgBnAGwAaQBzAHQAIAAtAFYAZQByAGIAIABSAHUAbgBBAHMA
```

錯誤示範

```powershell
# 安裝
powershell -Command Start-Process regsvr32 -ArgumentList "`"C:\Users\Taco Sync\.skcom\lib\SKCOM.dll`"" -Verb RunAs
# 移除
powershell -Command Start-Process regsvr32 -ArgumentList '"/u" "C:\Users\Taco Sync\.skcom\lib\SKCOM.dll"' -Verb RunAs
```

原先的指令搬過來直接用, -ArgumentList 參數值會失去 String[] 型態而執行失敗
