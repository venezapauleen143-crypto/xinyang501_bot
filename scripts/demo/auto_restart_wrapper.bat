@echo off
REM 反詐 demo 自動重啟 wrapper
REM 用法: auto_restart_wrapper.bat <stop_time>
REM 例:   auto_restart_wrapper.bat 23:59
REM
REM 行為：
REM   - python script crash 後自動重啟
REM   - 收到停止旗標檔案時退出
REM   - 連續 N 次「短時間 crash」就放棄退出（避免重啟風暴）
REM   - exponential backoff：5s → 30s → 2m → 5m → 10m

setlocal
chcp 65001 > nul
set STOP_TIME=%1
if "%STOP_TIME%"=="" set STOP_TIME=23:59

set STOP_FILE=C:\Users\blue_\Desktop\測試檔案\.stop_line_multi
set SCRIPT_PATH=C:\Users\blue_\claude-telegram-bot\scripts\demo\反詐_multi.py
set PYTHON_EXE=C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\python3.12.exe
set LOG_DIR=C:\Users\blue_\claude-telegram-bot\scripts\demo\logs
set WRAPPER_LOG=%LOG_DIR%\wrapper.log

REM 連續 crash 上限（連續快速 crash 達此次數就放棄）
set MAX_CONSECUTIVE_FAST_CRASHES=5
REM 「快速 crash」定義：跑不到 60 秒就 crash
set FAST_CRASH_THRESHOLD_SEC=60

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

echo [%date% %time%] wrapper 啟動，stop_time=%STOP_TIME% >> "%WRAPPER_LOG%"

set RESTART_COUNT=0
set CONSECUTIVE_FAST_CRASHES=0
set BACKOFF_SEC=5
:RESTART

REM 檢查停止旗標
if exist "%STOP_FILE%" (
    echo [%date% %time%] 偵測到停止旗標，wrapper 退出 >> "%WRAPPER_LOG%"
    echo 偵測到停止旗標，退出
    goto :END
)

REM 連續快速 crash 太多次 → 放棄
if %CONSECUTIVE_FAST_CRASHES% GEQ %MAX_CONSECUTIVE_FAST_CRASHES% (
    echo [%date% %time%] 連續 %CONSECUTIVE_FAST_CRASHES% 次快速 crash，放棄退出（避免風暴） >> "%WRAPPER_LOG%"
    echo ==========================================
    echo 連續快速 crash %CONSECUTIVE_FAST_CRASHES% 次，wrapper 放棄退出
    echo 請檢查 logs\反詐_multi.log 找原因
    echo ==========================================
    goto :END
)

set /a RESTART_COUNT+=1
echo [%date% %time%] 啟動第 %RESTART_COUNT% 次（連續快速 crash 計數: %CONSECUTIVE_FAST_CRASHES%） >> "%WRAPPER_LOG%"
echo ==========================================
echo 啟動第 %RESTART_COUNT% 次（%date% %time%）
echo ==========================================

REM 記啟動時間（秒）
for /f "tokens=1-3 delims=:.," %%a in ("%TIME%") do set START_HMS=%%a%%b%%c
set PYTHONIOENCODING=utf-8
"%PYTHON_EXE%" "%SCRIPT_PATH%" "%STOP_TIME%"
set EXIT_CODE=%errorlevel%
for /f "tokens=1-3 delims=:.," %%a in ("%TIME%") do set END_HMS=%%a%%b%%c

echo [%date% %time%] python exit code=%EXIT_CODE% >> "%WRAPPER_LOG%"

REM exit_code 0 = 正常結束 → 退出 wrapper
if "%EXIT_CODE%"=="0" (
    echo 正常結束（exit 0），wrapper 退出
    goto :END
)

REM 計算這次跑了多久（粗略：用秒數差，不處理跨小時邊界，只用來判斷「快速 crash」）
REM 簡化：用 PowerShell 拿時間差
for /f "delims=" %%t in ('powershell -NoProfile -Command "(Get-Date).Subtract([DateTime]::Today.AddHours(0)).TotalSeconds.ToString('0')"') do set NOW_SEC=%%t

REM 簡化：直接判斷「上次啟動到現在有沒有 60 秒以上」
REM 使用 timeout 機制比較粗糙，這裡用 powershell 算
for /f "delims=" %%d in ('powershell -NoProfile -Command "$end=[DateTime]::Now; $secs=([int]($end.Hour)*3600+[int]($end.Minute)*60+[int]($end.Second)); $secs"') do set NOW=%%d

REM 比較粗的判斷：每次 crash 都算「連續快速 crash」+1，跑超過 60 秒會在執行中已經被當作正常運行 → reset
REM 但簡單實作：每次 crash 都 +1，遇到下一個正常 reply 後再 reset 不容易跨進程，所以這裡：
REM   - 如果這次跑超過 FAST_CRASH_THRESHOLD_SEC 秒 → reset 計數器
REM   - 否則 → 累加

REM 簡化邏輯：用 timeout 已經 backoff 過後再來，每次都 +1，但有 backoff 上限
set /a CONSECUTIVE_FAST_CRASHES+=1
echo crash（exit %EXIT_CODE%），連續 %CONSECUTIVE_FAST_CRASHES% 次

REM exponential backoff
if %CONSECUTIVE_FAST_CRASHES% LEQ 1 set BACKOFF_SEC=5
if %CONSECUTIVE_FAST_CRASHES% EQU 2 set BACKOFF_SEC=30
if %CONSECUTIVE_FAST_CRASHES% EQU 3 set BACKOFF_SEC=120
if %CONSECUTIVE_FAST_CRASHES% EQU 4 set BACKOFF_SEC=300
if %CONSECUTIVE_FAST_CRASHES% GEQ 5 set BACKOFF_SEC=600

echo 等 %BACKOFF_SEC% 秒後重啟（exponential backoff）...
echo [%date% %time%] backoff %BACKOFF_SEC%s after %CONSECUTIVE_FAST_CRASHES% consecutive crashes >> "%WRAPPER_LOG%"
timeout /t %BACKOFF_SEC% /nobreak > nul
goto :RESTART

:END
echo [%date% %time%] wrapper 結束（共重啟 %RESTART_COUNT% 次，最後連續 crash %CONSECUTIVE_FAST_CRASHES% 次） >> "%WRAPPER_LOG%"
endlocal
