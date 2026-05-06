@echo off
REM 反詐 demo 自動重啟 wrapper
REM 用法: auto_restart_wrapper.bat <stop_time>
REM 例:   auto_restart_wrapper.bat 23:59
REM
REM 行為：
REM   - python script crash 後自動重啟（無限）
REM   - 收到停止旗標檔案時退出
REM   - 每次重啟印一條 log
REM   - 重啟之間等 5 秒避免快速重啟風暴

setlocal
chcp 65001 > nul
set STOP_TIME=%1
if "%STOP_TIME%"=="" set STOP_TIME=23:59

set STOP_FILE=C:\Users\blue_\Desktop\測試檔案\.stop_line_multi
set SCRIPT_PATH=C:\Users\blue_\claude-telegram-bot\scripts\demo\反詐_multi.py
set PYTHON_EXE=C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\python3.12.exe
set LOG_DIR=C:\Users\blue_\claude-telegram-bot\scripts\demo\logs
set WRAPPER_LOG=%LOG_DIR%\wrapper.log

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

echo [%date% %time%] auto_restart_wrapper 啟動，stop_time=%STOP_TIME% >> "%WRAPPER_LOG%"

set RESTART_COUNT=0
:RESTART
REM 用戶手動 touch stop file 就退出
if exist "%STOP_FILE%" (
    echo [%date% %time%] 偵測到停止旗標，wrapper 退出 >> "%WRAPPER_LOG%"
    echo 偵測到停止旗標，退出
    goto :END
)

set /a RESTART_COUNT+=1
echo [%date% %time%] 啟動第 %RESTART_COUNT% 次... >> "%WRAPPER_LOG%"
echo ==========================================
echo 啟動第 %RESTART_COUNT% 次（%date% %time%）
echo ==========================================

set PYTHONIOENCODING=utf-8
"%PYTHON_EXE%" "%SCRIPT_PATH%" "%STOP_TIME%"
set EXIT_CODE=%errorlevel%

echo [%date% %time%] python exit code=%EXIT_CODE% >> "%WRAPPER_LOG%"

REM exit_code 0 = 正常結束（到時間或停止信號）→ 退出 wrapper
if "%EXIT_CODE%"=="0" (
    echo 正常結束（exit 0），wrapper 退出
    goto :END
)

REM exit_code != 0 = crash → 等 5 秒後重啟
echo crash（exit %EXIT_CODE%），5 秒後重啟...
timeout /t 5 /nobreak > nul
goto :RESTART

:END
echo [%date% %time%] wrapper 結束（共重啟 %RESTART_COUNT% 次） >> "%WRAPPER_LOG%"
endlocal
