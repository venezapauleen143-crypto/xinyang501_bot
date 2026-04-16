' 小牛馬 Bot 開機自啟腳本
' 功能：等網路就緒 → 檢查是否已在跑 → 啟動 bot.py
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

botDir = "C:\Users\blue_\claude-telegram-bot"
pythonw = "C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\pythonw3.12.exe"
botScript = botDir & "\bot.py"
logFile = botDir & "\startup.log"

Sub WriteLog(msg)
    Set f = fso.OpenTextFile(logFile, 8, True)
    f.WriteLine Now() & " - " & msg
    f.Close
End Sub

' 等待 30 秒讓系統穩定
WScript.Sleep 30000
WriteLog "開機自啟腳本啟動"

' 等待網路就緒（最多等 120 秒）
networkReady = False
For i = 1 To 12
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://api.telegram.org", False
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        networkReady = True
        On Error GoTo 0
        Exit For
    End If
    On Error GoTo 0
    Set http = Nothing
    WScript.Sleep 10000
Next

If Not networkReady Then
    WriteLog "警告：網路未就緒，仍嘗試啟動 Bot"
Else
    WriteLog "網路已就緒"
End If

' 檢查是否已有 bot.py 在跑（防重複啟動）
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set processes = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='pythonw.exe' OR Name='pythonw3.12.exe'")
For Each proc In processes
    cmdLine = proc.CommandLine
    If InStr(LCase(cmdLine), "bot.py") > 0 Then
        WriteLog "Bot 已在運行 (PID=" & proc.ProcessId & ")，跳過啟動"
        WScript.Quit
    End If
Next

' 啟動 Bot
WriteLog "啟動 Bot..."
WshShell.CurrentDirectory = botDir
WshShell.Run """" & pythonw & """ """ & botScript & """", 0, False
WriteLog "Bot 啟動指令已發送"
