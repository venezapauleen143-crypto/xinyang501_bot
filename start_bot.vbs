' NiuMa Bot startup script
' Wait for network -> check duplicate -> start bot.py
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

' Wait 30s for system stability
WScript.Sleep 30000
WriteLog "Startup script started"

' Wait for network (max 120s)
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
    WriteLog "WARNING: Network not ready, still trying to start Bot"
Else
    WriteLog "Network ready"
End If

' Check if bot.py already running (prevent duplicate)
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set processes = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='pythonw.exe' OR Name='pythonw3.12.exe'")
For Each proc In processes
    cmdLine = proc.CommandLine
    If InStr(LCase(cmdLine), "bot.py") > 0 Then
        WriteLog "Bot already running (PID=" & proc.ProcessId & "), skip"
        WScript.Quit
    End If
Next

' Start Bot
WriteLog "Starting Bot..."
WshShell.CurrentDirectory = botDir
WshShell.Run """" & pythonw & """ """ & botScript & """", 0, False
WriteLog "Bot start command sent"
