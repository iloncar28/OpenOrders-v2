On Error Resume Next 

Sub KillProcess(processName)
    Dim objShell, command
    Set objShell = CreateObject("WScript.Shell")
    command = "taskkill /F /IM " & processName
    objShell.Run command, 0, True
    Set objShell = Nothing
End Sub

KillProcess("excel.exe")
KillProcess("wscript.exe")

Dim objWMIService, objProcessList, objProcess
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objProcessList = objWMIService.ExecQuery("Select * from Win32_Process WHERE Name = 'excel.exe'")

For Each objProcess In objProcessList
    objProcess.Terminate()
Next

Set objProcessList = Nothing
Set objWMIService = Nothing
