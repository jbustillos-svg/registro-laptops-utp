Set WshShell = CreateObject("WScript.Shell")

Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
logPath = fso.BuildPath(scriptPath, "arranque.log")

On Error Resume Next
Set logFile = fso.OpenTextFile(logPath, 8, True)
logFile.WriteLine Now & " VBS iniciado | ruta=" & WScript.ScriptFullName
logFile.Close
On Error GoTo 0

On Error Resume Next
WshShell.Run """" & scriptPath & "\inicio_laptop.bat""", 0, False
If Err.Number <> 0 Then
    Set logFile = fso.OpenTextFile(logPath, 8, True)
    logFile.WriteLine Now & " ERROR VBS al solicitar BAT: " & Err.Description
    logFile.Close
End If
On Error GoTo 0

Set WshShell = Nothing
