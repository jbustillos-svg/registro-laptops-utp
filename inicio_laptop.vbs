Set WshShell = CreateObject("WScript.Shell")

Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

WshShell.Run """" & scriptPath & "\inicio_laptop.bat""", 0

Set WshShell = Nothing
