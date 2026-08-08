Option Explicit

Dim shell, fso, appRoot, launcherPath, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

appRoot = fso.GetParentFolderName(WScript.ScriptFullName)
launcherPath = appRoot & "\Launch-HQueex-Hub.ps1"

If Not fso.FileExists(launcherPath) Then
    MsgBox "Launcher script not found: " & launcherPath, vbCritical, "H-Queex Hub"
    WScript.Quit 1
End If

command = "powershell -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & launcherPath & Chr(34)
shell.CurrentDirectory = appRoot
shell.Run command, 0, False