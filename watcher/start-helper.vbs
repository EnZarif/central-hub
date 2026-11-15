' Central Hub local helper - launcher
'
' Starts the watcher (Pixel Office + the Outlook mail bridge) with no
' console window at all.
'
' Why this exists rather than a plain shortcut to node.exe: a shortcut
' gives the helper its own console window. It sits minimised in the
' taskbar looking like something you could tidy away, and closing it kills
' the helper - at which point the Hub quietly falls back to the old
' copy-and-paste buttons with no obvious reason why. A window that must
' not be closed is a bad design; better to not have one.
'
' The "0" below is the window style (hidden) and "False" means do not wait
' for it to finish - this launcher exits immediately, leaving node running.
'
' To check it is running:  Task Manager > Details > node.exe
' To stop it:              end that node.exe task, or just log out.

Dim sh, fso, here, nodeExe, script

Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Resolve everything relative to this file, so moving the project folder
' does not silently break startup.
here    = fso.GetParentFolderName(WScript.ScriptFullName)
script  = fso.BuildPath(here, "watch.js")
nodeExe = "C:\Users\ZarifErtesam\node-v24.17.0-win-x64\node.exe"

' Fall back to whatever node is on PATH if that specific build is gone,
' which is what happens after a Node upgrade.
If Not fso.FileExists(nodeExe) Then
  nodeExe = "node.exe"
End If

If Not fso.FileExists(script) Then
  MsgBox "Central Hub helper could not start: watch.js not found at" & vbCrLf & script, _
         vbExclamation, "Central Hub"
  WScript.Quit 1
End If

sh.CurrentDirectory = here
sh.Run """" & nodeExe & """ """ & script & """", 0, False
