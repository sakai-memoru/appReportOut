Attribute VB_Name = "Helper"
Option Explicit

Sub OpenCurrentFolder()
'''' **********************************************
Package.Include
''
Dim wbDir As String
Let wbDir = C_File.GetParentFolder(C_File.GetLocalName(ThisWorkbook.FullName))
'Console.log (wbDir)
Shell "cmd.exe /c start """" """ & wbDir & """", vbNormalFocus
''
Package.Terminate
End Sub

Sub OpenConfigJson()
'''' **********************************************
Package.Include
''
Dim wbDir As String
Let wbDir = C_File.GetParentFolder(C_File.GetLocalName(ThisWorkbook.FullName))
Dim configJson As String
Let configJson = C_File.BuildPath(wbDir, "config.json")
Console.log (configJson)
Shell "sakura.exe " & configJson, vbNormalFocus
''
Package.Terminate
End Sub

