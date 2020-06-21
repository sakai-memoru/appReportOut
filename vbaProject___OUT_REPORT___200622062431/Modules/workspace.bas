Attribute VB_Name = "workspace"
Option Explicit

Sub dev()

'Console.log "Hello!"

'Call unittest
''Call DebugArea.Process
Call DebugArea.Kick
''
End Sub

Private Sub unittest()
'''' *************************************************
Console.info "-------------------- start !!"
Package.Include
''
Dim C_FieldType As C_FieldType
Set C_FieldType = New C_FieldType
''
Dim shtName As String
Let shtName = "User"

Dim wbFolder As String
Let wbFolder = "G:\Users\sakai\Desktop\ExcelVbaApp\FormDefApp\Input"
Dim wbName As String
Let wbName = "SSE_SD_User.xlsx"
Dim wbPath As String
Let wbPath = C_File.BuildPath(wbFolder, wbName)
''
Dim ary() As Variant
''
Console.info wbName
Call C_Book.OpenBook(wbPath)
Let ary = TransReportMain.GetFields(wbName, shtName)
Call C_Book.CloseBook(wbPath, False)

Console.dump ary
''
Package.Terminate
Console.info "-------------------- end ...."
''
End Sub
