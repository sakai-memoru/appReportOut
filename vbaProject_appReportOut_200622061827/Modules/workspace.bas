Attribute VB_Name = "workspace"
Option Explicit

Sub dev()

'Console.log "Hello!"

Call unittest

End Sub

Private Sub unittest()
'''' *************************************************
Console.log "-------------------- start !!"
Package.Include
''
''Call FieldDefMain.Batch("FIELD_DEF")
Call ReportDefMain.Batch("FIELD_DEF", True)
''
Package.Terminate
Console.log "-------------------- end ...."
''
End Sub

