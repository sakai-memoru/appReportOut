Attribute VB_Name = "EventHandler"
Option Explicit

Sub button1_click()
'''' ***************************************
'''' Output JSON Def
''
Call ReportOutMain.Batch("REPORT_DEF")
'Call FieldDefMain.Batch("REPORT_DEF", False, True)
End Sub

Sub Button5_Click()
'''' ***************************************
'''' Output Generate HTML/CSS Template
''
'' FIXME
Call ReportOutMain.Batch("REPORT_DEF", outTemplOn:=True)
End Sub

Sub Button8_Click()
'''' ***************************************
'''' Dump Simple JSON Def
''
Call ReportOutMain.Batch("REPORT_DEF", dumpOn:=True)
End Sub

Sub Button9_Click()
'''' ***************************************
'''' Generate Easy Template
''
Call ReportOutMain.Batch("REPORT_DEF", outTemplHtmlOn:=True)
End Sub

