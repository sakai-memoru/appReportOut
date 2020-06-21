Attribute VB_Name = "ReportOutModule"
Option Explicit

'''' **********************************************
'' @file ReportOutModule.bas
'' @parent appReportOut.xlms
''

''Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Public Sub CopyHtmlTemplate()
'''' **********************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_File As C_File
Set C_File = New C_File
Dim C_Sheet As C_Sheet
Set C_Sheet = New C_Sheet
Dim objWShell As Object
Set objWShell = CreateObject("WScript.Shell")
''
Dim dictControl As Dictionary
Set dictControl = cdset.GetValue("control_dic")
Dim dictData As Dictionary
Set dictData = cdset.GetValue("data_dic")
'Console.dump dictControl
'Console.dump dictData
Dim CONS_CMD_STATEMENT As String
Let CONS_CMD_STATEMENT = "powershell node ./js/index.js --run --tempdir ${input_folder} --tempfile ${input_file} --output ${output_folder}"
''
Dim current_folder As String
Let current_folder = ThisWorkbook.Path
Dim input_folder As String
Dim input_file As String
Dim input_path As String
Dim output_folder As String
Let output_folder = cdset.GetValue("output_folder")
''
Dim wbName As String
Dim aryShtName() As Variant
Dim shtName As String
Dim aryTemp() As Variant
Dim dict As Dictionary
Set dict = New Dictionary
Dim cmd_statement As String
Dim key As String
Dim i As Long
Dim j As Long
For i = LBound(dictControl.keys) To UBound(dictControl.keys)
    Let wbName = dictControl.keys(i)
    Let aryShtName = dictControl.item(wbName)
    For j = LBound(aryShtName) To UBound(aryShtName)
        Let shtName = aryShtName(j)
        Let key = C_Sheet.GetBookSheetExpression(wbName, shtName)
        Let aryTemp = dictData.item(key)
        Let input_path = aryTemp(LBound(aryTemp))
        Let input_file = C_File.GetFileName(input_path)
        Let input_folder = C_File.GetParentFolder(input_path)
        'Console.info input_file
        'Console.info input_folder
        'Console.info output_folder
        Call dict.Add("input_file", input_file)
        Call dict.Add("input_folder", input_folder)
        Call dict.Add("output_folder", output_folder)
        ''
        Let cmd_statement = C_String.RenderTemplate(CONS_CMD_STATEMENT, dict)
        ''Console.info cmd_statement
        ''Console.info C_File.GetCurrentDirectory
        Call C_File.ChDirectory(current_folder)
        ''Call objWShell.Run("powershell", WaitOnReturn:=True)
        Call objWShell.Run(cmd_statement, WaitOnReturn:=True)
    Next j
Next i
''
End Sub
