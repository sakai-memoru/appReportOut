Attribute VB_Name = "OutReport"
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file OutReport.bas
'''  * @parent __OUT_REPORT__.xlsm
'''  *
'''  */
'

Public Function ExportAsPdf(ByVal wbName As String, ByVal shtName As String) As Variant
'''' ****************************************
Dim C_File As C_File
Set C_File = New C_File
''
Dim C_Config As C_Config
Set C_Config = New C_Config
Call C_Config.GetConfig(ThisWorkbook.name)
Dim CONF As Dictionary
Set CONF = cdset.GetValue("CONF")
''
Console.log "> ---------------------// Start!"
''
Dim aryRtn() As Variant
''
Dim C_Report As C_Report
Set C_Report = New C_Report
''
Dim wb As Workbook
Set wb = Workbooks(wbName) '' Have already opened
''
Dim base_folder As String
If CONF.Exists("BASE_FOLDER") Then
    Let base_folder = CONF.item("BASE_FOLDER")
Else
    Let base_folder = wb.Path
End If
Dim output_folder As String
If CONF.Exists("OUTPUT_FOLDER") Then
    Let output_folder = CONF.item("OUTPUT_FOLDER")
Else
    Let output_folder = ""
End If
Dim output_path
Let output_path = C_File.BuildPath(base_folder, output_folder)
''
Dim output_realpath As String
Let output_realpath = C_Report.ExportAsPdf(shtName, wb, C_File.GetAbsolutePathName(output_path))
''
Let aryRtn = Array(output_realpath)
Let ExportAsPdf = aryRtn
''
Console.log "> ---------------------// End..."
''
End Function

Public Function SaveAsHtmlTemplate(ByVal wbName As String, ByVal shtName As String) As Variant
'''' ****************************************
''
Dim C_File As C_File
Set C_File = New C_File
''
Dim C_Config As C_Config
Set C_Config = New C_Config
Call C_Config.GetConfig(ThisWorkbook.name)
Dim CONF As Dictionary
Set CONF = cdset.GetValue("CONF")
''
Console.log "> ---------------------// Start!"
''
Dim aryRtn() As Variant
''
Dim C_Report As C_Report
Set C_Report = New C_Report
''
Dim wb As Workbook
Set wb = Workbooks(wbName) '' Have already opened
''
Dim base_folder As String
If CONF.Exists("BASE_FOLDER") Then
    Let base_folder = CONF.item("BASE_FOLDER")
Else
    Let base_folder = wb.Path
End If
Dim temp_folder As String
If CONF.Exists("TEMP_FOLDER") Then
    Let temp_folder = CONF.item("TEMP_FOLDER")
Else
    Let temp_folder = ""
End If
Dim sub_folder As String
Let sub_folder = C_File.GetBaseName(wbName) & "_" & shtName & format(Now, "_YYMMDDHHmmSS")
Dim temp_path
Let temp_path = C_File.BuildPath(base_folder, temp_folder)
Let temp_path = C_File.BuildPath(temp_path, sub_folder)
Call C_File.CreateFolder(temp_path)
''
Dim outputAry() As Variant
Let outputAry = C_Report.SaveAsHtmlTemplate(shtName, wb, C_File.GetAbsolutePathName(temp_path))
''
Let aryRtn = outputAry
Let SaveAsHtmlTemplate = aryRtn
''
Console.log "> ---------------------// End..."
''
End Function

Public Function GetDef(ByVal wbName As String, ByVal shtName As String) As Variant
'''' ****************************************
''
Console.log "> ---------------------// Start!"
Dim C_Book As C_Book
Set C_Book = New C_Book
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Name As C_Name
Set C_Name = New C_Name
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim C_Report As C_Report
Set C_Report = New C_Report
Dim C_FieldType As C_FieldType
Set C_FieldType = New C_FieldType
Dim C_Template As C_Template
Set C_Template = New C_Template
''
Dim C_Config As C_Config
Set C_Config = New C_Config
Call C_Config.GetConfig(ThisWorkbook.name)
Dim CONF As Dictionary
Set CONF = cdset.GetValue("CONF")
''
Dim aryRtn() As Variant
''
Dim wb As Workbook
Set wb = Workbooks(wbName) '' Have already opened
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)
Dim nmStr As String
Let nmStr = CONF.item("TARGET_NAMEDRANGE")
''
Dim nmStatement As String
Let nmStatement = C_Name.GetNameStatement(shtName, nmStr, wb)
Console.log nmStatement
Dim rngLayout As Range
Set rngLayout = sht.Range(nmStatement)
Dim colStart As Long
Let colStart = rngLayout.column
Dim colCntOfRow As Long
Let colCntOfRow = rngLayout.Columns.Count
Dim colEnd As Long
Let colEnd = colStart + colCntOfRow - 1
Dim aryary() As Variant
Let aryary = C_Report.GetContents(nmStatement, shtName, wb)
''Console.dump aryary
Dim dict As Dictionary
Dim rng As Range
Dim tagStr As String
Dim nameStr As String
Dim borderWeightStr As String
Dim borderWeightAry As Variant
Dim aryLabel() As Variant
Dim colPosRight As Long
Dim i As Long
For i = LBound(aryary, 1) To UBound(aryary, 1)
    Set rng = wb.Worksheets(shtName).Range(aryary(i, 1))
    Set dict = C_Range.GetCellDetail(rng)
    Call dict.Add("id", dict.item("sheetName") & "!" & dict.item("address"))
    Let nameStr = C_Template.GetFieldTemplateKey(dict.item("value"))
    Call dict.Add("name", nameStr)
    If nameStr <> "" Then
        Let tagStr = "output"
        Let colPosRight = dict.item("colPos") + dict.item("colCnt") - 1
        Let aryLabel = C_Report.GetLabel(dict.item("rowPos"), dict.item("colPos"), colPosRight, shtName, wb, rngLayout)
        ''Console.dump aryLabel
        Call dict.Add("displayString", aryLabel(LBound(aryLabel)))
        Call dict.Add("for", aryLabel(UBound(aryLabel) - 1))
        Call dict.Add("lookfor", aryLabel(UBound(aryLabel)))
    Else
        Call dict.Add("displayString", dict.item("value"))
        If dict.item("font-size") > 11 Then
            If dict.item("font-weight") = "bold" Then
                Let tagStr = "h2"
            ElseIf dict.item("font-weight") = "normal" Then
                Let tagStr = "h3"
            End If
        Else
            Let borderWeightStr = dict.item("border-weight")
            Let borderWeightAry = Split(borderWeightStr, " ")
            If Not (borderWeightAry(3) = "Thin") Then '' Left Border
                Let tagStr = "div"
            Else
                If dict.item("colPos") = colStart Then
                    Let tagStr = "p"
                Else
                    Let tagStr = "p" ''FIXME
                End If
            End If
        End If
    End If
    Call dict.Add("tag", tagStr)
    Call dict.Add("_id", C_Commons.CreateGUID)
    Call C_Array.Add(aryRtn, C_Dictionary.DeepCopy(dict))
    Call dict.RemoveAll
Next i
Let GetDef = aryRtn

''
Console.log "> ---------------------// End..."
''
End Function

Public Function DumpSimpleJson(ByVal wbName As String, ByVal shtName As String) As Variant
'''' ****************************************
''
Console.log "> ---------------------// Start!"
''
Dim aryRtn() As Variant
Dim dictRtn As Dictionary
Set dictRtn = New Dictionary
''
Dim ary() As Variant
Let ary = GetDef(wbName, shtName)
Dim dict As Dictionary
Set dict = New Dictionary
Dim i As Long
For i = LBound(ary) To UBound(ary)
    Set dict = ary(i)
    If dict.item("tag") = "output" Then
        Call dictRtn.Add(dict.item("name"), dict.item("id"))
    End If
Next i
Let aryRtn = Array(dictRtn)
Let DumpSimpleJson = aryRtn
''
Console.log "> ---------------------// End..."
''
End Function

