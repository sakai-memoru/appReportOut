Attribute VB_Name = "DebugArea"
Option Explicit

Public Sub Kick2()
'''' ****************************************
''
Package.Include
''
Dim targetFolder As String
Let targetFolder = "G:\Users\sakai\Desktop\ExcelVbaApp\ReportOutApp\Input"
Dim targetFile As String
Let targetFile = "RequestSheet.xlsx"
Dim targetPath As String
Let targetPath = C_File.BuildPath(targetFolder, targetFile)
''
Call C_Book.OpenBook(targetPath, False, True)
Dim wb As Workbook
Set wb = Workbooks(targetFile)
Dim wbName As String
Let wbName = wb.name
Dim shtName As String
Let shtName = "Sheet"
Dim output_path As Variant
Let output_path = OutReport.SaveAsHtmlTemplate(wbName, shtName)
''Let output_path = OutReport.ExportAsPdf(wbName, shtName)
Console.dump output_path
Call C_Book.CloseBook(targetPath, False)
''
Package.Terminate
End Sub

Public Sub Kick()
'''' ****************************************
''
Package.Include
''
Dim targetFolder As String
Let targetFolder = "G:\Users\sakai\Desktop\ExcelVbaApp\ReportOutApp\Input"
Dim targetFile As String
Let targetFile = "RequestSheet.xlsx"
Dim targetPath As String
Let targetPath = C_File.BuildPath(targetFolder, targetFile)
''
Call C_Book.OpenBook(targetPath, False, True)
Dim wb As Workbook
Set wb = Workbooks(targetFile)
Dim wbName As String
Let wbName = wb.name
Dim shtName As String
Let shtName = "Sheet"
Dim ary() As Variant
Let ary = OutReport.GetDef(wbName, shtName)
'Console.dump ary
Let ary = OutReport.DumpSimpleJson(wbName, shtName)
Console.dump ary
Call C_Book.CloseBook(targetPath, False)
''
Package.Terminate
End Sub

Public Sub TestC_Range()
'''' ****************************************
Dim C_Range As C_Range
Set C_Range = New C_Range
''
Dim wb As Workbook
Set wb = ThisWorkbook
Dim shtName As String
Let shtName = "report"
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)
Dim rng As Range
'Set rng = sht.Range("C16")
'Set rng = sht.Range("C23")
'Set rng = sht.Range("J23")
'Set rng = sht.Range("AB23")
'Set rng = sht.Range("C40")
'Set rng = sht.Range("C41")
'Set rng = sht.Range("C18")
'Set rng = sht.Range("U17")
'Set rng = sht.Range("C17")
'Set rng = sht.Range("F17")
'Set rng = sht.Range("F18")
'Set rng = sht.Range("X17")
Set rng = sht.Range("F21")

Console.dump C_Range.GetCellDetail(rng)

Dim colEndPos As Long
Let colEndPos = 37
''
''Console.dump C_Range.GetRangeBlock(rng, colEndPos)
End Sub


Public Sub Process()
'''' ****************************************
''
Console.log "> ---------------------// Start!"
Package.Include
Dim C_Report As C_Report
Set C_Report = New C_Report
Dim C_FieldType As C_FieldType
Set C_FieldType = New C_FieldType
''
Dim wb As Workbook
Set wb = ThisWorkbook
Dim shtName As String
Let shtName = "report"
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)
Dim nmStr As String
Let nmStr = "layout"
Dim nmStatement As String
Let nmStatement = C_Name.GetNameStatement(shtName, nmStr)
Dim rngLayout As Range
Set rngLayout = sht.Range(nmStatement)
Dim colStart As Long
Let colStart = rngLayout.column
Dim colCntOfRow As Long
Let colCntOfRow = rngLayout.Columns.Count
Dim colEnd As Long
Let colEnd = colStart + colCntOfRow - 1
Dim aryary() As Variant
Let aryary = C_Report.GetContents(nmStr, shtName)
''Console.dump aryary
Dim dict As Dictionary
Dim rng As Range
Dim ary() As Variant
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
        Call dict.Add("for", aryLabel(UBound(aryLabel)))
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
    Call C_Array.Add(ary, C_Dictionary.DeepCopy(dict))
    Call dict.RemoveAll
Next i
Console.dump ary

''
Package.Terminate
Console.log "> ---------------------// End..."
''
End Sub

