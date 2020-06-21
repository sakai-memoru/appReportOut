VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Report.cls
'''  *<pre>Report utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/06/30
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Report
'''  * <pre> Report utility Class
'''  * </pre>
'''  *
'''  * @require C_Array,C_Range, C_File, C_FileIO
'''  */
'

Public Function SaveAsHtmlTemplate(ByVal shtName As String, Optional wb As Variant, Optional output_folder As Variant) As Variant
'' ****************************************************
''
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Report.SaveAsHtmlTemplate"
Dim C_File As C_File
Set C_File = New C_File
Dim C_Book As C_Book
Set C_Book = New C_Book
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)
''
'Dim strDtm As String
'Let strDtm = format(Now(), "_yymmddhhmmss")
''
On Error GoTo EXCEPTION
''
'get active workbook folder, if saved
If output_folder = "" Then
  output_folder = Application.DefaultFilePath
End If
''
Dim outputFile As String
Let outputFile = Replace(sht.name, " ", "")
Let outputFile = Replace(outputFile, ".", "_")
Dim outputbase As String
Dim outputPath As String
Dim outputHtmldir As String
'Let outputbase = outputFile & strDtm
Let outputbase = outputFile
Let outputPath = C_File.BuildPath(output_folder, outputbase & ".html")
Let outputHtmldir = C_File.BuildPath(output_folder, outputbase & "_files")
''
Application.DisplayAlerts = False
Call sht.Copy
''
Call ActiveWorkbook.SaveAs( _
    FileName:=outputPath, _
    FileFormat:=xlHtml _
)
ActiveWorkbook.Close
Application.DisplayAlerts = True
''
Let SaveAsHtmlTemplate = Array(C_File.BuildPath(outputHtmldir, "Sheet001.html"), C_File.BuildPath(outputHtmldir, "stylesheet.css"))
Exit Function
''
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function ExportAsPdf(ByVal shtName As String, Optional wb As Variant, Optional outputFolder As Variant) As String
'''' ****************************************
''https://www.contextures.com/excelvbapdf.html
''www.contextures.com
''for Excel 2010 and later
''
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Report.ExportAsPdf"
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)

If VBA.IsMissing(outputFolder) Then
    Let outputFolder = wb.Path
End If

''
Dim outputFile As String
Dim outputPath As String
Dim strDtm As String
Let strDtm = format(Now(), "_yymmddhhmmss")
''
On Error GoTo EXCEPTION
''
'get active workbook folder, if saved
Let outputFolder = outputFolder & "\"
''
'replace spaces and periods in sheet name
Let outputFile = Replace(sht.name, " ", "")
Let outputFile = Replace(outputFile, ".", "_")

'create default name for savng file
Let outputFile = outputFile & strDtm & ".pdf"
Let outputPath = outputFolder & outputFile

Call sht.ExportAsFixedFormat( _
        Type:=xlTypePDF, _
        FileName:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False _
    )
''
Let ExportAsPdf = outputPath
Exit Function
''
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.description, Err.HelpFile, Err.HelpContext)
    '
End Function




Public Function GetContents( _
    ByVal nmStr As String, _
    ByVal shtName As String, _
    Optional wb As Variant _
) As Variant
'''' *************************************************
Dim C_Name As C_Name
Set C_Name = New C_Name
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim nmStatement As String
If VBA.IsMissing(wb) Then
    Set wb = ThisWorkbook
    Let nmStatement = C_Name.GetNameStatement(shtName, nmStr)
Else
    Let nmStatement = C_Name.GetNameStatement(shtName, nmStr, wb)
End If
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)
''
'Console.log nmStatement
''
Dim rngLayout As Range
If Not nmStatement = "" Then
    If VBA.IsMissing(wb) Then
        Set rngLayout = C_Name.GetRange(nmStatement, shtName)
    Else
        Set rngLayout = C_Name.GetRange(nmStatement, shtName, wb)
    End If
End If
''Console.dump wb.Names
''Console.dump rngLayout

On Error Resume Next
''
Dim cellsFormulas As Range
Set cellsFormulas = rngLayout.SpecialCells(xlCellTypeFormulas)
If Err.Number <> 0 Then
    Set cellsFormulas = rngLayout.SpecialCells(xlCellTypeConstants)
    Err.Clear
End If
Dim cellsConstants As Range
Set cellsConstants = rngLayout.SpecialCells(xlCellTypeConstants)
Dim cellsFilled As Range
Set cellsFilled = Union(cellsFormulas, cellsConstants)
''Console.dump cellsFilled
Dim i As Long
Dim r As Long
Dim c As Long
Dim clt As Areas
Dim rng As Range
Dim rn As Range
Dim jagary() As Variant
If cellsFilled.Areas.Count > 0 Then
    Set clt = cellsFilled.Areas
    For i = 1 To clt.Count
        Set rng = clt.item(i)
        For r = 1 To rng.Rows.Count
            For c = 1 To rng.Columns.Count
                Set rn = rng.Cells(r, c)
                If Not VBA.IsEmpty(rn) Then
                    Call C_Array.Add(jagary, Array(rn.Address, rn.value))
                End If
            Next c
        Next r
    Next i
End If
Let jagary = C_Array.ChangeIndexOfBegin(jagary)
''
Dim addr As String
Dim val As String
Dim aryTemp() As Variant
Dim arTemp() As Variant
For i = LBound(jagary) To UBound(jagary)
    Let arTemp = C_Array.ChangeIndexOfBegin(jagary(i))
    Let addr = arTemp(LBound(arTemp))
    Set rng = sht.Range(addr)
    Let aryTemp = C_Range.ToR1C1(rng)
    Call C_Array.AddArray(arTemp, aryTemp(LBound(aryTemp)))
    Let jagary(i) = arTemp
Next i
''Console.dump jagary
''
Dim aryary() As Variant
Let aryary = C_Array.JagArrayToArrayArray(jagary)
Dim aryaryRtn() As Variant
Let aryaryRtn = C_Array.SortbyArrayArray(aryary, 3, 1, 4, 1)
''
Let GetContents = aryaryRtn
''
End Function

Public Function GetLabel(rowPos As Long, colPos As Long, colPosRight As Long, _
            ByVal shtName As String, _
            ByVal wb As Workbook, _
            Optional ByRef rngGroup As Variant) As Variant
'''' *****************************************
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Range As C_Range
Set C_Range = New C_Range
''
Dim aryRtn() As Variant
''
If IsMissing(rngGroup) Then
    Set rngGroup = wb.Worksheets(shtName).UsedRange
End If
Dim rngTargetRowLeft As Range
Set rngTargetRowLeft = wb.Worksheets(shtName).Range(Cells(rowPos, 1).Address, Cells(rowPos, colPos - 1).Address)
Dim rngTargetRowAbove As Range
Set rngTargetRowAbove = wb.Worksheets(shtName).Range(Cells(rowPos - 1, colPos).Address)
'Dim rngIntersected1 As Range
'Set rngIntersected1 = Application.Intersect(rngGroup, rngTargetRowAbove)
'Dim rngIntersected2 As Range
'Set rngIntersected2 = Application.Intersect(rngGroup, rngTargetRowLeft)
''
Dim rng As Range
Dim displayString As String
Let displayString = ""
Dim displayAddr As String
Dim displayFor As String
Dim r As Long
Dim c As Long
Dim rngTemp As Range
On Error Resume Next
    Set rng = rngTargetRowAbove
    Let displayFor = "above"
    ''Console.info "RowAbove"
    ''Console.dump C_Range.GetRangeDetail(rng)
    For r = 1 To rng.Rows.Count '' r=1 only
        For c = 1 To rng.Columns.Count Step -1
            Set rngTemp = rng.Cells(r, c)
            If Not VBA.IsEmpty(rngTemp.value) Then
                If C_Range.IsMerged(rngTemp) Then
                    Let displayString = Replace(rngTemp.MergeArea.Cells(1, 1).value, vbLf, "")
                    Let displayAddr = shtName & "!" & rngTemp.AddressLocal
                    Let aryRtn() = Array(displayString, displayAddr)
                    Let GetLabel = aryRtn()
                    Exit Function
                Else
                    If Not C_String.StartsWith(rngTemp.value, "{{") Then
                        Let displayString = rngTemp.value
                        Let displayAddr = shtName & "!" & rngTemp.AddressLocal
                        Let aryRtn() = Array(displayString, displayAddr, displayFor)
                    End If
                End If
            End If
        Next c
    Next r
    If displayString = "" Then
        Set rng = rngTargetRowLeft
        Let displayFor = "left"
        ''Console.info "RowLeft"
        ''Console.dump C_Range.GetRangeDetail(rng)
        For r = 1 To rng.Rows.Count  '' r=1 only
            For c = rng.Columns.Count To 1 Step -1
                Set rngTemp = rng.Cells(r, c)
                If Not VBA.IsEmpty(rngTemp.value) Then
                    If C_Range.IsMerged(rngTemp) Then
                        Let displayString = Replace(rngTemp.MergeArea.Cells(1, 1).value, vbLf, "")
                        Let displayAddr = shtName & "!" & rngTemp.AddressLocal
                        Let aryRtn() = Array(displayString, displayAddr)
                        Exit For
                    Else
                        If Not C_String.StartsWith(rngTemp.value, "{{") Then
                            Let displayString = rngTemp.value
                            Let displayAddr = shtName & "!" & rngTemp.AddressLocal
                            Let aryRtn() = Array(displayString, displayAddr, displayFor)
                        End If
                        Exit For
                    End If
                End If
            Next c
        Next r
        If C_Array.TypeArrayPattern(aryRtn) = "ArrayNotInit" Then
            Let aryRtn() = Array(displayString, displayAddr)
        End If
    End If
    Let GetLabel = aryRtn
    ''
End Function



