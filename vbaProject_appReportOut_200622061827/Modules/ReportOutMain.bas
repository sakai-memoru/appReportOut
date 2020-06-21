Attribute VB_Name = "ReportOutMain"
Option Explicit

'''' **********************************************
'' @file ReportOutMain.bas
'' @parent appReportOut.xlms
''

Public Function Batch( _
        ByVal datatype As String, _
        Optional ByVal outTemplOn As Variant = False, _
        Optional ByVal outTemplHtmlOn As Variant = False, _
        Optional ByVal dumpOn As Variant = False, _
        Optional ByVal moveOn As Variant = False _
    ) As Variant
'''' **********************************************
'' @function batch
'' @param datatype {String} 処理データタイプ
'' @param outTemplOn {Variant<boolean>}
''            Template出力flag
'' @param dumpOn {Variant<boolean>}
''            Check用Dump出力flag
'' @param moveOn  {Variant<boolean>}
''            Inputファイル移動flag
''
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "ReportOutMain.Batch"
''
Package.Include
' ## C_Config.GetConfig
Dim C_Config As C_Config
Set C_Config = New C_Config
Call C_Config.GetConfig
Dim CONF As Dictionary
Set CONF = cdset.GetValue("CONF")
''
On Error GoTo EXCEPTION
    Console.info "-------------------- start !!"
    ' ## CommonModule.SetCommonDataSet
    Call CommonModule.SetCommonDataSet
    ''Console.dump CONF
    ''
    ' ## CommonModule.SetEnvVariables
    Call CommonModule.SetEnvVariables
    ''Console.dump cdset
    ''
    ' - Set Variables
    Dim input_folder As String
    Let input_folder = cdset.GetValue("input_folder")
    Dim backup_folder As String
    Let backup_folder = cdset.GetValue("backup_folder")
    Dim form_folder As String
    Let form_folder = cdset.GetValue("form_folder")
    ''
    Dim sheet_type As String
    Let sheet_type = CONF.item(datatype)("SHEET_TYPE")
    Dim input_like As String
    Let input_like = CONF.item(datatype)("INPUT_LIKE")
    ''
    Dim form_fileName As String
    Dim form_path As String
    Dim isFormAdapted As Boolean
    If CONF.item(datatype).Exists("FORM_FILE") Then
        Let isFormAdapted = True
        Let form_fileName = CONF.item(datatype)("FORM_FILE")
        Let form_path = C_File.BuildPath(form_folder, form_fileName)
    Else
        Let isFormAdapted = False
    End If
    ''
    Dim macroName As String
    Let macroName = CONF.item(datatype)("MACRO_GET_METHOD")
    If outTemplOn Then
        Let macroName = CONF.item(datatype)("MACRO_OUT_METHOD")
    End If
    If outTemplHtmlOn Then
        Let macroName = CONF.item(datatype)("MACRO_OUTHTML_METHOD")
    End If
    If dumpOn Then
        Let macroName = CONF.item(datatype)("MACRO_DUMP_METHOD")
    End If
    ''
    ' ## C_Book.GetXlsxes
    Dim aryWbFullName() As Variant
    Let aryWbFullName = C_Book.GetXlsxes(input_folder, input_like)
    'Console.dump aryWbFullName
    Dim i As Long
    ''
    Application.ScreenUpdating = False
    If isFormAdapted Then
        ' ## C_Book.OpenBook
        Call C_Book.OpenBook(form_path, False, False)
        'C_Book.OpenBook (form_path)
        ''Dim wbForm As Workbook
        ''Set wbForm = Workbooks(form_fileName) ''FIXME
        ' (())-Loop Each books
        For i = LBound(aryWbFullName) To UBound(aryWbFullName)
            Call ProcessForReportSheet(aryWbFullName(i), datatype, macroName)
        Next i
        ' ## C_Book.CloseBook
        C_Book.CloseBook filePath:=form_path, saveChanges:=False
    End If
    Application.ScreenUpdating = True
    ''
    ''Console.dump cdset.GetValue("control_dic")
    ''Console.dump cdset.GetValue("data_dic")
    ''
    If Not (outTemplOn Or outTemplHtmlOn) Then
        ' ## CommonModule.OutputJson
        Call CommonModule.OutputJson(macroName)
    ElseIf outTemplHtmlOn Then
        Call ReportOutModule.CopyHtmlTemplate
    Else
        ' ## FieldDefModule.OutputTemplate
        'Call ReportOutModule.OutputTemplate
    End If
    ''
    If moveOn Then
        For i = LBound(aryWbFullName) To UBound(aryWbFullName)
            ' ## C_File.MoveFile
            Call C_File.MoveFile(aryWbFullName(i), backup_folder & "/")
        Next i
    End If
    ''
    GoSub FINALLY
    Let Batch = 0 '' return
    Exit Function
    ''
FINALLY: 'Label
    Package.Terminate
    Console.info "-------------------- end ...."
    Return
    ''
EXCEPTION: 'Label
    GoSub FINALLY
    Console.error ("!!!!!! An error has occured !!")
    Console.error ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.error ("Err.Number = " & Err.Number)
    Console.error ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.description)
    Let Batch = -1 '' return
    ''
End Function

Public Sub ProcessForReportSheet( _
      ByVal wbPath As String, _
      ByVal datatype As String, _
      ByVal macroName As String _
    )
'''' **********************************************
'' @function ProcessForReportSheet
'' @param wbPath {String} book path
'' @param datatype {String} 処理データタイプ
'' @param macroName {String}
''
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "ReportOutMain.ProcessForReportSheet"
''
Dim CONF As Dictionary
Set CONF = cdset.GetValue("CONF")
Console.dump CONF
''
On Error GoTo EXCEPTION
    ' - Set Common Dataset
    Dim dicControl As Dictionary
    Set dicControl = cdset.GetValue("control_dic")
    Dim dicData As Dictionary
    Set dicData = cdset.GetValue("data_dic")
    ''
    Dim wbNameForm As String
    Dim shtNameFormSheet As String
    Dim sheetType As String
    If CONF.item(datatype).Exists("FORM_FILE") Then
        Let wbNameForm = CONF.item(datatype)("FORM_FILE")
        Let shtNameFormSheet = CONF.item(datatype)("FORM_SHEET")
        Let sheetType = CONF.item(datatype)("SHEET_TYPE")
    End If
    ''
    Dim wbName As String
    Let wbName = C_File.GetFileName(wbPath)
    ' - Start Each Book Process
    Console.info ""
    Console.info "..... For """ & wbName & """ !!"
    ' ## C_Book.OpenBook
    Call C_Book.OpenBook(filePath:=wbPath, updateLinks:=False, readOnly:=True)
    Dim wb As Workbook
    Set wb = Workbooks(wbName)
    Dim aryShtName As Variant
    Let aryShtName = C_Book.GetSheetsByPrefix(wbName)
    ''
    Dim data_records() As Variant
    Dim detail_records() As Variant
    ' ## C_Book.GetBookProperties
    Dim wbUpdateDate As Date
    Let wbUpdateDate = C_Book.GetBookProperties(wbName).item("Last save time")
    Dim createData As Date
    Let createData = Now()
    Dim keyWbSht As String
    Dim guid As String
    ' (())-Loop Each Sheet
    Dim r As Long
    Dim i As Long
    For i = LBound(aryShtName) To UBound(aryShtName)
        Console.info ".....   : '" & wbName & "'!" & aryShtName(i) & "....."
        Let keyWbSht = C_Sheet.GetBookSheetExpression(wbName, aryShtName(i))
        Let guid = C_Commons.CreateGUID
        Dim j As Long
        ''
        '' Operate each sheets
        Console.log macroName
        If sheetType = "REPORT" Then
            ' ## Application.Run Macro of another book
            Let data_records = Application.Run("'" & wbNameForm & "'!" & macroName, wbName, aryShtName(i))
        Else
            Console.log ("Can not be applicable in config.json setting.")
        End If
        ''
        For j = LBound(data_records) To UBound(data_records)
            If VBA.IsObject(data_records(j)) Then
                data_records(j).Add "_source", keyWbSht
                data_records(j).Add "_source_date", wbUpdateDate
                data_records(j).Add "_created", createData
            End If
        Next j
        '' set "data_dic"
        dicData.Add keyWbSht, data_records
        ''
    Next i
    '' set "control_dic"
    dicControl.Add wbName, aryShtName
    ''
    GoSub FINALLY
    Exit Sub
    ''
FINALLY: 'Label
    ' ## C_Book.CloseBook
    Call C_Book.CloseBook(filePath:=wbPath, saveChanges:=False)
    Console.info "..... For """ & wbName & """ , it has done....."
    Console.info ""
    ''Package.Terminate
    Return
    ''
EXCEPTION: 'Label
    GoSub FINALLY
    Console.error ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.error ("Err.Number = " & Err.Number)
    Console.error ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.description)
    ''
End Sub

