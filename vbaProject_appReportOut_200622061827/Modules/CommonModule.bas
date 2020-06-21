Attribute VB_Name = "CommonModule"
Option Explicit

Public Function SetCommonDataSet()
'''' **********************************************
''
cdset.AppendDictionary "control_dic"
cdset.AppendDictionary "data_dic"
cdset.AppendDictionary "output_dic"
''
End Function

Public Function SetEnvVariables()
'''' **********************************************
Dim CONF As Dictionary
Set CONF = cdset.GetValue("CONF")
Dim C_String As C_String
Set C_String = New C_String
Dim C_File As C_File
Set C_File = New C_File
''
Dim book_path As String
Let book_path = C_File.GetLocalName(ThisWorkbook.FullName)
Dim book_folder As String
Let book_folder = C_File.GetParentFolder(book_path)
''
Dim base_folder As String
If CONF.Exists("BASE_FOLDER") Then
    Let base_folder = CONF.item("BASE_FOLDER")
Else
    Let base_folder = ""
End If
Let base_folder = C_String.DefaultString(base_folder, book_folder)
''
Dim input_folder As String
Let input_folder = C_File.BuildPath(base_folder, CONF.item("INPUT_FOLDER"))
''
Dim output_folder As String
Let output_folder = C_File.BuildPath(base_folder, CONF.item("OUTPUT_FOLDER"))
''
Dim temp_folder As String
Let temp_folder = C_File.BuildPath(base_folder, CONF.item("TEMP_FOLDER"))
''
Dim backup_folder As String
Let backup_folder = C_File.BuildPath(base_folder, CONF.item("BACKUP_FOLDER"))
''
Dim form_folder As String
Let form_folder = C_File.BuildPath(base_folder, CONF.item("FORM_FOLDER"))
''
Dim data_folder As String
Let data_folder = C_File.BuildPath(base_folder, CONF.item("DATA_FOLDER"))
Dim backup_data_folder As String
Let backup_data_folder = C_File.BuildPath(base_folder, CONF.item("BACKUP_DATA_FOLDER"))
''
Call C_File.CreateFolder(base_folder)
Call C_File.CreateFolder(input_folder)
Call C_File.CreateFolder(temp_folder)
Call C_File.CreateFolder(backup_folder)
Call C_File.CreateFolder(form_folder)
Call C_File.CreateFolder(data_folder)
Call C_File.CreateFolder(backup_data_folder)
''
Call cdset.PutValue(book_folder, "book_folder")
Call cdset.PutValue(input_folder, "input_folder")
Call cdset.PutValue(output_folder, "output_folder")
Call cdset.PutValue(form_folder, "form_folder")
Call cdset.PutValue(temp_folder, "temp_folder")
Call cdset.PutValue(backup_folder, "backup_folder")
Call cdset.PutValue(data_folder, "data_folder")
Call cdset.PutValue(backup_data_folder, "backup_data_folder")
''
End Function

Public Function OutputJson(Optional ByVal macroName As Variant)
'''' **********************************************
''
Dim C_Sheet As C_Sheet
Set C_Sheet = New C_Sheet
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
Dim dictControl As Dictionary
Set dictControl = cdset.GetValue("control_dic")
Dim dictData As Dictionary
Set dictData = cdset.GetValue("data_dic")
''
Dim json_str As String
''
Dim output_folder As String
Let output_folder = cdset.GetValue("output_folder")
Dim output_file_ext As String
If VBA.IsMissing(macroName) Then
    Let output_file_ext = "_Report" & format(Now(), "_YYMMDDHHmmSS") & ".json"
ElseIf macroName = "GetDef" Then
    Let output_file_ext = "_Report" & format(Now(), "_YYMMDDHHmmSS") & ".json"
ElseIf macroName = "DumpSimpleJson" Then
    Let output_file_ext = "_SimpleReport" & format(Now(), "_YYMMDDHHmmSS") & ".json"
Else
    ''
End If
''
Dim clt As Collection
Set clt = New Collection
Dim vntTemp As Variant
Dim output_file As String
Dim output_path As String
Dim wbNameAry As Variant
Let wbNameAry = dictControl.keys
Dim wbName As Variant
Dim shtNameAry As Variant
Dim wbShtkey As String
Dim i As Long
Dim j As Long
For Each wbName In wbNameAry
    Let shtNameAry = dictControl.item(wbName)
    For i = LBound(shtNameAry) To UBound(shtNameAry)
        Let wbShtkey = C_Sheet.GetBookSheetExpression(wbName, shtNameAry(i))
        If VBA.IsObject(dictData.item(wbShtkey)) Then
            Set vntTemp = dictData.item(wbShtkey)
            clt.Add C_Dictionary.DeepCopy(vntTemp)
        ElseIf VBA.IsArray(dictData.item(wbShtkey)) Then
            Let vntTemp = dictData.item(wbShtkey)
            For j = LBound(vntTemp) To UBound(vntTemp)
                clt.Add vntTemp(j)
            Next j
        End If
        '' convert to json
        Let json_str = JsonConverter.ConvertToJson(clt, Whitespace:=4)
        '' output json format
        Let output_file = C_File.GetBaseName(wbName) & "_" & shtNameAry(i) & output_file_ext
        Let output_path = C_File.BuildPath(output_folder, output_file)
        C_FileIO.WriteTextAllAsUTF8NoneBOM output_path, json_str
        Set clt = New Collection
    Next i
Next wbName
''
End Function

