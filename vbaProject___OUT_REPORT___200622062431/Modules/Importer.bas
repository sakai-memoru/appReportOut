Attribute VB_Name = "Importer"
Option Explicit
'-----------------------------------------------------------
' tool : importer
'-----------------------------------------------------------

Public Sub Export()
''' *************************************************
Dim exporter As Tool_ModuleExporter
Set exporter = New Tool_ModuleExporter
exporter.Export

End Sub

Private Sub autoRefered()
''' *************************************************
On Error Resume Next
Application.VBE.ActiveVBProject.References.AddFromFile _
"C:\WINDOWS\System32\scrrun.dll"
''
Debug.Print "Have referred to Microsoft.Scripting.Runtime"
On Error GoTo 0
''
End Sub

Public Function DeleteClasses()
'''' *************************************************
''
Dim CONS_ARY_CLASS_MODULES() As Variant
Let CONS_ARY_CLASS_MODULES = Array( _
    "C_Commons.cls", "C_Console.cls", "C_String.cls", "O_DataSet.cls", "C_Config.cls", _
    "C_Array.cls", "C_Range.cls", "C_Name.cls", "C_ListObject.cls", _
    "C_Dictionary.cls", "C_Collection.cls", "C_Book.cls", "C_Sheet.cls", _
    "C_Template.cls", "O_Template.cls", "MiniTemplator.cls", _
    "C_File.cls", "C_FileIO.cls", "Tool_ModuleExporter.cls", "O_StringBuilder.cls")
Dim itm As Variant
Dim itmBase() As String
ReDim itmBase(LBound(CONS_ARY_CLASS_MODULES) To UBound(CONS_ARY_CLASS_MODULES))
Dim i As Long
For i = LBound(CONS_ARY_CLASS_MODULES) To UBound(CONS_ARY_CLASS_MODULES)
    Let itmBase(i) = Split(CONS_ARY_CLASS_MODULES(i), ".")(0)
Next i
Set C_File = Nothing
''
For i = LBound(CONS_ARY_CLASS_MODULES) To UBound(CONS_ARY_CLASS_MODULES)
    Debug.Print "> remove ... " & CONS_ARY_CLASS_MODULES(i)
    Call RemoveModule(itmBase(i))
Next i
''
End Function

Public Sub Import(Optional ByVal classOnly As Variant = True, Optional ByVal removeClassOn As Variant = False)
''' *************************************************
Const CONS_MODULES As String = "Modules"
Const CONS_CLASS_MODULES As String = "ClassModules"
Dim CONS_ARY_MODULES As Variant
Dim CONS_ARY_CLASS_MODULES As Variant
'CONS_ARY_MODULES = Array("GlobalVariableMin.bas")
CONS_ARY_MODULES = Array("workspace.bas", "JsonConverter.bas", "Package.bas")
CONS_ARY_CLASS_MODULES = Array( _
    "C_Commons.cls", "C_Console.cls", "C_String.cls", "O_DataSet.cls", "C_Config.cls", _
    "C_Array.cls", "C_Range.cls", "C_Name.cls", "C_ListObject.cls", _
    "C_Dictionary.cls", "C_Collection.cls", "C_Book.cls", "C_Sheet.cls", _
    "C_Template.cls", "O_Template.cls", "MiniTemplator.cls", _
    "C_File.cls", "C_FileIO.cls", "Tool_ModuleExporter.cls", "O_StringBuilder.cls")
''
Dim wbTarget As Excel.Workbook
Dim strTargetWorkbook As String
Dim strImportPathOfModules As String
Dim strImportPathOfClassModules As String
Dim strFilename As String
Dim VBComponents As Object
Set VBComponents = ThisWorkbook.VBProject.VBComponents
''
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
''
''Get the path to the folder with modules
If FolderWithVBAProject = "Error" Then
    Debug.Print "Import Folder not exist"
    Exit Sub
End If
''
'''' NOTE: This workbook must be open in Excel.
strTargetWorkbook = ActiveWorkbook.name
Set wbTarget = Application.Workbooks(strTargetWorkbook)
''
If wbTarget.VBProject.Protection = 1 Then
    Debug.Print "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
End If
''
strImportPathOfModules = FolderWithVBAProject & "\" & CONS_MODULES
strImportPathOfClassModules = FolderWithVBAProject & "\" & CONS_CLASS_MODULES
''
If objFSO.GetFolder(strImportPathOfModules).Files.Count = 0 Then
   MsgBox "There are no files to import"
   Exit Sub
End If
''
''Delete all modules/Userforms from the ActiveWorkbook
''Call DeleteVBAModulesAndUserForms
''
Set VBComponents = wbTarget.VBProject.VBComponents
''
'''' Import all the code modules in the specified path
'''' to the ActiveWorkbook.
Dim itm As Variant
If Not classOnly Then
    For Each itm In CONS_ARY_MODULES
        Debug.Print "> import ... " & itm
        VBComponents.Import strImportPathOfModules & "\" & itm
    Next itm
End If
''
If removeClassOn Then
    Call DeleteClasses
End If
''
For Each itm In CONS_ARY_CLASS_MODULES
    Debug.Print "> import ... " & itm
    VBComponents.Import strImportPathOfClassModules & "\" & itm
Next itm
''
Call autoRefered
''
Debug.Print "Have Refered to Microsoft.Scripting.Runtime."
Debug.Print "--------"
Debug.Print "Have completed ! You have taken in required basic modules."
Debug.Print "Let's execute 'call helloWorld'"
''
'' release and end process
Set objFSO = Nothing
End Sub

Private Function FolderWithVBAProject() As String
'''' *************************************************
Dim objWShell As Object
Dim objFSO As Object
Dim special_folder As String
'
Set objWShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("scripting.filesystemobject")
special_folder = objWShell.SpecialFolders("Desktop")
''
If Right(special_folder, 1) <> "\" Then
    special_folder = special_folder & "\"
End If
''
If objFSO.FolderExists(special_folder & "VBAProject") = False Then
    On Error Resume Next
    MkDir special_folder & "VBAProject"
    On Error GoTo 0
End If
''
If objFSO.FolderExists(special_folder & "VBAProject") = True Then
    FolderWithVBAProject = special_folder & "VBAProject"
Else
    FolderWithVBAProject = "Error"
End If
Set objWShell = Nothing
Set objFSO = Nothing
''
End Function

Public Function RemoveModule(ByVal itemName As String)
'''' *************************************************
''
Dim cmp As Object
Dim VBComps As Object
Set VBComps = ThisWorkbook.VBProject.VBComponents
''
For Each cmp In VBComps
    If cmp.name = itemName Then
        VBComps.Remove cmp
    End If
Next cmp
''
End Function

