Option Explicit

'' this vbs is modified from below.
'' https://qiita.com/jinoji/items/23099771f78401bf0d34

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

'' get gook path
Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))

Dim bookPath
'' if empty, export all modules.
'' if set, export modulePath module
'' module include sht.cls
Dim moduleFileName
Dim pathToExport
Dim isExportSheet
Dim isExportForm

'' default parameter
bookPath = ""
'' if empty src_bookFileName
pathToExport = ""
'' if empty all module, .cls, .frm, .sht.cls
moduleFileName = ""
isExportSheet = True
isExportForm = True

If WScript.Arguments.Count = 1 Then
    bookPath = WScript.Arguments(0)
ElseIf WScript.Arguments.Count = 2 Then
    bookPath = WScript.Arguments(0)
    pathToExport = WScript.Arguments(1)
ElseIf WScript.Arguments.Count = 3 Then
    bookPath = WScript.Arguments(0)
    pathToExport = WScript.Arguments(1)
    moduleFileName = WScript.Arguments(2)
ElseIf WScript.Arguments.Count = 5 Then
    bookPath = WScript.Arguments(0)
    pathToExport = WScript.Arguments(1)
    moduleFileName = WScript.Arguments(2)
    isExportSheet = WScript.Arguments(3)
    isExportForm =  WScript.Arguments(4)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    pathToExport = fso.GetParentFolderName(bookPath) & "\src_" & "XXXXXXXXXXXXXXXXXXX"
    ''moduleFileName = "sheet1.sht.cls"
End If

'' now all export
isExportSheet = True
isExportForm = True

'' debug output information
DebugWriteLine "################", WScript.ScriptName
DebugWriteLine "bookPath", bookPath
DebugWriteLine "pathToExport", pathToExport
DebugWriteLine "moduleFileName", moduleFileName
DebugWriteLine "isExportSheet", isExportSheet
DebugWriteLine "isExportForm", isExportForm

'' test if bookPath exists
If fso.FileExists(bookPath) = False Then
    WScript.StdErr.WriteLine ("File does not exists: " & bookPath)
    WScript.Quit(10)
End If

'' create export folders
On Error Resume Next
Dim dirModules
If pathToExport = "" Then
  dirModules = fso.GetParentFolderName(bookPath) & "\src_" & fso.GetFileName(bookPath)
Else
  dirModules = pathToExport
End If
CreateFolder dirModules
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not Create folders"
    WScript.Quit(Err.Number)
End If
On Error Goto 0

On Error Resume Next
If moduleFileName = "" Then
  '' clear folder
  DeleteFilesInFolder dirModules
End if
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not delete module files."
    WScript.StdErr.WriteLine Err.description
    WScript.Quit(Err.Number)
End If
On Error Goto 0

'' from vbsCommon.vbs 
On Error Resume Next
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel 
set objExcel = book.Application

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can Not Open Excel File."
    WScript.Quit(Err.Number)
End If
On Error Goto 0


Dim vbComponent
Dim TypeOfModule
Dim VBComponents
Set VBComponents = book.VBProject.VBComponents
Dim sheetObjContents
dim modulePath

On Error Resume Next
For Each vbComponent In VBComponents
    TypeOfModule = ResolveExtension(vbComponent.Type) 'include .sht.cls
    modulePath = dirModules & "\" & vbComponent.Name & TypeOfModule

    Dim isExport
    isExport = moduleFileName = "" Or LCase(modulePath) = LCase(dirModules & "\" & moduleFileName)

    If vbComponent.CodeModule.CountOfLines > 0 And isExport Then
        If TypeOfModule = "" Then
            ' do nothing
            DebugWriteLine "Next is not exported1", vbComponent.Name
        ElseIF vbComponent.Type = 100 And isExportSheet Then
            sheetObjContents = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
            ExportSheetModule modulePath, sheetObjContents
        ElseIf vbComponent.Type = 1  Then 'bas
            ExportNormalModule modulePath, vbComponent
        ElseIf vbComponent.Type = 2  Then 'cls
            ExportNormalModule modulePath, vbComponent
        ElseIf vbComponent.Type = 3   And isExportForm Then 'frm
            ExportNormalModule modulePath, vbComponent
        End If
    Else
        DebugWriteLine "Next is not exported2", vbComponent.Name
        DebugWriteLine "modulePath", modulePath
        DebugWriteLine "moduleFileName", moduleFileName
    End If
Next

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not Export or Checkout."
    WScript.Quit(Err.Number)
End If
On Error Goto 0

Set book = Nothing
Set objExcel = Nothing
Set fso = Nothing
WScript.StdOut.WriteLine "Export Complete."
WScript.Quit(0)

'' modules

Function ResolveExtension(cType)
    Select Case cType
        Case 1      : ResolveExtension = ".bas"
        Case 2      : ResolveExtension = ".cls"
        Case 3      : ResolveExtension = ".frm"
        Case 100    : ResolveExtension = ".sht.cls"        
        Case Else   : ResolveExtension = ""
    End Select
End Function


Function ExportNormalModule(filePath, vbComponent)
    If filePath = "" Then
        Exit Function
    End if
    vbComponent.Export filePath
    DebugWriteLine "Export normal module", filePath
End Function

Function ExportSheetModule(filePath, contents)
    If filePath = "" Then
        Exit Function
    End if
    '' todo error
    Dim strFile
    strFile = filePath
    Dim objFS
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Dim objFile
    Set objFile = objFS.CreateTextFile(strFile)
    objFile.Write(contents)
    objFile.Close
    DebugWriteLine "Export sheet module", filePath
End Function

Function CreateFolder(folderPath)
    If (fso.FileExists(folderPath)) Then
      WScript.StdErr.WriteLine "Can not Create Folder. Same file exists"
      WScript.Quit(10)
    End If

    If (Not fso.FolderExists(folderPath)) Then
        fso.CreateFolder folderPath 
    End If

    '' already the folder exists
    '' good and return
End Function

