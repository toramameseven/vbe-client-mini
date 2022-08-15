Option Explicit

'' this vbs is modified from below.
'' https://qiita.com/jinoji/items/23099771f78401bf0d34

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\startExcelOpen.vbs").ReadAll()

'' get gook path
Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))
Dim bookPath

'' if empty, export all modules.
'' if set, export modulePath module
'' module include sht.cls
Dim modulePathSelect
Dim isAllExport ' 1: all(work, base, current), 0: only current
If WScript.Arguments.Count = 1 Then
    bookPath = WScript.Arguments(0)
    isAllExport = 1
    modulePathSelect = ""
ElseIf WScript.Arguments.Count = 2 Then
    bookPath = WScript.Arguments(0)
    isAllExport = WScript.Arguments(1)
    modulePathSelect = ""
ElseIf WScript.Arguments.Count = 3 Then
    bookPath = WScript.Arguments(0)
    isAllExport = WScript.Arguments(1)
    modulePathSelect = WScript.Arguments(2)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    isAllExport = 0
    modulePathSelect = ""
End If
isAllExport = 0

'' debug output information
DebugWriteLine "bookPath", bookPath
DebugWriteLine "isAllExport", isAllExport
DebugWriteLine "modulePathSelect", modulePathSelect

If fso.FileExists(bookPath) = False Then
    WScript.StdErr.WriteLine ("File does not exists: " & bookPath)
    WScript.Quit(10)
End If

On Error Resume Next
'' create export folders
Dim dirModules
dirModules = fso.GetParentFolderName(bookPath) & "\src_" & fso.GetFileName(bookPath)
Dim dirModulesBase
dirModulesBase = dirModules & "\.base"
Dim dirModulesCurrent
dirModulesCurrent = dirModules & "\.current"

CreateFolder dirModules
''for commit check, if dirModulesBase files differ from dirModulesCurrent, dirModulesCurrent is modified.
'CreateFolder dirModulesBase
'CreateFolder dirModulesCurrent

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not Create folders"
    WScript.Quit(Err.Number)
End If
On Error Goto 0


On Error Resume Next
If modulePathSelect = "" Then
  '' clear folder
  If isAllExport Then
      DeleteFilesInFolder dirModulesBase
      DeleteFilesInFolder dirModulesCurrent
  End if
  DeleteFilesInFolder dirModules
End if

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not delete module files."
    WScript.StdErr.WriteLine Err.description
    WScript.Quit(Err.Number)
End If
On Error Goto 0

'' from startExcelOpen.vbs
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


'' export modules

' If book.VBProject.Protection <> 0 Then
'     set objExcel = Nothing
'     WScript.Quit(10)
' end If

Dim vbComponent
Dim TypeOfModule
Dim VBComponents
Set VBComponents = book.VBProject.VBComponents
Dim sheetObjContents
dim modulePath
dim modulePathBase
dim modulePathCurrent


On Error Resume Next
For Each vbComponent In VBComponents

    TypeOfModule = ResolveExtension(vbComponent.Type)
    modulePath = dirModules & "\" & vbComponent.Name & TypeOfModule
    modulePathBase = "" ''dirModulesBase & "\" & vbComponent.Name & TypeOfModule
    modulePathCurrent ="" '' dirModulesCurrent & "\" & vbComponent.Name & TypeOfModule

    Dim isExport
    isExport = modulePathSelect = "" Or LCase(modulePath) = LCase(modulePathSelect)

    If vbComponent.CodeModule.CountOfLines > 0 And isExport Then
        If TypeOfModule = "" Then
            ' do nothing
            DebugWriteLine "Next is not exported", vbComponent.Name
        ElseIF vbComponent.Type = 100 Then
            sheetObjContents = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
            If isAllExport Then
                ExportSheetModule modulePathCurrent, sheetObjContents
                ExportSheetModule modulePathBase, sheetObjContents
            End If
            ExportSheetModule modulePath, sheetObjContents
        Else
            If isAllExport Then
                ExportNormalModule vbComponent, modulePathCurrent
                ExportNormalModule vbComponent, modulePathBase
            End if
            ExportNormalModule vbComponent, modulePath
        End If
    Else
        DebugWriteLine "Next is not exported", vbComponent.Name
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
WScript.StdOut.WriteLine "Export Complete"
WScript.Quit(0)

Function ResolveExtension(cType)
    Select Case cType
        Case 1      : ResolveExtension = ".bas"
        Case 2      : ResolveExtension = ".cls"
        Case 3      : ResolveExtension = ".frm"
        Case 100    : ResolveExtension = ".sht.cls"        
        Case Else   : ResolveExtension = ""
    End Select
End Function


Function ExportNormalModule(vbComponent, filePath)
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

