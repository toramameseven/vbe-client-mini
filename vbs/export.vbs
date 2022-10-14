Option Explicit

'' this vbs is modified from below.
'' https://qiita.com/jinoji/items/23099771f78401bf0d34

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

Dim bookPath
'' if empty, export all modules.
'' if set, export modulePath module
'' module include sht.cls
Dim moduleFileName
Dim pathToExport

'' default parameter
bookPath = ""
'' if empty src_bookFileName
pathToExport = ""
'' if empty all module, .cls, .frm, .sht.cls
moduleFileName = ""

'' get arguments
If WScript.Arguments.Count = 1 Then
    bookPath = WScript.Arguments(0)
ElseIf WScript.Arguments.Count = 2 Then
    bookPath = WScript.Arguments(0)
    pathToExport = WScript.Arguments(1)
ElseIf WScript.Arguments.Count = 3 Then
    bookPath = WScript.Arguments(0)
    pathToExport = WScript.Arguments(1)
    moduleFileName = WScript.Arguments(2)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    pathToExport = fso.GetParentFolderName(bookPath) & "\src_" & "XXXXXXXXXXXXXXXXXXX"
    moduleFileName = "sheet1.sht.cls"
End If

'' debug output information
LogDebug "################", WScript.ScriptName
LogDebug "bookPath", bookPath
LogDebug "pathToExport", pathToExport
LogDebug "moduleFileName", moduleFileName

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
call Catch("Can not Create folders", 9999)

'' delete all modules or not
If moduleFileName = "" Then
  DeleteFilesInFolder dirModules
End if
call Catch("Can not delete module files.", 9999)

'' from vbsCommon.vbs 
OpenExcelFileWE bookPath
Dim book
Set book = GetObject(bookPath)
Dim objExcel 
set objExcel = book.Application

call Catch("Can Not Open Excel File.", 9999)

Dim vbComponent
Dim VBComponents
Set VBComponents = book.VBProject.VBComponents
Dim sheetObjContents
dim modulePath

Dim VBComponentToExport
If moduleFileName = "" Then
  Set VBComponentToExport = Nothing
Else
  Dim refIsSheetClass
  LogInfo "moduleFileName", GetModuleName(moduleFileName, refIsSheetClass)
  Set VBComponents = Nothing
  Set VBComponentToExport = book.VBProject.VBComponents.Item(GetModuleName(moduleFileName, refIsSheetClass))
End if

catch "Can not set VBComponent", 9999

If Not (VBComponents Is Nothing) Then
    For Each vbComponent In VBComponents
        call ExportFromVBComponent(vbComponent, dirModules, "")
    Next
End If
call Catch("Can not Export modules.", 9999)
On Error Goto 0

IF Not (VBComponentToExport Is Nothing)  Then
    call ExportFromVBComponent(VBComponentToExport, dirModules, moduleFileName)
End if
call Catch("Can not Export a module.", 9999)
On Error Goto 0

Set book = Nothing
Set objExcel = Nothing
Set fso = Nothing
WScript.StdOut.WriteLine "Export Complete."
WScript.Quit(0)
''---------------------------------------------end-------------------------------------


Function ExportFromVBComponent(vbComponent, dirModules, moduleFileName)
    Dim TypeOfModule
    Dim modulePath
    TypeOfModule = ResolveExtension(vbComponent.Type) 'include .sht.cls
    modulePath = dirModules & "\" & vbComponent.Name & TypeOfModule

    Dim isExport
    isExport = moduleFileName = "" Or LCase(modulePath) = LCase(dirModules & "\" & moduleFileName)

    If vbComponent.CodeModule.CountOfLines > 0 And isExport Then
        If TypeOfModule = "" Then
            ' do nothing
            LogDebug "Next is not exported1", vbComponent.Name
        ElseIF vbComponent.Type = 100  Then
            sheetObjContents = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
            ExportSheetModule modulePath, sheetObjContents
        ElseIf vbComponent.Type = 1  Then 'bas
            ExportNormalModule modulePath, vbComponent
        ElseIf vbComponent.Type = 2  Then 'cls
            ExportNormalModule modulePath, vbComponent
        ElseIf vbComponent.Type = 3    Then 'frm
            ExportNormalModule modulePath, vbComponent
        End If
    Else
        LogDebug "Next is not exported2", vbComponent.Name
        LogDebug "modulePath", modulePath
        LogDebug "moduleFileName", moduleFileName
    End If
End Function

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
    LogDebug "Export normal module: ", filePath
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
    LogDebug "Export sheet module: ", filePath
End Function

Function CreateFolder(folderPath)
    If (fso.FileExists(folderPath)) Then
      WScript.StdErr.WriteLine "Can not Create Folder. Same file exists"
      WScript.Quit(10)
    End If

    If (Not fso.FolderExists(folderPath)) Then
        fso.CreateFolder folderPath 
    End If
End Function

