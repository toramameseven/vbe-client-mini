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
DebugWriteLine "################", WScript.ScriptName
DebugWriteLine "bookPath", bookPath
DebugWriteLine "pathToExport", pathToExport
DebugWriteLine "moduleFileName", moduleFileName

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
call ErrorHandle("Can not Create folders", True)

'' delete all modules or not
If moduleFileName = "" Then
  '' clear folder
  DeleteFilesInFolder dirModules
End if
call ErrorHandle("Can not delete module files.", True)

'' from vbsCommon.vbs 
OpenExcelFile bookPath
Dim book
Set book = GetObject(bookPath)
Dim objExcel 
set objExcel = book.Application

call ErrorHandle("Can Not Open Excel File.", True)



Dim vbComponent
Dim VBComponents
Set VBComponents = book.VBProject.VBComponents
Dim sheetObjContents
dim modulePath


Dim VBComponentToExport

If moduleFileName = "" Then
  Set VBComponentToExport = Nothing
Else
  Set VBComponents = Nothing
  Set VBComponentToExport = book.VBProject.VBComponents.Item(GetModuleName(moduleFileName))
End if

call ErrorHandle("Can not set VBComponent", True)

If Not (VBComponents Is Nothing) Then
    For Each vbComponent In VBComponents
        call ExportFromVBComponent(vbComponent, dirModules, "")
    Next
End If
call ErrorHandle("Can not Export modules.", True)
On Error Goto 0

IF Not (VBComponentToExport Is Nothing)  Then
    call ExportFromVBComponent(VBComponentToExport, dirModules, moduleFileName)
End if
call ErrorHandle("Can not Export a module.", True)
On Error Goto 0

Set book = Nothing
Set objExcel = Nothing
Set fso = Nothing
WScript.StdOut.WriteLine "Export Complete."
WScript.Quit(0)
''---------------------------------------------end-------------------------------------

Function ErrorHandle(message, isQuit)
  If Err.Number <> 0 Then
    WScript.StdErr.WriteLine message
    WScript.StdErr.WriteLine Error.Description
    If isQuit Then
      WScript.Quit(Err.Number)
    End If
  End if
End Function

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
            DebugWriteLine "Next is not exported1", vbComponent.Name
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
        DebugWriteLine "Next is not exported2", vbComponent.Name
        DebugWriteLine "modulePath", modulePath
        DebugWriteLine "moduleFileName", moduleFileName
    End If
End Function

'' sheet1.sht.cls > sheet1
'' module1.cls > module1
Function GetModuleName(moduleFileName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim base1
    Dim moduleName
    base1 = fso.GetBaseName(moduleFileName)

    If fso.GetExtensionName(base1) = "" Then
      moduleName = base1
    Else
      moduleName = fso.GetBaseName(base1)
    End If

    GetModuleName = moduleName
    Set fso = Nothing
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
    DebugWriteLine "Export normal module: ", filePath
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
    DebugWriteLine "Export sheet module: ", filePath
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

