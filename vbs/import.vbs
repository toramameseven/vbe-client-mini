Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\startExcelOpen.vbs").ReadAll()

'' get project root
Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))

'' get book path
Dim bookPath
'' if modulePath is empty, import all modules
Dim modulePath
Dim isUseFromModule
Dim isUseSheetModule
modulePath = ""
isUseFromModule = True
isUseSheetModule = True
If WScript.Arguments.Count = 2 Then
    bookPath = WScript.Arguments(0)
    '' module path include xxxx.sht.cls
    modulePath = WScript.Arguments(1)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    modulePath = "C:\projects\toramame-hub\vbe-client-mini\xlsms\src_macrotest.xlsm\Module2.bas"
End If


'' get modules folder
Dim moduleFolderPath
moduleFolderPath = fso.GetParentFolderName(bookPath) & "\src_" & fso.GetFileName(bookPath) 

IF fso.FolderExists(moduleFolderPath) = False Then
    WScript.StdErr.WriteLine ("No src Folder: " & moduleFolderPath)
    WScript.Quit(10)
End If


'' 
On Error Resume Next

'' from startExcelOpen.vbs
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel
Set objExcel = book.Application

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not Open Excel: " & bookPath)
    book.Close
    WScript.Quit(Err.Number)
End If
On Error Goto 0

On Error Resume Next
'' delete modules form xlsm

Call DeleteVbaModules(book, modulePath, isUseFromModule, isUseSheetModule)

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not delete modules: " & bookPath)
    WScript.Quit(Err.Number)
End If
On Error Goto 0

On Error Resume Next
''' Import VBA module files
Call importVbaModules(fso.GetFolder(moduleFolderPath), book, bookPath, modulePath, isUseFromModule, isUseSheetModule)
book.Save

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not import modules: " & bookPath)
    DebugWriteLine "importVbaModules Err", Err.description
    WScript.Quit(Err.Number)
End If
On Error Goto 0

WScript.StdOut.WriteLine "Import Complete"
WScript.Quit(0)


''///////////////////////////
'' if moduleName is empty, delete all modules.
Function DeleteVbaModules(book, modulePath, isUseFromModule, isUseSheetModule)
    Dim vBComponents 
    Set vBComponents = book.VBProject.VBComponents

    Dim moduleName
    Dim isSrcSheetCls
    moduleName = GetModuleName(modulePath, isSrcSheetCls) 'filename without .ext
    
    Dim vbComponent 
    For Each vbComponent In vBComponents
        If (moduleName = "" Or LCase(moduleName) = LCase(vbComponent.Name)) Then
            If vbComponent.Type = 100 Then
                If isUseSheetModule Then
                    vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
                    DebugWriteLine "Delete Content",vbComponent.Name
                End If
            ElseIf vbComponent.Type = 3 Then
                If isUseFromModule Then
                    DebugWriteLine "Remove vbComponent",vbComponent.Name
                    vBComponents.Remove vbComponent
                End If
            Else
                '' 2(cls), 1(bas)
                DebugWriteLine "Remove vbComponent",vbComponent.Name
                vBComponents.Remove vbComponent
            End If
        End If
    Next 
End Function

'' if modulePath is empty, import all modules.
Public Sub importVbaModules(modulesFolder, book, excelBookPath, modulePath, isUseFromModule, isUseSheetModule)
    'As VBIDE.VBComponents
    Dim vBComponents
    Set vBComponents = book.VBProject.VBComponents
    
    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '' xxxx.sht.cls must convert to xxx
    Dim moduleName
    moduleName = ""
    If excelBookPath <> "" Then
      moduleName = LCase(fso.GetFileName(modulePath))
    End If

    Dim objFile
    Dim fileExtension
    For Each objFile In fso.GetFolder(modulesFolder).Files
        '' selected module or all modules
        If moduleName = LCase(objFile.Name) Or moduleName = "" Then
          fileExtension = LCase(fso.GetExtensionName(objFile.Name))
          If  fileExtension = "bas" Then
              Call vBComponents.Import (objFile.Path)
              DebugWriteLine "Import bas module", objFile.Path
          ElseIf fileExtension = "frm" Then
              Call importFormModule(vBComponents, objFile.Path, isUseFromModule)
          ElseIf (fileExtension = "cls") Then
              '' cls file include a normal cls or sheet cls
              importClassModule excelBookPath, objFile.Path, isUseSheetModule
          End If
        End if
    Next

    Set modulesFolder = Nothing
    Set objFile = Nothing
End Sub


Function GetModuleName(modulePath, refIsSheetClass)
    '' sheet module is exported with .sht.cls
    '' test this extension
    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim withoutFirstExt
    withoutFirstExt = fso.GetBaseName(modulePath) 'filename without .ext
    GetModuleName = withoutFirstExt

    Dim secondExt
    secondExt = LCase(fso.GetExtensionName(withoutFirstExt))
    
    refIsSheetClass = secondExt = "sht"
    If refIsSheetClass  Then
      GetModuleName = fso.GetBaseName(withoutFirstExt)
    End If
End Function

'' import class module
Public Function importClassModule(excelBookPath, modulePath, isUseSheetModule)
    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '' check .sht.cls or normal.cls
    Dim moduleName
    Dim isSrcSheetCls
    moduleName = GetModuleName(modulePath, isSrcSheetCls)

    Dim book
    Set book = GetObject(excelBookPath)

    Dim vbComponent
    Dim IsSheet
    IsSheet = False
    For Each vbComponent in book.VBProject.VBComponents
        If moduleName = vbComponent.Name And vbComponent.Type = 100 Then
            IsSheet = True
            Exit For
        End if
    Next

    If IsSheet and isSrcSheetCls then
        If isUseSheetModule Then
            Dim sourceFile
            Set sourceFile  = fso.OpenTextFile(modulePath)
            Dim fileContent
            fileContent = sourceFile.ReadAll
            sourceFile.Close
            Set sourceFile = Nothing
            DebugWriteLine "Import sheet cls Content", vbComponent.Name
            vbComponent.CodeModule.InsertLines 1, fileContent
        End If
    ElseIf isSrcSheetCls Then
        '' module is sheet cls, but the book has not the sheet name
        '' do not import
    Else
        DebugWriteLine "Import cls module", modulePath
        book.VBProject.VBComponents.Import modulePath
    End if

    Set book = Nothing
    Set vbComponent = Nothing
    set fso = Nothing
End Function

Public Function importFormModule(vBComponents, modulePath, isOptionImport)
    If isOptionImport = False Then
        Exit Function
    End if

    Call vBComponents.Import (modulePath)
    DebugWriteLine "Import frm module", modulePath 

    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim moduleName
    moduleName = fso.GetBaseName(modulePath) 'filename without .ext

    With VBComponents(moduleName).CodeModule
        If .CountOfLines <= 1 Then
            Exit Function
        End IF

        '' when import, one line is added. here delete the line. 
        If Trim(.Lines(1, 1)) = "" Then
            Call .DeleteLines(1, 1)
        End if
    End With
End Function



