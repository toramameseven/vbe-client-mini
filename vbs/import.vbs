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
Dim isUseFromModule
Dim isUseSheetModule
isUseFromModule = True
isUseSheetModule = True
If WScript.Arguments.Count = 3 Then
    bookPath = WScript.Arguments(0)
    isUseFromModule = WScript.Arguments(1)
    isUseSheetModule = WScript.Arguments(2)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
End If

On Error Resume Next

'' from startExcelOpen.vbs
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel
Set objExcel = book.Application

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not Open Excel: " & fso.GetFileName(bookPath)
    book.Close
    WScript.Quit(Err.Number)
End If
On Error Goto 0

On Error Resume Next
'' delete all modules form xlsm
Call DeleteVbaModules(book, isUseFromModule, isUseSheetModule)
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not delete modules: " & fso.GetFileName(bookPath)
    book.Close
    WScript.Quit(Err.Number)
End If
On Error Goto 0

'' get modules folder
Dim moduleFolderPath
moduleFolderPath = fso.GetParentFolderName(bookPath) & "\src_" & fso.GetFileName(bookPath) 

''' Import VBA module files
Call importVbaModules(fso.GetFolder(moduleFolderPath), book, bookPath, isUseFromModule, isUseSheetModule)

book.Save
WScript.StdOut.WriteLine "Import Complete"
WScript.Quit(0)


''///////////////////////////  
Function DeleteVbaModules(book, isUseFromModule, isUseSheetModule)
    Dim vBComponents 
    Set vBComponents = book.VBProject.VBComponents
    
    Dim vbComponent 
    For Each vbComponent In vBComponents
        If vbComponent.Type = 100 Then
            If isUseSheetModule Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
            End If
        ElseIf vbComponent.Type = 3Then
            If isUseFromModule Then
                vBComponents.Remove vbComponent
            End If
        Else
            '' 2(cls), 1(bas)
            vBComponents.Remove vbComponent
        End If
    Next 
End Function

Public Sub importVbaModules(modulesFolder, book, excelBookPath, isUseFromModule, isUseSheetModule)
    'As VBIDE.VBComponents
    Dim vBComponents
    Set vBComponents = book.VBProject.VBComponents
    
    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Dim objFile
    Dim fileExtension
    For Each objFile In fso.GetFolder(modulesFolder).Files
        
        fileExtension = LCase(fso.GetExtensionName(objFile.Name))
        If  fileExtension = "bas" Then
            Call vBComponents.Import (objFile.Path)
        ElseIf fileExtension = "frm" Then
            Call importFormModule(vBComponents, objFile.Path, isUseFromModule)
        ElseIf (fileExtension = "cls") Then
            importCodeToExcelObjects excelBookPath, objFile.Path, isUseSheetModule
        End If
    Next

    Set modulesFolder = Nothing
    Set objFile = Nothing
End Sub

'' import sheet module or book module
Public Function importCodeToExcelObjects(excelBookPath, modulePath, isUseSheetModule)
    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim sourceFile
    Set sourceFile  = fso.OpenTextFile(modulePath)
    Dim fileContent
    fileContent = sourceFile.ReadAll
    Dim className
    className = fso.GetBaseName(modulePath) 'filename without .ext
    sourceFile.Close
    Set sourceFile = Nothing
    set fso = Nothing

    Dim book
    Set book = GetObject(excelBookPath)

    Dim vbComponent
    Dim IsSheet
    IsSheet = False
    For Each vbComponent in book.VBProject.VBComponents
        If className = vbComponent.Name And vbComponent.Type = 100 Then
            IsSheet = True
            Exit For
        End if
    Next

    If IsSheet then
        If isUseSheetModule Then
            vbComponent.CodeModule.InsertLines 1, fileContent
        End If
    Else
        book.VBProject.VBComponents.Import modulePath
    End if

    Set book = Nothing
    Set vbComponent = Nothing
End Function

Public Function importFormModule(vBComponents, modulePath, isOptionImport)
    If isOptionImport = False Then
        Exit Function
    End if

    Call vBComponents.Import (modulePath)

    Dim fso    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim moduleName
    moduleName = fso.GetBaseName(modulePath) 'filename without .ext

    With VBComponents(moduleName).CodeModule
        If .CountOfLines <= 1 Then
            Exit Function
        End IF
        If Trim(.Lines(1, 1)) = "" Then
            Call .DeleteLines(1, 1)
        End if
    End With
End Function



