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
modulePath = ""
If WScript.Arguments.Count = 2 Then
    bookPath = WScript.Arguments(0)
    '' module path include xxxx.sht.cls
    modulePath = WScript.Arguments(1)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    modulePath = "C:\projects\toramame-hub\vbe-client-mini\xlsms\src_macrotest.xlsm\Module2.bas"
End If

'' debug output information
DebugWriteLine "################", WScript.ScriptName
DebugWriteLine "bookPath", bookPath
DebugWriteLine "modulePath", modulePath

On Error Resume Next
'' from startExcelOpen.vbs
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel
Set objExcel = book.Application

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not Open Excel: " & bookPath)
    WScript.Quit(Err.Number)
End If
On Error Goto 0

On Error Resume Next
Call DeleteVbaModules(book, modulePath)

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not delete modules: " & bookPath)
    WScript.Quit(Err.Number)
End If
On Error Goto 0

WScript.StdOut.WriteLine "Remove Complete"
WScript.Quit(0)

''///////////////////////////
'' if moduleName is empty, delete all modules.
Function DeleteVbaModules(book, modulePath)
    Dim vBComponents 
    Set vBComponents = book.VBProject.VBComponents

    Dim moduleName
    Dim isSrcSheetCls
    moduleName = GetModuleName(modulePath, isSrcSheetCls) 'filename without .ext
    
    Dim vbComponent 
    For Each vbComponent In vBComponents
        If (moduleName = "" Or LCase(moduleName) = LCase(vbComponent.Name)) Then
            If vbComponent.Type = 100 Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
                DebugWriteLine "Delete Content",vbComponent.Name
            ElseIf vbComponent.Type = 3 Then
                DebugWriteLine "Remove vbComponent",vbComponent.Name
                vBComponents.Remove vbComponent
            Else
                '' 2(cls), 1(bas)
                DebugWriteLine "Remove vbComponent",vbComponent.Name
                vBComponents.Remove vbComponent
            End If
        End If
    Next 
End Function

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






