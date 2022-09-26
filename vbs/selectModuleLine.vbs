'Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

'' get book path
Dim projectRoot
projectRoot = GetProjectRoot(fso)
Dim moduleName
Dim lineCodePane

Dim bookPath
If WScript.Arguments.Count = 3 Then
  bookPath = WScript.Arguments(0)
  moduleName = WScript.Arguments(1) ''"module2"
  lineCodePane = WScript.Arguments(2) ''120
Else
  '' for debug, run with no arguments
  bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
  moduleName = "module2"
  lineCodePane = 120
End If

Dim bookName
bookName = fso.GetFileName(bookPath)

'' debug output information
DebugWriteLine "################", WScript.ScriptName
DebugWriteLine "bookPath", bookPath
DebugWriteLine "moduleName", moduleName
DebugWriteLine "lineCodePane", lineCodePane


'' open the book or attach the book
On Error Resume Next
'' from vbsCommon.vbs
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel
Set objExcel = book.Application

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not Open: " & bookName & " : " & bookPath)
    WScript.Quit(Err.Number)
End If


'' FindControl(type, id. ......)
'' open vbe
Dim ctrl
Set ctrl = objExcel.Application.CommandBars.FindControl(, 1695)
If Err.Number <> 0 Or ctrl Is Nothing Then
    WScript.StdErr.WriteLine ("Check objExcel Object Model Security")
    WScript.StdErr.WriteLine Err.Description
Else
    If ctrl.Enabled = True Then
        ctrl.Execute
    Else
        '' not enable
    End IF
End If

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not open vbe"
    WScript.Quit(Err.Number)
End if

'' select line
objExcel.VBE.ActiveVBProject.VBComponents(moduleName).Activate
objExcel.VBE.ActiveVBProject.VBComponents(moduleName).CodeModule.CodePane.SetSelection lineCodePane, 1, lineCodePane, 10
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not select line"
End if


WScript.Sleep 1000
WScript.StdOut.WriteLine "goto complete"
On Error GoTo 0
WScript.Quit(0)



