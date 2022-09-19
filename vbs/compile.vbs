Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\startExcelOpen.vbs").ReadAll()

'' get book path
Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))
Dim bookPath
If WScript.Arguments.Count = 1 Then
  bookPath = WScript.Arguments(0)
Else
  '' for debug, run with no arguments
  bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
End If

Dim bookName
bookName = fso.GetFileName(bookPath)

'' open the book or attach the book
On Error Resume Next
'' from startExcelOpen.vbs
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel
Set objExcel = book.Application

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine ("Can not Open: " & bookName & " : " & bookPath)
    book.Close
    WScript.Quit(Err.Number)
End If
On Error Goto 0

'' compile project
On Error Resume Next
'' FindControl(type, id. ......)
Dim ctrl
Set ctrl = objExcel.VBE.ActiveVBProject.VBE.CommandBars.FindControl(, 578)
If Err.Number <> 0 Or ctrl Is Nothing Then
    WScript.StdErr.WriteLine ("Check objExcel Object Model Security")
Else
    If ctrl.Enabled = True Then
        ctrl.Execute
    Else
        '' already complied!!
    End IF
End If

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not compile VBA"
End if

WScript.Sleep 1000
WScript.StdOut.WriteLine "Compile complete"
On Error GoTo 0
WScript.Quit(0)



