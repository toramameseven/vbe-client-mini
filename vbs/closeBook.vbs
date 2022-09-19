Option Explicit

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\startExcelOpen.vbs").ReadAll()

Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))

Dim bookPath
Dim bookName
If WScript.Arguments.Count = 1 Then
    bookPath = WScript.Arguments(0)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
End If

bookName = fso.GetFileName(bookPath)

DebugWriteLine "################ start", WScript.ScriptName
DebugWriteLine "bookName", bookName
DebugWriteLine "bookPath", bookPath

'On Error Resume Next
Dim r
CloseExcelFile bookName
DebugWriteLine "----------------r", r

If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not close excel."
    DebugWriteLine "----------------Err", "Can not close excel."
    WScript.Quit(Err.Number)
End If

WScript.StdOut.WriteLine "Excel Close Complete."
DebugWriteLine "----------------End", WScript.ScriptName
WScript.Quit(0)
'////////////////////////////////////////////////////////////////////////////





