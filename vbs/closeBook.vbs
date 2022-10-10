Option Explicit

Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs", 1).ReadAll()

' declare at vbCommon
' Dim projectRoot

Dim bookPath
If WScript.Arguments.Count = 1 Then
    bookPath = WScript.Arguments(0)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
End If

Dim bookName
bookName = fso.GetFileName(bookPath)

LogDebug "################ start", WScript.ScriptName
LogDebug "bookName", bookName
LogDebug "bookPath", bookPath

On Error Resume Next
CloseExcelFileWE bookName
Catch "Can not close excel.",9765

WScript.StdOut.WriteLine "Excel Close Complete."
LogDebug "----------------End", WScript.ScriptName
Catch "closeBook.vbs", 9999'
WScript.Quit(0)
'////////////////////////////////////////////////////////////////////////////





