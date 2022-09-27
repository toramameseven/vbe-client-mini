Option Explicit

Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))

'' get book path and function to run
Dim bookPath
Dim FunctionName
If WScript.Arguments.Count = 2 Then
    bookPath = WScript.Arguments(0)
    FunctionName = WScript.Arguments(1)
Else
    '' debug for click this script.
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    FunctionName = "Module1.Test4"
End If

'' debug output information
DebugWriteLine "################", WScript.ScriptName
DebugWriteLine "bookPath", bookPath
DebugWriteLine "FunctionName", FunctionName


'' from vbsCommon.vbs
Call OpenExcelFile(bookPath)

'' get book instance
Dim myWorkBook
Set myWorkBook = GetObject(bookPath)
Dim objExcel
Set objExcel = myWorkBook.Application

'' test vba run or not
'' in running, stop(end) quit script
call TestCompiled(objExcel)
call TestRunningVba(objExcel)

' set Excel top view
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.AppActivate myWorkBook.Name, True



'' run functionName on myWorkBook
Dim ret
ret = objExcel.Run(myWorkBook.Name & "!" & FunctionName)
' if run Function return value
'' and output the value, stdout
DebugWriteLine "Return Value:", CStr(ret)
DebugWriteLine "If Return Value:", "out the value to stdout"
WScript.StdOut.WriteLine ret
WScript.Quit(0)



