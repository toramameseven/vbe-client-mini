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

'' from vbsCommon.vbs
Call OpenExcelFile(bookPath)

'' get book instance
Dim myWorkBook
Set myWorkBook = GetObject(bookPath)
Dim objExcel
Set objExcel = myWorkBook.Application

' set Excel top view
Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.AppActivate myWorkBook.Name, True

'' run functionName on myWorkBook
Dim ret
ret = objExcel.Run(myWorkBook.Name & "!" & FunctionName)
WScript.StdOut.WriteLine ret
WScript.Quit(0)



