Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

' declare at vbCommon
' Dim projectRoot

Dim bookPath
If WScript.Arguments.Count = 1 Then
  bookPath = WScript.Arguments(0)
Else
  '' for debug, run with no arguments
  bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
End If

Dim bookName
bookName = fso.GetFileName(bookPath)

'' debug output information
LogDebug "################", WScript.ScriptName
LogDebug "bookPath", bookPath

'' open the book or attach the book
On Error Resume Next
OpenExcelFileWE bookPath
On Error Goto 0

'' compile project
On Error Resume Next
ActivateVbeProjectWE bookPath

'' FindControl(type, id. ......)
'' 578 compile
Dim ctrlCompile
Set ctrlCompile = GetVbeMenuCtrlWE(578)

If ctrlCompile.Enabled = True Then
    ctrlCompile.Execute
Else
    '' already complied!!
End IF
Catch "Can not compile VBA", 8888

'' check Compile
If ctrlCompile.Enabled = True Then
    WScript.StdErr.WriteLine "Can not compile VBA. May be a compile error!! Check VBE"
    WScript.Quit(100)
End IF

WScript.StdOut.WriteLine "Compile complete"
On Error GoTo 0
WScript.Quit(0)



