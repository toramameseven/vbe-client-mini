'Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

'' get book path
' declare at vbCommon
' Dim projectRoot
Dim moduleName
Dim lineCodePane

Dim bookPath
If WScript.Arguments.Count = 3 Then
  bookPath = WScript.Arguments(0)
  moduleName = WScript.Arguments(1) ''"module2"
  lineCodePane = WScript.Arguments(2) ''120
Else
  '' for debug, run with no arguments
  bookPath = fso.BuildPath(projectRoot, "xlsms\xlsAmMenu.xlam")
  moduleName = "clsAllAdoc"
  lineCodePane = 120
End If

Dim bookName
bookName = fso.GetFileName(bookPath)

'' debug output information
LogDebug "################", WScript.ScriptName
LogDebug "bookPath", bookPath
LogDebug "moduleName", moduleName
LogDebug "lineCodePane", lineCodePane


'' open the book or attach the book
On Error Resume Next
'' from vbsCommon.vbs
OpenExcelFileWE bookPath


call OpenVbe()

ActivateVbeProjectWE bookPath

'' select line
objExcel.VBE.ActiveVBProject.VBComponents(moduleName).Activate
objExcel.VBE.ActiveVBProject.VBComponents(moduleName).CodeModule.CodePane.SetSelection lineCodePane, 1, lineCodePane, 10
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "Can not select line"
    WScript.StdErr.WriteLine Err.Description
End if


WScript.Sleep 1000
WScript.StdOut.WriteLine "goto complete"
On Error GoTo 0
WScript.Quit(0)



