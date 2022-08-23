Option Explicit
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
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

WScript.echo bookName

CloseExcelFile bookName

Sub CloseExcelFile(bookName)
  '' is objExcel running?
  Dim objExcel
  Dim IsRunExcel
  On Error Resume Next
  Set objExcel = GetObject(,"Excel.Application")
  If objExcel Is Nothing Then
      WScript.Quit(0)
  End If
  If Err.Number <> 0 Then
      WScript.Quit(Err.Number)
  End If
  On Error Goto 0

  Dim wb, IsOpenFile
  IsOpenFile = False
  For Each wb in objExcel.Workbooks
    If LCase(wb.Name) = LCase(bookName) Then    
      IsOpenFile = True
      wb.Close()
      Exit For
    End if
  Next

  WScript.Quit(0)
End Sub





