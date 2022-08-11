Option Explicit

'' test
''TestFileOpen

'' for debug
Function TestFileOpen()
    WScript.StdOut.WriteLine "run vbs"
    Dim fso
    Set fso = createObject("Scripting.FileSystemObject")
    Dim projectRoot
    projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))

    Dim bookPath
    If WScript.Arguments.Count = 1 Then
        bookPath = WScript.Arguments(0)
    Else
        bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
    End If
    OpenExcelFile bookPath
End Function

Function OpenExcelFile(bookPath)
  '' is objExcel running
  Dim objExcel
  Dim IsRunExcel
  On Error Resume Next
  Set objExcel = GetObject(,"Excel.Application")
  IsRunExcel = True
  If Err.Number <> 0 OR objExcel Is Nothing Then
    IsRunExcel = False
    WScript.StdOut.WriteLine "not run objExcel"
  End If
  On Error Goto 0

  '' if not running, start objExcel
  On Error Resume Next
  If IsRunExcel = False Then
    Set objExcel = CreateObject("Excel.Application")
  End If

  If Err.Number <> 0 OR objExcel Is Nothing Then
    WScript.StdErr.WriteLine "Can not run objExcel"
    objExcel = Nothing
    WScript.Quit(10)
  End If
  objExcel.visible = True
  On Error Goto 0

  '' check, is open target Excel file 
  Dim fso
  Set fso = createObject("Scripting.FileSystemObject")
  Dim bookName
  bookName = fso.GetFileName(bookPath)

  '' if xlam and no book, create new book.
  Dim ext
  Dim bookTemp
  ext = LCase(fso.GetExtensionName(bookPath)) ' without dot extension
  DebugWriteLine "ext", ext
  DebugWriteLine "Workbooks", objExcel.Workbooks.Count
  If objExcel.Workbooks.Count = 0  and ext = "xlam" Then
    set bookTemp = objExcel.Workbooks.Add()
    DebugWriteLine "", "add book"
  End if

  Dim wb, IsOpenFile
  IsOpenFile = False
  For Each wb in objExcel.Workbooks
    If LCase(wb.Name) = LCase(bookName) Then    
      IsOpenFile = True
      Exit For
    End if
  Next
  
  '' get book instance
  On Error Resume Next
  Dim myWorkBook
  DebugWriteLine "bookPath", bookPath
  If IsOpenFile = False Then
    Set myWorkBook = objExcel.Workbooks.Open(bookPath)
  Else
    Set myWorkBook = GetObject(bookPath)
  End if

  If Err.Number <> 0 OR myWorkBook Is Nothing Then
    WScript.StdErr.WriteLine "Can not Get Excel Book."
    Set objExcel = Nothing
    WScript.Quit(10)
  End If
  
  Set objExcel = Nothing
  On Error Goto 0
End Function

''for debug
''DeleteFilesInFolder "C:\projects\toramame-hub\xy\xlsms\src_macroTest.xlsm"

Function DeleteFilesInFolder(folderPath)
    If folderPath = "" Then
        Exit Function
    End if
    Dim objFolder, fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder(folderPath)
    Dim objFile
    For each objFile in objFolder.files
        fso.DeleteFile folderPath & "\" & objFile.Name
    Next
End Function

Sub DebugWriteLine(title, value)
    '' WScript.StdOut.WriteLine "DEBUG:: " & title & " : " & value
End Sub




