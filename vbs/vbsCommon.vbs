Option Explicit

Sub OpenExcelFile(bookPath)
  '' is objExcel running?
  Dim objExcel
  Dim IsRunExcel
  On Error Resume Next
  Set objExcel = GetObject(,"Excel.Application")
  IsRunExcel = True
  If Err.Number <> 0 OR objExcel Is Nothing Then
    '' if false, launch excel
    IsRunExcel = False
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
    DebugWriteLine "CreateObject error", "Can not get excel instance"
    WScript.Quit(100)
  End If
  objExcel.visible = True

  Dim fso
  Set fso = createObject("Scripting.FileSystemObject")

  On Error Goto 0
  '' if xlam and no book, create new book.
  Dim ext
  Dim bookTemp
  ext = LCase(fso.GetExtensionName(bookPath)) ' without dot extension
  DebugWriteLine "Workbooks in Excel", objExcel.Workbooks.Count
  If objExcel.Workbooks.Count = 0 And ext = "xlam" Then
    set bookTemp = objExcel.Workbooks.Add()
    DebugWriteLine "Add book to Excel instance", ""
  End if

  '' check, is open target Excel file 
  Dim bookName
  bookName = fso.GetFileName(bookPath)


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
  If IsOpenFile = False Then
    Set myWorkBook = objExcel.Workbooks.Open(bookPath)
  Else
    Set myWorkBook = GetObject(bookPath)
  End if

  If Err.Number <> 0 OR myWorkBook Is Nothing Then
    WScript.StdErr.WriteLine "Can not Get Excel Book."
    Set objExcel = Nothing
    WScript.Quit(100)
  End If
  
  Set objExcel = Nothing
  On Error Goto 0
End Sub


Sub CloseExcelFile(bookName)
  '' is objExcel running?
  Dim objExcel
  Dim IsRunExcel
  On Error Resume Next
  Set objExcel = GetObject(,"Excel.Application")
  If objExcel Is Nothing Then
      DebugWriteLine "Err", Err.Description
      WScript.Quit(100)
  End If
  
  If Err.Number <> 0 Then
      DebugWriteLine "Err", Err.Description
      WScript.Quit(Err.Number)
  End If
  On Error Goto 0

  On Error Resume Next
  Dim wb, IsOpenFile
  IsOpenFile = False
  For Each wb in objExcel.Workbooks
    If LCase(wb.Name) = LCase(bookName) Then    
      IsOpenFile = True
      wb.Close False
      Exit For
    End if
  Next

  If Err.Number <> 0 Then
      DebugWriteLine "Err", Err.Description
      WScript.Quit(Err.Number)
  End If  
End Sub

''for debug
''DeleteFilesInFolder "C:\projects\toramame-hub\xy\xlsms\src_macroTest.xlsm"

Function DeleteFilesInFolder(folderPath)
    DebugWriteLine "To Removed", folderPath
    If folderPath = "" Then
        Exit Function
    End if
    Dim objFolder, fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder(folderPath)
    Dim objFile
    Dim objFileName
    For each objFile in objFolder.files
        objFileName = objFile.Name
        fso.DeleteFile folderPath & "\" & objFileName
        DebugWriteLine "Removed", objFileName
    Next
End Function

Sub DebugWriteLine(title, value)
    ''exit Sub
    Dim outTitle
    Dim outValue
    outTitle = title
    outValue = value
    If outTitle = "" Then
        outTitle = "(_empty_)"
    End if
    If outValue = "" Then
        outValue = "(_empty_)"
    End if
    WScript.StdOut.WriteLine "VBS::    " & outTitle & " : " & outValue
End Sub

Function GetProjectRoot(fso)
  GetProjectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))
End Function




