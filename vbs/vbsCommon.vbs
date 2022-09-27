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
    exit Sub
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


Sub OpenVbe(objExcel)
  '' FindControl(type, id. ......)
  '' open vbe
  Dim ctrl
  Set ctrl = objExcel.Application.CommandBars.FindControl(, 1695)
  If Err.Number <> 0 Or ctrl Is Nothing Then
      WScript.StdErr.WriteLine ("Check objExcel Object Model Security")
      WScript.StdErr.WriteLine Err.Description
  Else
      If ctrl.Enabled = True Then
          ctrl.Execute
      Else
          '' not enable
      End IF
  End If
End Sub


Sub TestCompiled(objExcel)

    on error resume Next
    Dim ctrlCompile
    ''compile 578
    Set ctrlCompile = objExcel.VBE.ActiveVBProject.VBE.CommandBars.FindControl(, 578)
    If Err.Number <> 0 Or ctrlCompile Is Nothing Then
        WScript.StdErr.WriteLine ("Can not find compile command")
        WScript.Quit(1001)
    Else
        If ctrlCompile.Enabled = True Then
          ctrlCompile.Execute
        Else
          '' now compiled go ahead
        End IF
    End If

    If ctrlCompile.Enabled = True Then
      Call OpenVbe(objExcel)
      WScript.StdErr.WriteLine "Can not compile VBA. May be a compile error!! Check VBE"
      WScript.Quit(100)
    End if

    on error goto 0
End Sub

Sub TestRunningVba(objExcel)

    on error resume Next
    Dim ctrlPause
    ''pause 189
    Set ctrlPause = objExcel.VBE.ActiveVBProject.VBE.CommandBars.FindControl(, 189)
    If Err.Number <> 0 Or ctrlPause Is Nothing Then
        WScript.StdErr.WriteLine ("Can not find pause command")
        WScript.Quit(1001)
    Else
        If ctrlPause.Enabled = True Then
          '' vba not running, go a head.
        Else
          '' now pause, so not enabled.
          Call OpenVbe(objExcel)
          WScript.StdErr.WriteLine ("Can not run. May be paused in VBE.")
          WScript.Quit(1001)
        End IF
    End If
    on error goto 0

    '' test vba run or not
    on error resume Next
    Dim ctrlContinue
    ''continue 186
    Set ctrlContinue = objExcel.VBE.ActiveVBProject.VBE.CommandBars.FindControl(, 186)
    If Err.Number <> 0 Or ctrlContinue Is Nothing Then
        WScript.StdErr.WriteLine ("Can not find continue command")
        WScript.Quit(1001)
    Else
        If ctrlContinue.Enabled = True Then
          '' vba not running, go a head.
        Else
          '' now pause, so not enabled.
          Call OpenVbe(objExcel)
          WScript.StdErr.WriteLine ("Can not run. May be running in VBE.")
          WScript.Quit(1001)
        End IF
    End If
    on error goto 0
End Sub




