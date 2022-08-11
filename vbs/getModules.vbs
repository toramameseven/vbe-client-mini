Option Explicit

'' this vbs is modified from below.
'' https://qiita.com/jinoji/items/23099771f78401bf0d34

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\startExcelOpen.vbs").ReadAll()

'' get gook path
Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))
Dim bookPath
If WScript.Arguments.Count = 1 Then
    bookPath = WScript.Arguments(0)
Else
    '' for debug
    bookPath = fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm")
End If

If fso.FileExists(bookPath) = False Then
    WScript.StdEr.WriteLine "File does not exists: " & fso.GetFileName(bookPath) 
    WScript.Quit(10)
End If

Dim dirModules
dirModules = fso.GetParentFolderName(bookPath) & "\src_" & fso.GetFileName(bookPath)

'' from startExcelOpen.vbs
OpenExcelFile bookPath

Dim book
Set book = GetObject(bookPath)
Dim objExcel 
set objExcel = book.Application

'' export modules

If book.VBProject.Protection <> 0 Then
    objExcel.Close
    set objExcel = Nothing
    WScript.Quit(10)
end If

Dim VBComponents
Set VBComponents = book.VBProject.VBComponents
Dim count
count = VBComponents.Count

Set book = Nothing
Set objExcel = Nothing
Set fso = Nothing
WScript.StdOut.WriteLine count
WScript.Quit(0)
