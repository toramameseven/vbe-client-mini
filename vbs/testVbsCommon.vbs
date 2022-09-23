Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

''test
TestFileOpen
CloseExcelFile "macrotest.xlsm"

'' for debug
Function TestFileOpen()
    WScript.StdOut.WriteLine "Test run vbs"
    Dim projectRoot
    projectRoot = GetProjectRoot(fso)
    DebugWriteLine "OpenExcelFile:", OpenExcelFile(fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm"))
End Function