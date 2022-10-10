Option Explicit

'' load function
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\vbsCommon.vbs").ReadAll()

''test
LogDebug "All_Test>>>>>>>>>>>>>>>>>>>>>>>>>>>>", "Start"

TestFileOpen
call ActivateVbeProject_Test()
Call TestIfCompiled()
call TestRunningVba()
CloseExcelFileWE "macrotest.xlsm"

LogDebug "All_Test<<<<<<<<<<<<<<<<<<<<<<<<<<<<", "End"


''for debug
''DeleteFilesInFolder "C:\projects\toramame-hub\xy\xlsms\src_macroTest.xlsm"

'' for debug
Function TestFileOpen()
    WScript.StdOut.WriteLine "Test run vbs"
    Dim projectRoot
    projectRoot = GetProjectRoot
    call OpenExcelFileWE(fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm"))
    'LogDebug "OpenExcelFileWE:", OpenExcelFileWE(fso.BuildPath(projectRoot, "xlsms\macroTest.xlsm"))
End Function

Sub ActivateVbeProject_Test()
    Dim bookPath
    Dim moduleName
    bookPath = fso.BuildPath(GetProjectRoot, "xlsms\xlsAmMenu.xlam")
    moduleName = "clsAllAdoc"
    OpenExcelFileWE bookPath
    call OpenVbe()
    ActivateVbeProjectWE bookPath
    Call Catch("ActivateVbeProject_Test", 9999)
    LogDebug "ActivateVbeProject_Test", "End"
End Sub