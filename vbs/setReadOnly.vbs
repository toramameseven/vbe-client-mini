Option Explicit

'' load external vbs.
Dim fso
Set fso = createObject("Scripting.FileSystemObject")
Execute fso.OpenTextFile(fso.getParentFolderName(WScript.ScriptFullName) & "\startExcelOpen.vbs").ReadAll()


Dim filePath
Dim isReadOnly
filePath = ""
isReadOnly = ""

' Normal	0	
' ReadOnly	1	
' Hidden	2	
' System	4	
' Volume	8	
' Directory	16	
' Archive	32	
' Alias	      1024	
' Compressed  2048	

' value Or Flg  ON
' value And Flg GET
' value And Not Flg OFF

Dim projectRoot
projectRoot = fso.getParentFolderName(fso.getParentFolderName(WScript.ScriptFullName))

'' get parameters
If WScript.Arguments.Count = 2 Then
    filePath = WScript.Arguments(0)
    '1: readonly, 0:not readonly
    isReadOnly = WScript.Arguments(1)
Else
    '' for debug
    filePath = projectRoot + "\vbs\test.txt"
    isReadOnly = "0"
End If

'' debug output information
DebugWriteLine "vbs Script", "setReadOnly.vbs"
DebugWriteLine "filePath", filePath
DebugWriteLine "isReadOnly", isReadOnly


'' test if path exists
If fso.FileExists(filePath) = False Then
    WScript.StdErr.WriteLine ("File does not exists: " & filePath)
    WScript.Quit(10)
End If

'' set read only on or off
Dim objFile
Dim thisAttribute
set objFile = fso.GetFile(filePath)
thisAttribute = objFile.Attributes
If isReadOnly Then
    objFile.Attributes = thisAttribute Or 1
Else
    objFile.Attributes = thisAttribute And Not 1
End if

set objFile = Nothing
set fso = Nothing

WScript.Quit(0)
