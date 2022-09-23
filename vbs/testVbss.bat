pushd  %~dp0
set PATH_VBS=%~dp0
@REM pause
cscript /nologo %PATH_VBS%closeBook.vbs
copy %PATH_VBS%..\xlsms\macrotest_org.xlsm %PATH_VBS%..\xlsms\macrotest.xlsm

cscript /nologo %PATH_VBS%vbsCommon.vbs  
cscript /nologo %PATH_VBS%runVba.vbs
cscript /nologo %PATH_VBS%compile.vbs
cscript /nologo %PATH_VBS%export.vbs
cscript /nologo %PATH_VBS%getModules.vbs
cscript /nologo %PATH_VBS%import.vbs
cscript /nologo %PATH_VBS%remove.vbs