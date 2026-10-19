@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" >nul
cd /d "%~dp0"
rc /nologo /fo CdwPdf.res CdwPdf.rc || exit /b 1
cl /nologo /O2 /EHsc /utf-8 /LD /DUNICODE /D_UNICODE CdwPdf.cpp CdwPdf.res /link /DEF:CdwPdf.def /OUT:..\CdwPdf.rtw user32.lib ole32.lib || exit /b 1
del /q *.obj *.res *.exp *.lib 2>nul
