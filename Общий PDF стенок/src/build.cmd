@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" >nul
cd /d "%~dp0"
rc /nologo /fo WallsPdf.res WallsPdf.rc || exit /b 1
cl /nologo /O2 /EHsc /utf-8 /LD /DUNICODE /D_UNICODE WallsPdf.cpp WallsPdf.res /link /DEF:WallsPdf.def /OUT:..\WallsPdf.rtw user32.lib ole32.lib || exit /b 1
del /q *.obj *.res *.exp *.lib 2>nul
