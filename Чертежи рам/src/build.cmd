@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" >nul
cd /d "%~dp0"
rc /nologo /fo FrameDrawings.res FrameDrawings.rc || exit /b 1
cl /nologo /O2 /EHsc /utf-8 /LD /DUNICODE /D_UNICODE FrameDrawings.cpp FrameDrawings.res /link /DEF:FrameDrawings.def /OUT:..\FrameDrawings.rtw user32.lib ole32.lib || exit /b 1
del /q *.obj *.res *.exp *.lib 2>nul
