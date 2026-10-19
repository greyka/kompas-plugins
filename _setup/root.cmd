@echo off
rem Определяет папку установки плагинов в переменную KP_ROOT:
rem C:\KompasPlugins, а если туда нельзя писать — %LOCALAPPDATA%\KompasPlugins
set "KP_ROOT=C:\KompasPlugins"
mkdir "%KP_ROOT%" 2>nul
(echo.>"%KP_ROOT%\.write_test") 2>nul && (del "%KP_ROOT%\.write_test" & exit /b 0)
set "KP_ROOT=%LOCALAPPDATA%\KompasPlugins"
mkdir "%KP_ROOT%" 2>nul
exit /b 0
