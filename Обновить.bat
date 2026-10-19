@echo off
chcp 65001 >nul
rem Обновить плагины КОМПАС с GitHub (то же самое делается автоматически при входе в Windows).
call "%~dp0_setup\root.cmd"
set "PY=%KP_ROOT%\Python\python.exe"
if not exist "%PY%" (echo Плагины не установлены: сначала запустите «Установить.bat». & pause & exit /b 1)
set "UPD=%KP_ROOT%\_setup\update.py"
if not exist "%UPD%" set "UPD=%~dp0_setup\update.py"
"%PY%" "%UPD%" --root "%KP_ROOT%" %*
