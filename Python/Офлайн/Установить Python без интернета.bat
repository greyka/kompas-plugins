@echo off
chcp 65001 >nul
rem Поставить Python для плагинов из этой папки (runtime), интернет не нужен.
rem Обычно не нужно — «Установить.bat» сам ставит Python. Этот файл — переустановить вручную.
call "%~dp0..\..\_setup\root.cmd"
call "%~dp0offline-setup.cmd" "%KP_ROOT%\Python"
if errorlevel 1 (echo. & echo ОШИБКА: Python не установлен.)
pause
