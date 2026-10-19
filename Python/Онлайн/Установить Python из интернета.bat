@echo off
chcp 65001 >nul
rem Скачать с python.org нужную версию Python (+ pywin32) и поставить для плагинов.
rem Обычно не нужно — «Установить.bat» сам ставит Python. Этот файл — переустановить вручную.
call "%~dp0..\..\_setup\root.cmd"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0python-setup.ps1" -Target "%KP_ROOT%\Python"
if errorlevel 1 (echo. & echo ОШИБКА: не удалось скачать. Используйте папку Python\Офлайн.)
pause
