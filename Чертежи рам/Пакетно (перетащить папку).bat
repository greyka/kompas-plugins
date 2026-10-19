@echo off
chcp 65001 >nul
rem Пакетно: перетащите на этот файл папку со сборками .a3d (или один файл).
set "SRC=%~1"
if "%SRC%"=="" (echo Перетащите папку со сборками на этот файл. & pause & exit /b)
rem Python: рядом (установленный плагин) или в папке установки плагинов
set "PY=%~dp0..\Python\python.exe"
if not exist "%PY%" set "PY=C:\KompasPlugins\Python\python.exe"
if not exist "%PY%" set "PY=%LOCALAPPDATA%\KompasPlugins\Python\python.exe"
if not exist "%PY%" (echo Python для плагинов не установлен — запустите «Установить.bat». & pause & exit /b 1)
set PYTHONIOENCODING=utf-8
"%PY%" "%~dp0frame2cdw.py" "%SRC%"
pause
