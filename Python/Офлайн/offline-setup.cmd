@echo off
rem Ставит Python для плагинов из папки runtime (без интернета). Параметр: папка назначения.
set "TARGET=%~1"
if "%TARGET%"=="" (echo Не указана папка назначения & exit /b 1)
if not exist "%~dp0runtime\python.exe" (echo Нет офлайн-копии Python: %~dp0runtime & exit /b 1)
echo Устанавливаю Python из офлайн-копии в %TARGET% ...
if exist "%TARGET%" rmdir /s /q "%TARGET%"
xcopy "%~dp0runtime" "%TARGET%\" /e /i /q /y >nul || exit /b 1
"%TARGET%\python.exe" -c "import win32com.client, pythoncom" || (echo Python скопирован, но pywin32 не загружается & exit /b 1)
echo Python готов: %TARGET%
exit /b 0
