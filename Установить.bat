@echo off
chcp 65001 >nul
rem Установка всех плагинов из этой папки в КОМПАС-3D (КОМПАС должен быть закрыт).
rem Python ставится один раз: сначала скачивается из интернета (Python\Онлайн),
rem если интернета нет — берётся из Python\Офлайн. Повторный запуск — обновить/добавить плагины.
call "%~dp0_setup\root.cmd"
set "PY=%KP_ROOT%\Python\python.exe"
if exist "%PY%" goto plugins
echo Python для плагинов не найден.
if exist "%~dp0Python\Онлайн\python-setup.ps1" (
    echo Пробую скачать из интернета...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Python\Онлайн\python-setup.ps1" -Target "%KP_ROOT%\Python"
)
if exist "%PY%" goto plugins
echo Интернета нет или скачать не удалось — ставлю из офлайн-копии...
call "%~dp0Python\Офлайн\offline-setup.cmd" "%KP_ROOT%\Python"
if not exist "%PY%" (echo. & echo ОШИБКА: Python не установлен ни из интернета, ни из папки Python\Офлайн. & pause & exit /b 1)
:plugins
"%PY%" "%~dp0_setup\install.py" --root "%KP_ROOT%" %*
