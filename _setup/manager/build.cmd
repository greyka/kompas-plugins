@echo off
chcp 65001 >nul
rem Сборка «Менеджер плагинов КОМПАС.exe» в корень папки плагинов (csc из .NET Framework 4.x).
cd /d "%~dp0"
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
"%CSC%" /nologo /codepage:65001 /target:winexe /optimize+ /win32icon:icon.ico /out:"..\..\Менеджер плагинов КОМПАС.exe" /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:System.Core.dll Manager.cs
