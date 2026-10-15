@echo off
chcp 65001 >nul
title Payment Tool - PVI Digital
cd /d "%~dp0"
if not exist "%~dp0app\pyembed\pythonw.exe" goto missing
if not exist "%~dp0app\node.exe" goto missing
start "" "%~dp0app\pyembed\pythonw.exe" "app\server.py"
exit /b

:missing
echo Missing files in the app folder. Extract the WHOLE zip again (to a short path, e.g. C:\Payment-Tool), then run this file.
pause
