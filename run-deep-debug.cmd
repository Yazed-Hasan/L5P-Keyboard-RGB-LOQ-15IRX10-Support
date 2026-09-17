@echo off
setlocal
cd /d "%~dp0"
powershell -NoExit -ExecutionPolicy Bypass -File "%~dp0enable-deep-debug.ps1"
