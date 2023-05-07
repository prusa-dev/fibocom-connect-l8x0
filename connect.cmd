@echo off

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0%scripts\main.ps1"

if ERRORLEVEL 1 pause
