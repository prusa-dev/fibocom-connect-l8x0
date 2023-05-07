@echo off

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0%scripts\main.ps1" -OnlyMonitor

if ERRORLEVEL 1 pause
