@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0publish.ps1" %*
