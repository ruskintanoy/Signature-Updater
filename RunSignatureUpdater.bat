@echo off
set "ScriptPath=%~dp0ClassicOutlookSignatureUpdater-GPO.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%ScriptPath%"