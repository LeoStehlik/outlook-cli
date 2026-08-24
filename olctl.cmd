@echo off
rem Native Windows entry point for cmd.exe. Put this folder on PATH, call `olctl ...`.
rem From a PowerShell prompt call olctl.ps1 directly instead: this wrapper adds a
rem cmd.exe hop, and cmd.exe re-parses & < > | ^ in arguments as redirection or
rem command separators, which can mangle folder names and subjects.
setlocal
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0olctl.ps1" %*
exit /b %errorlevel%
