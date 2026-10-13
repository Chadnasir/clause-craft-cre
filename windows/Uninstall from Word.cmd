@echo off
setlocal
title Remove Clause Craft from Word
rem Removes the Clause Craft add-in from desktop Word for this Windows user.

set "ADDIN_ID=EEFD4951-BE4B-4D39-8D49-24350738C07F"
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /f >nul 2>&1
if exist "%LOCALAPPDATA%\ClauseCraft\manifest.xml" del /q "%LOCALAPPDATA%\ClauseCraft\manifest.xml"
if exist "%LOCALAPPDATA%\ClauseCraft" rmdir "%LOCALAPPDATA%\ClauseCraft" 2>nul

echo.
echo   Clause Craft was removed from Word. Close and reopen Word to finish.
echo.
pause
exit /b 0
