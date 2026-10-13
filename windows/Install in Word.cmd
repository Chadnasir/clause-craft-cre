@echo off
setlocal
title Install Clause Craft in Word
rem Adds the Clause Craft add-in to desktop Word for this Windows user. No admin rights needed.
rem Word loads the add-in pages from https://chadnasir.github.io/clause-craft-cre/ .
rem The review itself runs inside Word; your documents are not uploaded anywhere.

set "ADDIN_ID=EEFD4951-BE4B-4D39-8D49-24350738C07F"
set "SRC=%~dp0Clause Craft manifest.xml"
if not exist "%SRC%" set "SRC=%~dp0..\word\manifest.xml"
set "DEST=%LOCALAPPDATA%\ClauseCraft"

if not exist "%SRC%" goto nomanifest

if not exist "%DEST%" mkdir "%DEST%"
copy /y "%SRC%" "%DEST%\manifest.xml" >nul
if errorlevel 1 goto failed

reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%DEST%\manifest.xml" /f >nul
if errorlevel 1 goto failed

echo.
echo   Clause Craft is installed for Word.
echo.
echo   1. Close Word completely - every Word window.
echo   2. Open Word again and open a lease or purchase agreement.
echo   3. On the Home tab, click Review Contract.
echo.
echo   Word needs an internet connection to open the Clause Craft panel.
echo.
pause
exit /b 0

:nomanifest
echo.
echo   Can't find "Clause Craft manifest.xml" next to this file.
echo   Unzip the whole Clause Craft folder first, then run this again.
echo.
pause
exit /b 1

:failed
echo.
echo   Windows blocked the install, so nothing was changed.
echo   You can still use Clause Craft in Word on the web:
echo   Home, Add-ins, More Settings, Upload My Add-in, then pick "Clause Craft manifest.xml".
echo.
pause
exit /b 1
