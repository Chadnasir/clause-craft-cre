@echo off
setlocal
title Install Clause Craft in Word
rem Adds the Clause Craft add-in to desktop Word for this Windows user (no admin rights needed),
rem then opens Word with the Clause Craft panel showing on a demo lease.
rem Word loads the add-in pages from https://chadnasir.github.io/clause-craft-cre/ .
rem The review itself runs inside Word; your documents are not uploaded anywhere.

set "ADDIN_ID=EEFD4951-BE4B-4D39-8D49-24350738C07F"
set "SRC=%~dp0Clause Craft manifest.xml"
if not exist "%SRC%" set "SRC=%~dp0..\word\manifest.xml"
set "DEMO=%~dp0Clause Craft Demo Lease.docx"
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

rem Word only loads new add-ins when it starts, so restart it if it is open.
tasklist 2>nul | find /i "WINWORD.EXE" >nul
if errorlevel 1 goto openword

echo   Word is open and needs to restart once to load Clause Craft.
echo   Save your documents first. Word will ask about anything unsaved.
echo.
choice /c YN /n /m "  Restart Word now? Press Y for yes, N for later: "
if errorlevel 2 goto later
taskkill /im WINWORD.EXE >nul 2>&1
for /l %%i in (1,1,90) do (
  tasklist 2>nul | find /i "WINWORD.EXE" >nul || goto openword
  timeout /t 1 /nobreak >nul
)
echo.
echo   Word is still open. Close it, then double-click "Clause Craft Demo Lease" in this folder.
echo.
pause
exit /b 0

:openword
if "%CLAUSECRAFT_NO_OPEN%"=="1" goto done
if exist "%DEMO%" (start "" "%DEMO%") else (start "" winword)
:done
echo   Word is opening with the Clause Craft panel on the right.
echo   From now on, open any lease or PSA and click Home, then Review Contract.
echo.
echo   This window closes by itself.
timeout /t 10 >nul
exit /b 0

:later
echo.
echo   OK. The next time you open Word, click Home, then Review Contract.
echo.
timeout /t 10 >nul
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
