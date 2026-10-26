@echo off
rem doc.haus launcher for Windows: selects x64 Bun 1.4.0 or newer like start.sh, then
rem runs scripts\windows\start.ts. An explicit BUN_BIN is never replaced.
setlocal
cd /d "%~dp0"
set "BUN_ROOT=%USERPROFILE%\.bun"
if defined BUN_INSTALL set "BUN_ROOT=%BUN_INSTALL%"
set "BUN_HOME=%BUN_ROOT%\bin\bun.exe"
:: Some native dependencies ship for x64 Windows only, so Windows on Arm runs
:: x64 Bun under emulation, kept beside Bun's own install.
set "ARM64="
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "ARM64=1"
if /i "%PROCESSOR_ARCHITEW6432%"=="ARM64" set "ARM64=1"
if defined ARM64 set "BUN_HOME=%BUN_ROOT%\x64\bun.exe"
if defined BUN_BIN call :supported "%BUN_BIN%" && goto run
if defined BUN_BIN echo doc.haus could not run x64 Bun 1.4.0 or newer at %BUN_BIN%. Set BUN_BIN to a working x64 Bun executable, then run start.cmd again. 1>&2
if defined BUN_BIN exit /b 1
for %%B in (bun.exe) do set "BUN_PATH=%%~$PATH:B"
if defined BUN_PATH call :supported "%BUN_PATH%" && goto run
call :supported "%BUN_HOME%" && goto run
echo doc.haus requires x64 Bun 1.4.0 or newer; none was found on PATH or at %BUN_HOME%. 1>&2
if /i "%DOCHAUS_INSTALL_TOOLS%"=="no" goto missing
:: Without a console, install only when DOCHAUS_INSTALL_TOOLS=yes, as start.sh does.
if /i not "%DOCHAUS_INSTALL_TOOLS%"=="yes" powershell -NoProfile -c "exit [int][Console]::IsInputRedirected" || goto missing
set "ANSWER=y"
if /i not "%DOCHAUS_INSTALL_TOOLS%"=="yes" set /p "ANSWER=doc.haus: install the latest Bun from https://bun.sh now? [Y/n] "
if /i not "%ANSWER:~0,1%"=="y" goto missing
if defined ARM64 goto install_x64
echo + powershell -NoProfile -ExecutionPolicy Bypass -c "irm bun.sh/install.ps1|iex" 1>&2
powershell -NoProfile -ExecutionPolicy Bypass -c "irm bun.sh/install.ps1|iex"
goto installed
:install_x64
:: Bun's installer picks the Arm build here, so fetch its x64 release directly.
echo + downloading bun-windows-x64.zip from https://github.com/oven-sh/bun/releases/latest to %BUN_HOME% 1>&2
powershell -NoProfile -ExecutionPolicy Bypass -c "$ErrorActionPreference = 'Stop'; $zip = Join-Path $env:TEMP 'bun-windows-x64.zip'; $dir = Join-Path $env:TEMP 'bun-windows-x64'; Invoke-WebRequest -UseBasicParsing https://github.com/oven-sh/bun/releases/latest/download/bun-windows-x64.zip -OutFile $zip; Expand-Archive $zip $dir -Force; New-Item -ItemType Directory -Force (Split-Path $env:BUN_HOME) | Out-Null; Copy-Item (Join-Path $dir 'bun-windows-x64\bun.exe') $env:BUN_HOME -Force"
:installed
call :supported "%BUN_HOME%" && goto run
echo doc.haus: the Bun install did not produce a working x64 Bun 1.4.0+ at %BUN_HOME%. 1>&2
:missing
echo Install Bun or set BUN_BIN to a working x64 Bun executable, then run start.cmd again. 1>&2
exit /b 1
:run
"%BUN_EXE%" scripts\windows\start.ts %*
exit /b %errorlevel%
:supported
"%~1" -e "const [major, minor] = Bun.version.split('.').map(Number); process.exit(process.arch === 'x64' && (major > 1 || (major === 1 && minor >= 4)) ? 0 : 1)" >nul 2>&1 || exit /b 1
set "BUN_EXE=%~1"
exit /b 0
