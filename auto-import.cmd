@echo off
setlocal
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"
start "" "%SCRIPT_DIR%\MidAutumnGiftBox.exe" --quick-run "%SCRIPT_DIR%"
endlocal
