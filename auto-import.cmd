@echo off
setlocal
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"
start "" "%SCRIPT_DIR%\MidAutumnGiftBox.exe" --auto-import "%SCRIPT_DIR%"
endlocal
