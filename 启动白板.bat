@echo off
cd /d "%~dp0"
echo Starting Whiteboard...
node_modules\electron\electron.exe .
pause