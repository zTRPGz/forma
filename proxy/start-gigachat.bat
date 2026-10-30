@echo off
cd /d "%~dp0"
if not exist "gigachat-key.txt" (
  echo GigaChat Authorization key is required.
  echo Get it at https://developers.sber.ru/ and paste into the opened file.
  echo ^# Paste the key on a new line below, save and close Notepad.> gigachat-key.txt
  notepad gigachat-key.txt
)
node gigachat-proxy.js
echo.
echo Proxy stopped. Press any key to close.
pause >nul
