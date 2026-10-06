@echo off
cd /d "%~dp0"
netstat -ano | findstr ":3100 " | findstr "LISTENING" >nul
if errorlevel 1 (
  start "Yong Dental Meeting Server" /min cmd /c "npm run dev -- -p 3100"
  timeout /t 10 /nobreak >nul
)
start "" "http://localhost:3100"
