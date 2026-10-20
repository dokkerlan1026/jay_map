@echo off
chcp 65001 > nul
echo 正在啟動花蓮店家地圖...
cd /d "%~dp0"
node server.js
pause
