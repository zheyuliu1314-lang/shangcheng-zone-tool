@echo off
chcp 65001 > nul
cd /d "%~dp0"
set "GAODE_KEY=617e58986e6e6d1872bb9464924692b0"
echo 正在启动上城区片区分类工具，请不要关闭此窗口。
echo 浏览器打开：http://127.0.0.1:5050
"C:\Users\wy98k\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "app.py" --port 5050
pause
