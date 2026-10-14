@echo off
chcp 65001 >nul
title 城市中国·湖北通城网 综合数据面板
cd /d "%~dp0"

echo ================================================================
echo   城市中国·湖北通城网 招聘频道 综合数据面板
echo   首页 http://127.0.0.1:8795  （平台数据 /city/  企业投递 /jd/）
echo ================================================================
echo.

rem ---- 优先使用系统 Python 3.14（已装 openpyxl / rapidocr 等依赖）----
set "PY=C:\Python314\python.exe"
if not exist "%PY%" set "PY=python"

echo 使用解释器: %PY%
echo 正在启动服务，请稍候...
echo.

"%PY%" -X utf8 "面板.py"

echo.
echo 服务已退出。按任意键关闭窗口。
pause >nul
