@echo off
chcp 65001 >nul
title Document Toolkit
cd /d "%~dp0"
echo.
echo   源码运行方式（需要本机已装 Python 3.9+）
echo   如果用的是打包版，直接双击 exe，不需要这个文件
echo.
python app.py
if errorlevel 1 (
  echo.
  echo   启动失败：没有找到 python 命令，或依赖未安装
  echo   依赖安装：pip install fastapi uvicorn pywin32
)
echo.
pause
