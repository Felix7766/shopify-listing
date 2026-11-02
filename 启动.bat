@echo off
chcp 65001 >nul
title Shopify表格生成 1.5
cd /d "%~dp0"

set "PY=C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "%PY%" set "PY=python"

echo.
echo   Shopify表格生成 1.5  --  自动上新表格生成
echo   ---------------------------------------------
echo   数据表     : DJLLWY依赖表格.xlsx
echo   目标文件夹 : DJLLWY\DJLLWY
echo   尺码文件夹 : DJLLWY尺码表
echo.
echo   正在启动...
echo.

"%PY%" "%~dp0shopify_generator.py" %*

if errorlevel 1 (
  echo.
  echo   [启动失败] 缺少依赖时执行下面这条命令后重试:
  echo   "%PY%" -m pip install openpyxl requests pillow
  echo.
  pause
)
