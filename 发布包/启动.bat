@echo off
chcp 65001 >nul
title Shopify表格生成 1.5
cd /d "%~dp0"

echo.
echo   Shopify表格生成 1.5  --  自动上新表格生成
echo   ---------------------------------------------
echo   数据表     : 请在窗口里选择供应商依赖表格 .xlsx
echo   目标文件夹 : 请在窗口里选择商品图片文件夹
echo   尺码文件夹 : 请在窗口里选择尺码表图片文件夹
echo.
echo   正在启动...
echo.

start "" "%~dp0shopify.exe"

echo.
echo   如果窗口没有弹出，请确认 shopify.exe 与本脚本、config.json、settings.json
echo   在同一个文件夹里，然后直接双击 shopify.exe。
echo.
timeout /t 3 >nul
