@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONLEGACYWINDOWSSTDIO=utf-8

echo.
echo ========================================
echo MCP-Sheet-Parser 演示系统 (Windows)
echo ========================================
echo.

python start_demo.py

pause 