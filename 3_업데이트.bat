@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 로켓 대시보드 업데이트

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py -3"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
    echo 파이썬을 찾지 못했습니다. 먼저 1_설치.bat 을 실행하세요.
    pause
    exit /b 1
)

echo 대시보드 창이 켜져 있으면 먼저 닫아 주세요.
echo.
%PYEXE% -u update.py
echo.
pause
