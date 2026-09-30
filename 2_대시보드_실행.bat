@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 쿠팡 로켓 대시보드 :5080 - 이 창 닫지 마세요

netstat -ano | findstr ":5080" | findstr "LISTENING" >nul
if %errorlevel%==0 (
    echo 이미 실행 중입니다. 브라우저를 엽니다.
    start "" http://127.0.0.1:5080
    timeout /t 2 /nobreak >nul
    exit /b
)

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

echo.
echo  ========================================
echo   쿠팡 로켓 대시보드
echo   브라우저: http://127.0.0.1:5080
echo  ========================================
echo   이 창을 닫으면 대시보드가 꺼집니다.
echo   작업 중에는 쿠팡 로그인용 Chrome 창이 따로 뜹니다.
echo   (프로그램은 로그인 버튼을 누르지 않습니다)
echo.
%PYEXE% -u dashboard.py
echo.
pause
