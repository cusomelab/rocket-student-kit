@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 로켓 대시보드 설치

echo.
echo  ========================================
echo   쿠팡 로켓 대시보드 - 처음 설치
echo  ========================================
echo.

set "PYEXE="
where py >nul 2>nul && set "PYEXE=py -3"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
    echo  [!] 파이썬이 없습니다.
    echo      https://www.python.org/downloads/ 에서 설치하고
    echo      설치 첫 화면의 "Add python.exe to PATH" 를 꼭 체크하세요.
    echo.
    pause
    exit /b 1
)

%PYEXE% --version
%PYEXE% -m pip install --upgrade pip
%PYEXE% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo  [!] 패키지 설치에 실패했습니다. 위 메시지를 캡처해 문의해 주세요.
    pause
    exit /b 1
)

echo.
echo  설치 완료. 이제 "2_대시보드_실행.bat" 을 더블클릭하세요.
echo.
pause
