@echo off
chcp 65001 > nul
title 교육 이수증 생성 프로그램

echo =====================================
echo  교육 이수증 생성 프로그램 실행
echo =====================================
echo.

REM 가상환경 활성화
if exist "venv" (
    call venv\Scripts\activate.bat
    echo 가상환경을 활성화했습니다.
    echo.
) else (
    echo 가상환경이 없습니다. install_windows.bat를 먼저 실행해주세요.
    echo.
    pause
    exit /b 1
)

REM 프로그램 실행
echo 프로그램을 실행합니다...
echo.
python certificate_generator_v2.py

REM 오류 발생 시 대기
if errorlevel 1 (
    echo.
    echo 프로그램 실행 중 오류가 발생했습니다.
    pause
)