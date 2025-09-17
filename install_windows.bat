@echo off
chcp 65001 > nul
title 교육 이수증 생성 프로그램 설치

echo =====================================
echo  교육 이수증 생성 프로그램 설치
echo =====================================
echo.

REM Python 설치 확인
echo [1/5] Python 설치 확인 중...
python --version > nul 2>&1
if errorlevel 1 (
    echo.
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python 3.8 이상을 먼저 설치해 주세요.
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Python이 설치되어 있습니다.
python --version
echo.

REM 가상환경 생성
echo [2/5] 가상환경 생성 중...
if not exist "venv" (
    python -m venv venv
    echo 가상환경이 생성되었습니다.
) else (
    echo 가상환경이 이미 존재합니다.
)
echo.

REM 가상환경 활성화
echo [3/5] 가상환경 활성화 중...
call venv\Scripts\activate.bat
echo.

REM pip 업그레이드
echo [4/5] pip 업그레이드 중...
python -m pip install --upgrade pip
echo.

REM 패키지 설치
echo [5/5] 필요한 패키지 설치 중...
pip install -r requirements.txt
echo.

echo =====================================
echo  설치가 완료되었습니다!
echo =====================================
echo.
echo 프로그램을 실행하려면 run.bat 파일을 실행하세요.
echo.
pause