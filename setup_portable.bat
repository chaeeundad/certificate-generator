@echo off
chcp 65001 > nul
title 휴대용 Python 환경 설정

echo =====================================
echo  휴대용 Python 환경 설정
echo =====================================
echo.
echo 이 스크립트는 Python을 포함한 완전한 휴대용 환경을 만듭니다.
echo.

REM 작업 디렉토리 생성
if not exist "portable" mkdir portable
cd portable

REM Python Embeddable Package 다운로드 (Windows x64)
echo [1/4] Python 휴대용 패키지 다운로드 중...
echo.
echo Python 3.11.5 Embeddable Package를 다운로드합니다.
echo 수동 다운로드: https://www.python.org/ftp/python/3.11.5/python-3.11.5-embed-amd64.zip
echo.

if not exist "python-3.11.5-embed-amd64.zip" (
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.5/python-3.11.5-embed-amd64.zip' -OutFile 'python-3.11.5-embed-amd64.zip'"
)

REM Python 압축 해제
echo [2/4] Python 압축 해제 중...
if not exist "python" mkdir python
powershell -Command "Expand-Archive -Path 'python-3.11.5-embed-amd64.zip' -DestinationPath 'python' -Force"

REM get-pip.py 다운로드
echo [3/4] pip 설치 준비 중...
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py'"

REM pip 설치
echo [4/4] pip 설치 중...
python\python.exe get-pip.py

REM python311._pth 파일 수정 (pip 사용 가능하도록)
echo import site >> python\python311._pth

REM 필요한 패키지 설치
echo.
echo 필요한 패키지 설치 중...
python\python.exe -m pip install pandas openpyxl pyinstaller

REM 실행 스크립트 생성
echo @echo off > ..\run_portable.bat
echo cd portable >> ..\run_portable.bat
echo python\python.exe ..\certificate_generator_v2.py >> ..\run_portable.bat
echo pause >> ..\run_portable.bat

REM 빌드 스크립트 생성
echo @echo off > ..\build_portable.bat
echo cd portable >> ..\build_portable.bat
echo python\python.exe -m PyInstaller ..\certificate_generator.spec --noconfirm >> ..\build_portable.bat
echo pause >> ..\build_portable.bat

cd ..

echo.
echo =====================================
echo  설정 완료!
echo =====================================
echo.
echo 휴대용 Python 환경이 생성되었습니다.
echo.
echo 실행 방법:
echo   - run_portable.bat 실행
echo.
echo EXE 빌드:
echo   - build_portable.bat 실행
echo.
echo 이 폴더를 USB나 다른 컴퓨터로 복사해도 동작합니다.
echo.
pause