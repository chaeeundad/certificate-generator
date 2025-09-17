@echo off
echo =======================================
echo 교육 이수증 생성기 - Windows 실행 파일 생성
echo =======================================
echo.

REM Python이 설치되어 있는지 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python을 https://www.python.org/ 에서 다운로드하여 설치해주세요.
    pause
    exit /b 1
)

echo [1단계] 필요한 패키지 설치 중...
pip install pandas openpyxl lxml Pillow pyinstaller

echo.
echo [2단계] 실행 파일 생성 중...
pyinstaller --onefile --windowed --name="교육이수증생성기" --add-data="교육이수증_템플릿.hwpx;." --add-data="sample_trainees.csv;." --hidden-import=tkinter --hidden-import=pandas --hidden-import=openpyxl --hidden-import=lxml --hidden-import=PIL certificate_generator.py

echo.
echo =======================================
echo 빌드 완료!
echo 실행 파일 위치: dist\교육이수증생성기.exe
echo =======================================
echo.
pause