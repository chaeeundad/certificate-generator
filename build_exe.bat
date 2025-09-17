@echo off
chcp 65001 > nul
title 실행 파일 빌드

echo =====================================
echo  교육 이수증 생성 프로그램 빌드
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

REM 기존 빌드 폴더 삭제
echo [1/3] 기존 빌드 폴더 정리 중...
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"
echo.

REM PyInstaller로 빌드
echo [2/3] 실행 파일 빌드 중... (시간이 걸릴 수 있습니다)
pyinstaller certificate_generator.spec --noconfirm
echo.

REM 빌드 확인
echo [3/3] 빌드 결과 확인 중...
if exist "dist\교육이수증생성기.exe" (
    echo.
    echo =====================================
    echo  빌드가 완료되었습니다!
    echo =====================================
    echo.
    echo 실행 파일 위치: dist\교육이수증생성기.exe
    echo.

    REM 샘플 파일 복사
    if exist "교육이수증_템플릿.hwpx" (
        copy "교육이수증_템플릿.hwpx" "dist\" > nul
        echo 템플릿 파일을 복사했습니다.
    )
    if exist "sample_trainees.csv" (
        copy "sample_trainees.csv" "dist\" > nul
        echo 샘플 CSV 파일을 복사했습니다.
    )
    echo.
) else (
    echo.
    echo [오류] 빌드에 실패했습니다.
    echo.
)

pause