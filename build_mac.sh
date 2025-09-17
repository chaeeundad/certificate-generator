#!/bin/bash

echo "======================================="
echo "교육 이수증 생성기 - Mac 실행 파일 생성"
echo "======================================="
echo ""

# Python이 설치되어 있는지 확인
if ! command -v python3 &> /dev/null; then
    echo "[오류] Python 3가 설치되어 있지 않습니다."
    echo "Python을 먼저 설치해주세요."
    exit 1
fi

echo "[1단계] 필요한 패키지 설치 중..."
pip3 install pandas openpyxl lxml Pillow pyinstaller

echo ""
echo "[2단계] 실행 파일 생성 중..."
pyinstaller --onefile --windowed \
    --name="교육이수증생성기" \
    --add-data="교육이수증_템플릿.hwpx:." \
    --add-data="sample_trainees.csv:." \
    --hidden-import=tkinter \
    --hidden-import=pandas \
    --hidden-import=openpyxl \
    --hidden-import=lxml \
    --hidden-import=PIL \
    certificate_generator.py

echo ""
echo "======================================="
echo "빌드 완료!"
echo "실행 파일 위치: dist/교육이수증생성기"
echo "======================================="