#!/bin/bash

# GitHub Push Script for chaeeundad
echo "GitHub 푸시 스크립트"
echo "======================"

# Git 초기화 확인
if [ ! -d ".git" ]; then
    echo "Git 저장소 초기화 중..."
    git init
fi

# 파일 추가
echo "파일 추가 중..."
git add .

# 커밋
echo "커밋 생성 중..."
git commit -m "교육 이수증 자동 생성 프로그램 v2.0"

# 리모트 설정
echo "GitHub 리모트 설정 중..."
git remote remove origin 2>/dev/null
git remote add origin https://github.com/chaeeundad/certificate-generator.git

# 브랜치 이름 확인 및 설정
branch_name=$(git rev-parse --abbrev-ref HEAD)
if [ "$branch_name" != "main" ]; then
    echo "브랜치를 main으로 변경 중..."
    git branch -M main
fi

# Push
echo "GitHub에 푸시 중..."
echo "리포지토리: https://github.com/chaeeundad/certificate-generator"
git push -u origin main

echo ""
echo "완료! 다음 단계:"
echo "1. https://github.com/chaeeundad 에서 certificate-generator 리포지토리 생성"
echo "2. Actions 탭에서 빌드 상태 확인"
echo "3. 빌드 완료 후 Artifacts에서 exe 파일 다운로드"