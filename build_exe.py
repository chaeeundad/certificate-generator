#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
실행 파일 생성 스크립트
Windows: .exe 파일 생성
Mac: .app 파일 생성
"""

import PyInstaller.__main__
import platform
import os
import shutil

def build_executable():
    system = platform.system()

    # PyInstaller 옵션
    options = [
        'certificate_generator.py',
        '--onefile',  # 단일 실행 파일로 생성
        '--windowed',  # 콘솔 창 숨기기 (GUI만 표시)
        '--name=교육이수증생성기',
        '--add-data=교육이수증_템플릿.hwpx:.',  # 템플릿 파일 포함
        '--add-data=sample_trainees.csv:.',  # 샘플 CSV 포함
        '--hidden-import=tkinter',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=lxml',
        '--hidden-import=PIL',
    ]

    if system == 'Windows':
        options.extend([
            '--icon=icon.ico',  # 아이콘 파일 (있는 경우)
            '--version-file=version.txt',  # 버전 정보 (있는 경우)
        ])
    elif system == 'Darwin':  # macOS
        options.extend([
            '--icon=icon.icns',  # 아이콘 파일 (있는 경우)
            '--osx-bundle-identifier=com.yourcompany.certificate-generator',
        ])

    # 이전 빌드 정리
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')

    # PyInstaller 실행
    print(f"Building executable for {system}...")
    PyInstaller.__main__.run(options)

    print("\n빌드 완료!")
    print(f"실행 파일 위치: dist/교육이수증생성기{'.exe' if system == 'Windows' else ''}")

if __name__ == "__main__":
    build_executable()