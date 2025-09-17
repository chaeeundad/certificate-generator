# 교육 이수증 자동 생성 프로그램

[![Build Windows Executable](https://github.com/chaeeundad/certificate-generator/actions/workflows/build.yml/badge.svg)](https://github.com/chaeeundad/certificate-generator/actions/workflows/build.yml)

HWPX 템플릿 기반 교육 이수증 자동 생성 프로그램

## 🚀 빠른 시작

### Windows 실행 파일 다운로드
1. [Releases](https://github.com/chaeeundad/certificate-generator/releases) 페이지 방문
2. 최신 `CertificateGenerator.exe` 다운로드
3. 실행 (Python 설치 불필요)

## 📋 주요 기능

- ✅ CSV 파일에서 교육생 정보 일괄 입력
- ✅ HWPX 템플릿 기반 자동 생성
- ✅ 개별 파일 또는 병합 파일 생성
- ✅ 이수증 번호 자동 부여 (시작 번호 지정 가능)
- ✅ GUI 인터페이스

## 💻 개발 환경 설정

### 필요 사항
- Python 3.8+
- Git

### 설치
```bash
git clone https://github.com/chaeeundad/certificate-generator.git
cd certificate-generator
pip install -r requirements.txt
```

### 실행
```bash
python certificate_generator_v2.py
```

## 🏗️ 빌드

### Windows에서 EXE 빌드
```cmd
# 의존성 설치
pip install -r requirements.txt
pip install pyinstaller

# 빌드
pyinstaller certificate_generator.spec

# 또는
build_exe.bat
```

### GitHub Actions 자동 빌드
- 코드 푸시 시 자동으로 Windows 실행 파일 빌드
- Actions 탭에서 다운로드 가능

## 📝 템플릿 설정

HWPX 템플릿에 다음 placeholder 사용:

- `{교육년도}` - 현재 연도
- `{이수증번호}` - 이수증 번호
- `{교육자소속}` - 소속
- `{교육자성명}` - 이름
- `{교육자생년월일}` - 생년월일
- `{과정명}` - 교육과정명
- `{교육기간시작일}` - 시작일
- `{교육기간종료일}` - 종료일
- `{발행일자}` - 발급일

## 📊 CSV 파일 형식

```csv
이름,생년월일,소속
홍길동,1990.01.01,개발팀
김철수,1985.05.15,기획팀
```

## 🤝 기여

이슈 및 PR은 언제나 환영합니다!

## 📄 라이선스

MIT License