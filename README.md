# 💎 절삭평가 쉽게보기 (S-Reposer) v1.0.0

[![Release](https://img.shields.io/github/v/release/shin9602/Easy-Cutting-Report?color=00ffc3&logo=github)](https://github.com/shin9602/Easy-Cutting-Report/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **복잡한 절삭평가 리포트 작업을 단 5분 만에.**  
엑셀에 박힌 고화질 원본 이미지를 자동으로 추출하고, 세련된 웹 인터페이스에서 자유롭게 레이아웃을 조립하세요.

---

## 🚀 주요 기능 (Key Features)

### 1. 초고화질 원본 이미지 자동 추출 (Auto Extractor)
- `.xls`, `.xlsx` 파일 내에 삽입된 소스 이미지를 **해상도 손실 없이 100% 원본**으로 추출합니다.
- 인서트 **상면(Top)**과 **측면(Side)**을 열 위치와 이미지 크기 기반으로 스마트하게 자동 분류하여 폴더별로 정리해줍니다.

### 2. 인터렉티브 리포트 빌더 (Web-based Builder)
- 드래그 앤 드롭으로 샘플의 순서와 배치를 자유롭게 조절합니다.
- 현장에서 즉시 확인 가능한 **실시간 미리보기**와 **고해상도 PNG/PDF 내보내기** 기능을 지원합니다.

### 3. 무설치 단일 실행 파일 (Single Portable EXE)
- 별도의 파이썬 설치나 복잡한 셋업이 필요 없습니다.
- `Start.bat` 실행 시 모든 환경이 자동 구축되며, 생성된 `.exe` 파일 하나만 가지고 어디서든 작업할 수 있습니다.

### 4. 스마트 자동 업데이트 (Auto Update)
- 프로그램 실행 시 GitHub의 최신 배포 버전을 체크합니다.
- 새로운 기능이 추가되면 클릭 한 번으로 자동 업데이트가 진행됩니다.

---

## 🛠️ 시작하기 (Quick Start)

### 사용 방법
1. 본 레포지토리를 **Download ZIP** 하거나 `git clone` 하세요.
2. 폴더 내의 **`Start.bat`** 파일을 실행합니다.
3. 최초 실행 시 필요한 환경이 자동 구성되며 웹 브라우저가 열립니다.
4. 이후 생성된 **`CuttingEval_App.exe`** 파일만 따로 복사하여 사용하셔도 됩니다.

---

## 📦 기술 스택 (Tech Stack)

- **Backend**: Python (Flask), OpenPyXL, Pillow, PyInstaller
- **Frontend**: Vanilla HTML5, CSS3 (Modern Glassmorphic Design), JavaScript (ES6+), html2canvas
- **Automation**: GitHub Actions (Auto-Release)

---

## ⚠️ 주의 사항
- **엑셀 호환성**: `.xls` 파일 처리 시 로컬에 Microsoft Excel이 설치되어 있으면 더 정확한 변환이 가능합니다. (미설치 시 오픈소스 엔진을 사용하여 대체 처리합니다.)
- **업데이트 설정**: `Start.bat` 내의 `REPO` 변수를 본인의 GitHub 주소에 맞게 수정해야 자동 업데이트가 작동합니다.

---

## 🤝 기여하기
버그 보고나 기능 제안은 **Issues** 탭을 이용해 주세요.

## 📄 라이선스
Copyright © 2024. MIT License.
