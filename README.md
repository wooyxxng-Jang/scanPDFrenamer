# SCAN PDF Renamer (v1.0)

스캔된 PDF 1페이지의 문서번호를 OCR로 추출하여, 엑셀(master_data.xlsx)의 제목과 매칭해 **파일명을 자동 변경**하는 Windows용 매크로 프로그램

## ✨ 주요 기능
- 좌표 기반 OCR로 문서번호 인식 (Tesseract / Poppler)
- 엑셀 매핑 후 일괄 파일명 변경
- 처리 로그/예외 처리

## 📦 실행 환경
- Windows 10/11
- Python 3.10+ (개발용)
- 배포 실행파일: Releases 탭 참고

## 🔧 빠른 시작(개발자)
```bash
git clone git clone https://github.com/wooyxxng-Jang/scanPDFrenamer.git
cd scanPDFrenamer
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
python scanPDFrenamer.py