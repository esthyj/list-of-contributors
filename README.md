# 헌금자 명단 (list-of-contributors)

교회 주보용 RTF 헌금 내역 파일을 읽어 A3 가로 규격의 Word 문서(.doc)로 자동 변환하는 스크립트입니다.

## ✨ 기능

- 폴더 내 `.rtf` 파일에서 헌금 내역을 파싱
- 항목별(십일조, 감사헌금, 일천번제 등)로 헌금자 이름 그룹화
- 제외 키워드(주일헌금, 특별헌금, 공과금, 구제헌금, 헌물) 자동 필터링
- 줄바꿈 등으로 합쳐진 항목명 정리 (예: `감사헌금십일조` → `십일조`)
- A3 가로, 6열 표 형태의 `.docx` 생성 후 `.doc`로 변환
- 문서 하단에 로고 이미지(`logo.png`) 삽입

## 📋 요구 사항

- Windows (Microsoft Word 설치 필수 — `.docx` → `.doc` 변환용)
- Python 3.10+
- 의존 패키지는 `requirements.txt` 참고 (`striprtf`, `python-docx`, `pywin32`)

## 🛠️ 설치

가상환경을 만들고 의존 패키지를 설치합니다.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 🚀 사용법

1. `data/` 폴더에 변환할 `.rtf` 파일을 둔다.
2. 로고가 필요하면 `data/logo.png`에 이미지를 둔다.
3. 실행:

```bash
python run.py
```

결과물은 프로젝트 루트에 `헌금자명단 (YYYY.MM.DD).doc` 파일로 저장됩니다.

폴더 구조 예시:

```
list-of-contributors/
├── run.py
├── requirements.txt
└── data/
    ├── logo.png
    └── 헌금자명단-YYYY-MM-DD.RTF
```

## 📁 파일 구성

- `run.py` — 메인 스크립트 (RTF 파싱 + Word 문서 생성)
- `requirements.txt` — Python 의존 패키지 목록
- `data/` — 입력 파일 폴더 (`.rtf` 파일과 `logo.png`를 여기에 둔다)
- `.gitignore` — 산출물(`.doc`, `.docx`, `.rtf` 등) 제외 설정

## ⚙️ 주요 설정 (run.py 상단)

| 상수 | 설명 | 기본값 |
| --- | --- | --- |
| `DATA_DIR` | 입력 데이터 폴더 | `"data"` |
| `FOLDER_PATH` | RTF 파일을 찾을 폴더 | `DATA_DIR` |
| `LOGO_FILE` | 삽입할 로고 파일 경로 | `data/logo.png` |
| `FONT_NAME` | 본문 폰트 | `"함초롬바탕"` |
| `NUM_COLUMNS` | 이름 표의 열 수 | `6` |
| `COL_WIDTH` | 열 너비 | `61.5mm` |
| `EXCLUDED_KEYWORDS` | 제외할 항목 키워드 | 주일헌금, 특별헌금, 공과금, 구제헌금, 헌물 |
| `PRIORITY_KEYS` | 출력 순서 우선 키워드 | 날짜, 십일조, 감사헌금, 일천번제 |
