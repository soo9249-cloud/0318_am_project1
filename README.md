# 명찰 자동 생성기 (Badge Maker)

엑셀 파일에서 인원 정보를 읽어 명찰(배지/랜야드)을 자동으로 PPT로 생성하는 웹 대시보드 툴.

## 기술 스택

| 라이브러리 | 용도 |
|---|---|
| FastAPI | 웹 서버 (REST API + HTML 서빙) |
| Jinja2 | HTML 템플릿 렌더링 |
| python-multipart | Form/파일 업로드 파싱 |
| python-pptx | PPT 슬라이드 생성 (도형/텍스트 직접 렌더링) |
| Pillow | 로고 원형 크롭 (텍스트 렌더링에는 미사용) |
| openpyxl | 엑셀 읽기 |
| lxml | PPT XML 조작 (점선, 그라데이션, 폰트 embed) |

## 폴더 구조

```
0318_am_project1/
  ├── main.py               # FastAPI 앱 진입점
  ├── badge_maker.py        # 핵심 로직 (엑셀 읽기, 배지 생성)
  ├── font_embed.py         # OOXML 폰트 embed (ECMA-376 §22.5.2)
  ├── download_fonts.py     # Google Fonts 자동 다운로드 스크립트
  ├── templates/
  │   └── index.html        # 메인 UI (사이드바 레이아웃, Lucide 아이콘)
  ├── designs/
  │   ├── __init__.py       # Python 패키지 선언
  │   ├── badge.py          # 배지 디자인 A/B/C (python-pptx 네이티브)
  │   ├── lanyard.py        # 랜야드 디자인 D/E/F (python-pptx 네이티브)
  │   └── design_spec.json  # 디자인 스펙 (레이아웃, 색상, 폰트)
  ├── fonts/
  │   ├── README.txt        # 폰트 다운로드 안내
  │   ├── Montserrat-Regular.ttf
  │   ├── Montserrat-SemiBold.ttf
  │   ├── Montserrat-Bold.ttf
  │   ├── NotoSerifKR-Bold.ttf
  │   └── NotoSansKR-Regular.ttf
  ├── output/               # 생성된 PPT 저장 경로 (.gitignore 적용)
  ├── requirements.txt
  ├── .gitignore
  └── Dockerfile
```

## 디자인 종류

### 배지 (Badge) — 소형 명찰

| ID | 이름 | 레이아웃 | 색상 옵션 |
|---|---|---|---|
| A | Classic | 로고 좌측(원형) + 세로 구분선 + 텍스트 우측 | Gold, RoseGold, Silver, Black |
| B | Centered | 전체 중앙 정렬 + 가로 구분선 | Gold, RoseGold, Silver, Black |
| C | Modern | 미니멀 좌측 로고 + 우측 텍스트 | Gold, RoseGold, Silver, Black |

### 랜야드 (Lanyard) — 목걸이형 대형 명찰

| ID | 이름 | 배경 | 색상 옵션 |
|---|---|---|---|
| D | Wave Corporate | 흰 배경 + 하단 웨이브 도형 | Coral, Mint, Sky, Lavender |
| E | Gradient Vivid | Pillow 그라데이션 이미지 | CoralPurple, MintBlue, SkyIndigo, PeachGold |
| F | Bold Solid | 단색 밝은 배경 + 흰 상단 바 | SkyBlue, Coral, Mint, Lavender |

## 설치 및 실행

### 로컬 실행

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. 폰트 파일 준비 (fonts/README.txt 참고)

# 3. 앱 실행 (방법 1 — 개발 모드, 코드 변경 시 자동 재시작)
uvicorn main:app --reload --port 8080
# → http://127.0.0.1:8080 접속

# 3. 앱 실행 (방법 2 — 직접 실행)
python main.py
# → http://127.0.0.1:8080 접속
```

### Docker 실행

```bash
docker build -t badge-maker .
docker run -p 8080:8080 badge-maker
```

## 폰트 준비

`download_fonts.py`를 실행하면 Google Fonts에서 자동으로 다운로드됩니다.

```bash
python download_fonts.py
```

수동으로 넣으려면 `fonts/` 폴더에 아래 파일을 [Google Fonts](https://fonts.google.com)에서 다운로드하세요.

- `NotoSerifKR-Bold.ttf`
- `NotoSansKR-Regular.ttf`
- `Montserrat-Bold.ttf`
- `Montserrat-Regular.ttf`
- `Montserrat-SemiBold.ttf`

> 폰트 파일이 없으면 시스템 한글 폰트(맑은 고딕 등)로 자동 fallback됩니다.

## 엑셀 입력 형식

| 컬럼명 | 설명 | 예시 |
|---|---|---|
| 이름(한글) | 한글 이름 | 홍길동 |
| 이름(영문) | 영문 이름 | HONG GILDONG |
| 직급 | 직급/직책 | 부장 |
| 부서 | 소속 부서 | 영업팀 |


## 구현 현황

| 파일 | 상태 | 내용 |
|---|---|---|
| `badge_maker.py` `read_excel` | ✅ | 엑셀 파싱 |
| `badge_maker.py` `generate_badges` | ✅ | PPT 생성 및 저장 (A~F 전 디자인 검증 완료) |
| `designs/badge.py` | ✅ | 디자인 A/B/C 렌더링 + 슬라이드 배치 |
| `designs/lanyard.py` | ✅ | 디자인 D/E/F 렌더링 + 슬라이드 배치 |
| `download_fonts.py` | ✅ | GitHub raw에서 Google Fonts 자동 다운로드 |

## 렌더링 방식

**python-pptx 네이티브 도형/텍스트** 방식 사용
- 텍스트는 벡터(편집 가능) — 래스터 이미지 아님
- 폰트는 PPT 파일 내 embed → 다른 컴퓨터에서 열어도 폰트 유지
- 둥근 모서리: ROUNDED_RECTANGLE 도형 (adjustments)
- 원형 로고: Pillow 원형 크롭 → PNG → add_picture (로고만 이미지)
- Design E 그라데이션: lxml XML 직접 조작 (gradFill)
- 로고 없으면 디자인별 fallback (A: 회사명 텍스트, B: 회사명, C: 컬러 원)
- 직급 없으면 부서로 자동 대체 (빈 여백 없음)
- 폰트 embed: ECMA-376 §22.5.2 GUID XOR 난독화 적용

## 주요 함수 (`badge_maker.py`)

### `read_excel(filepath)` ✅ 구현 완료
```python
read_excel(filepath: str) -> list[dict]
# 반환 예시:
# [{"name_kr": "김민준", "name_en": "Kim Minjun",
#   "dept": "마케팅팀", "rank": "팀장", "company": "회사명"}]
# 오류 시: [{"error": "메시지"}]
```
- 헤더 매핑: `이름` → `name_kr`, `영문이름` → `name_en`, `부서` → `dept`, `직급` → `rank`, `회사명` → `company`
- 빈 행 자동 스킵, 셀 값 strip() 처리

### `designs/badge.py` 공개 API ✅

```python
render_badge_image(person, design, color, size_key, logo_path, company_name) -> PIL.Image
# 명찰 1개를 Pillow 이미지로 렌더링

place_on_slide(slide, images, layout, badge_size_mm)
# A4 슬라이드에 이미지 배치 + 점선 컷 가이드 추가
```

### `generate_badges(...)` ✅ 구현 완료
```python
generate_badges(people, logo_path, company_name,
                badge_type, design, color, size,
                layout, event_mode, event_info,
                output_dir) -> str
# 배지/랜야드 PPT 파일 생성, 저장 경로 반환
```

## UI 구성 (`templates/index.html`) ✅ 구현 완료

순수 HTML/CSS/JS 싱글 페이지 앱 (프레임워크 없음). FastAPI가 Jinja2로 `design_spec.json` 스펙 데이터를 주입.

### 레이아웃

- **사이드바 (72px)**: Lucide 아이콘 3개 (파일 업로드 / 디자인 설정 / 생성 & 다운로드)
- **메인 영역**: 탭 전환 패널 (max-width: 1000px)

### 탭 패널

| 탭 | 내용 |
|---|---|
| 파일 업로드 | 엑셀 드래그&드롭 업로드 (미리보기 최대 5행) + 로고 업로드 / 회사명 입력 |
| 디자인 설정 | 배지형 / 랜야드형 전환 · 규격·배치·색상 토글 · 디자인 카드 선택 (A/B/C 또는 D/E/F) · 랜야드 행사 모드 |
| 생성 & 다운로드 | 설정 요약 확인 · 명찰 생성 버튼 · PPT 다운로드 |

### API 엔드포인트

| 메서드 | 경로 | 설명 |
|---|---|---|
| GET | `/` | 메인 UI 페이지 |
| POST | `/upload-excel` | 엑셀 파일 업로드 → 인원 목록 반환 |
| POST | `/upload-logo` | 로고 이미지 업로드 → base64 미리보기 반환 |
| POST | `/generate` | 명찰 PPT 생성 |
| GET | `/download/{filename}` | 생성된 PPT 다운로드 |
