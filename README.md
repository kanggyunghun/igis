# igis

이지스자산운용 2팀 퀀트 · 포트폴리오 도구 모음. 각 프로젝트는 단말(DataGuide / FnGuide / 아이타스)
추출 데이터를 입력으로 받아, 스크리닝 · 리스크 · 수급 · 대시보드 산출물을 생성한다.

## 프로젝트 구성

| 프로젝트 | 설명 | 입력 | 출력 |
|---|---|---|---|
| `start_brief.py` | Taking Profit 스크리너 (FDR, MA10 + RVOL 시그널) | FDR(자동) | `outputs/` |
| `canaria.py` | 카나리아 리스크 스코어링 (13개 카나리 시그널 → 국면 판정) | yfinance + FRED | `canaria/날짜/` |
| `ciss/` | CISS (Composite Indicator of Systemic Stress, 시스템 스트레스 종합지표) | ECOS + FRED + yfinance | `outputs/` |
| `post_ipo/` | 상장 2년 이내 종목 모니터 | `data/*.xlsx` | `outputs/` |
| `whole_stock/` | 전종목 수급 스코어링 | `전종목_수급.xlsx` | `outputs/` |
| `under_20w/` | 20주선 하회 스크리닝 | `all_stock.xlsx` | `outputs/under_20w/` |
| `dashboard/` | MP · 블랙ON 포트폴리오 대시보드 생성 | `*.xlsx` | `dashboard/*.html` |

공통: 루트의 `env_loader.py` 가 `.env` (API 키)를 로드한다. `.env` 는 git 에서 제외되며,
`.env.example` 을 복사해 채운다.

## 폴더 구조

```
igis/
├── env_loader.py              # 공통 .env 로더
├── .env                       # API 키 (git 제외)
├── .env.example               # 키 템플릿
├── requirements.txt
│
├── start_brief.py             # Taking Profit 스크리너
├── canaria.py                 # 카나리아 리스크 모델
│
├── ciss/                      # 시스템 스트레스 종합지표
│   ├── main.py
│   ├── data_loader_v2.py
│   ├── transforms.py
│   ├── dcc_garch.py
│   └── ciss_calculator.py
│
├── post_ipo/                  # 상장 2년 이내 모니터
│   ├── run.py
│   ├── screen_ipo.py
│   ├── post_ipo_daily/
│   └── data/                  # ←입력 (수급.xlsx, univ.xlsx)
│
├── whole_stock/               # 전종목 수급 스코어링
│   ├── whole_stock.py
│   ├── fill_data.py
│   ├── fill_data.xlsx         # ←입력
│   └── 전종목_수급.xlsx        # ←입력
│
├── under_20w/                 # 20주선 하회 스크리닝
│   ├── under_20w.py
│   ├── all_stock.xlsx         # ←입력
│   └── database/
│
└── dashboard/                 # 포트폴리오 대시보드
    ├── generate_mp_dashboard.py        # 2팀 MP 관리 대시보드
    ├── generate_black_on_dashboard.py  # 블랙ON #1 대시보드
    ├── hf2_mp__날짜.xlsx                # ←MP 입력 (매일 갱신)
    ├── black_on_날짜.xlsx               # ←블랙ON 입력 (매일 갱신)
    ├── vendor/                          # 블랙ON 엑셀 다운로드용 JS (필수)
    │   ├── xlsx.full.min.js
    │   └── blackon_xlsx.js
    ├── mp_dashboard.html               # MP 출력
    └── black_on_dashboard.html         # 블랙ON 출력
```

## 설치

```bash
git clone <repo-url> igis
cd igis
pip install -r requirements.txt
cp .env.example .env      # 그리고 .env 에 API 키 입력
```

`.env` 에 필요한 키:

```
ECOS_API_KEY=             # 한국은행 경제통계시스템 (ECOS)
FRED_API_KEY=             # Federal Reserve Economic Data (FRED)
```

## 실행

```bash
# 스크리너 / 리스크
python start_brief.py
python canaria.py
python ciss/main.py
python post_ipo/run.py
python whole_stock/whole_stock.py
python under_20w/under_20w.py

# 대시보드 (dashboard 폴더 안에서)
cd dashboard
python generate_mp_dashboard.py          # 최신 hf2_mp__*.xlsx 자동 선택
python generate_black_on_dashboard.py    # 최신 black_on_*.xlsx 자동 선택
```

### dashboard 사용 메모

- 두 생성기 모두 **인자 없이 실행하면** 폴더에서 가장 최신 입력 엑셀을 자동 선택한다.
  특정 파일을 쓰려면 인자로 파일명을 준다: `python generate_mp_dashboard.py hf2_mp__20260624.xlsx`
- 출력 html 은 `dashboard/` 폴더 안에 생성된다 (`mp_dashboard.html`, `black_on_dashboard.html`).
- 생성된 html 은 단독 파일로, 브라우저로 바로 열어 사용한다 (편집 · 브라우저 저장 · 엑셀/CSV 내보내기 지원).
- **블랙ON 은 `vendor/` 폴더가 스크립트 옆에 있어야 한다.** 없으면 html 생성 단계에서
  `FileNotFoundError` 가 난다. `blackon_xlsx.js` 는 커스텀 파일이므로 저장소에 포함한다.
- `openpyxl` 의 `DrawingML support is incomplete...` 경고는 엑셀 내 도형 관련 안내로, 무시해도 된다.

## 다른 PC 에서 사용

```bash
git clone <repo-url> igis
cd igis
pip install -r requirements.txt
cp .env.example .env      # API 키 입력
```

입력 엑셀(단말 추출분)은 저장소에 포함되어 있으므로, 클론 직후 바로 실행 가능하다.
다만 매일 단말에서 받는 최신 데이터로 입력 엑셀을 갱신해야 당일 기준 결과가 나온다.