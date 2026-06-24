# igis — 퀀트 스크리닝/스코어링 도구 모음

이지스자산운용 2팀 내부용. FinanceDataReader(FDR)·ECOS·FRED·데이터가이드 단말 데이터를
기반으로 한국 주식 수급/모멘텀/리스크를 스크리닝·스코어링하는 6개 프로젝트 모음.

> **private 저장소** — 데이터가이드(FnGuide) 단말 데이터가 포함되어 있으므로 외부 공개 금지.

---

## 프로젝트 구성

| 폴더/파일 | 설명 | 입력 | 출력 |
|---|---|---|---|
| `screener_brief.py` | 익절 스크리너 (MA10 + RVOL 신호) | FDR | `outputs/` |
| `canaria/risk_scoring_model.py` | 카나리아 리스크 스코어 (13개 시그널 → 국면) | yfinance + FRED | `outputs/canaria/날짜/` |
| `ciss/` | CISS 시스템적 스트레스 지표 (ECB식 5개 섹터) | ECOS + FRED + yfinance | `outputs/ciss/` |
| `post_ipo/` | 2년 이내 신규상장 모니터링 | `data/수급.xlsx`, 유니버스 엑셀 | `outputs/post_ipo/` |
| `whole_stock/` | 전종목 수급 스코어링 (update→stack 누적) | `전종목_수급.xlsx` | `outputs/whole_stock/` |
| `under_20w/` | 20주선 하회 종목 스크리닝 | `all_stock.xlsx` | `outputs/under_20w/` |

공통: 루트의 `env_loader.py`(의존성 없는 .env 로더), `.env`(API 키, 깃 제외).

---

## 다른 PC에서 셋업 (클론 후 바로 실행)

### 1. 클론
```bash
git clone <이 저장소 URL>
cd igis
```

### 2. 파이썬 환경 (3.12 권장)
```bash
# 가상환경 (선택)
python -m venv .venv
# Windows
.venv\Scripts\activate
# Mac/Linux
source .venv/bin/activate
```

### 3. 패키지 설치
```bash
pip install -r requirements.txt
```

### 4. API 키 설정 (canaria, ciss 만 필요)
```bash
# Windows
copy .env.example .env
# Mac/Linux
cp .env.example .env
```
그 후 `.env` 를 열어 ECOS_API_KEY, FRED_API_KEY 에 실제 값 입력.
(screener_brief, post_ipo, whole_stock, under_20w 는 FDR 무료라 키 불필요)

### 5. 실행 예시
```bash
# 익절 스크리너
py -3.12 screener_brief.py

# 카나리아 리스크
py -3.12 canaria/risk_scoring_model.py

# CISS
cd ciss && py -3.12 main.py

# Post IPO (기본 버전 B)
cd post_ipo && py -3.12 run.py        # 버전 A는 --a

# 전종목 수급 스코어링 (엑셀 닫고 실행)
cd whole_stock && py -3.12 whole_stock.py

# 20주선 하회 스크리닝
cd under_20w && py -3.12 under_20w.py
```

---

## 데이터 파일 안내

데이터가이드 단말에서 받는 엑셀들은 매일 갱신해야 최신 결과가 나옵니다.

- `post_ipo/data/수급.xlsx` — Refresh 후 저장
- `whole_stock/전종목_수급.xlsx` — update 시트 Refresh 후 저장 (**엑셀 닫고** 실행)
- `under_20w/all_stock.xlsx` — Refresh 후 저장 (시가총액 단위: 억원)

### whole_stock 거래일 공백 보강
빠진 거래일이 있으면 `whole_stock.py` 가 중단하고 누락 날짜를 출력합니다.
`fill_data.xlsx`(시트명 = YYYYMMDD, 각 시트는 update 시트와 동일 구성)를 만들어
`py -3.12 fill_data.py` 실행 → 보강 후 다시 `whole_stock.py` 실행.

---

## 주의

- `.env`(API 키)는 깃에 올라가지 않습니다. 새 PC에서는 `.env.example` 을 복사해 채우세요.
- 엑셀을 Excel 에서 열어둔 채 `whole_stock.py` 를 돌리면 stack 시트 기록이 실패합니다(별도 누적파일엔 저장됨).
- 캐시 파일(`.krx_*.json`, `under_20w/database/`)은 자동 생성되므로 지워도 무방합니다.
