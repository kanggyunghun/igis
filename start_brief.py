"""
Taking Profit Screener (단일 파일 버전)
=========================================================================
FinanceDataReader(FDR) 기반 10일선 이탈 + 거래량 폭증 익절 신호 스크리너.

기존 start_brief.py + src/analyzer.py + src/screener.py 를 하나로 통합하고
Bloomberg 잔재 / CSV·시간봉 변환 / crossover for-loop / 히트맵 등 죽은 코드와
비효율을 모두 제거한 재작성 버전. 출력은 콘솔 + Excel.

티커 입력은 '엑셀에서 한 줄로 복붙한 덩어리'까지 처리한다.
  예) "000660<탭>SK하이닉스087010<탭>펩트론005930<탭>삼성전자..."
  → 종목명 끝에 다음 코드가 공백 없이 붙어도, KRX 상장코드 마스터와 대조해
    진짜 코드만 추출하고 중복을 제거한다.

실행:  py -3.12 screener_brief.py
=========================================================================
"""
import os
import re
import sys
import json
from datetime import datetime, timedelta, time, date
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import pandas as pd
from tabulate import tabulate

import FinanceDataReader as fdr

# Windows 콘솔 UTF-8
if sys.platform == 'win32':
    import codecs
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

# =========================================================================
# 설정값 (한곳에 모음)
# =========================================================================
MA_PERIOD = 10          # 이동평균 기간
RVOL_PERIOD = 10        # 상대거래량 평균 기간
RVOL_SURGE = 1.5        # 거래량 폭증 임계값 (Condition 2)  ← 1.5로 통일
RVOL_WATCH = 1.0        # 거래량 관심 하한 (Condition 3)
RSI_PERIOD = 14         # RSI 기간
PERIOD_DAYS = 120       # FDR 다운로드 기간(캘린더일). 3개월 + 영업일 버퍼
MAX_WORKERS = 10        # 병렬 다운로드 스레드 수
MARKET_CLOSE = time(15, 30)

# Excel 노이즈 필터 (profit taking / 주의필요 에만 적용)
EXCEL_RECENT_DAYS = 5   # 이탈일이 최근 N일 이내
EXCEL_MIN_GAP = 3       # 돌파~이탈 간격 최소 일수

_CACHE_FILE = os.path.join(CURRENT_DIR, '.krx_name_cache.json')
_MASTER_FILE = os.path.join(CURRENT_DIR, '.krx_code_master.json')  # 코드→(이름,시장) 마스터 캐시
_MASTER_TTL_DAYS = 3    # 마스터 캐시 유효기간(일). 지나면 재다운로드


# =========================================================================
# 1. KRX 코드 마스터 (상장 종목 코드 ↔ 이름/시장)
#    엑셀 복붙처럼 '코드+이름'이 붙어버린 입력에서 진짜 코드만 가려내는 기준.
# =========================================================================
def _load_master_cache() -> dict:
    if not os.path.exists(_MASTER_FILE):
        return {}
    try:
        with open(_MASTER_FILE, 'r', encoding='utf-8') as f:
            blob = json.load(f)
        ts = datetime.fromisoformat(blob.get('_ts', '2000-01-01'))
        if datetime.now() - ts > timedelta(days=_MASTER_TTL_DAYS):
            return {}                       # 만료
        return blob.get('codes', {})
    except Exception:
        return {}


def _save_master_cache(codes: dict) -> None:
    try:
        with open(_MASTER_FILE, 'w', encoding='utf-8') as f:
            json.dump({'_ts': datetime.now().isoformat(), 'codes': codes}, f, ensure_ascii=False)
    except Exception:
        pass


def build_code_master(verbose: bool = True) -> dict:
    """
    KRX 전체 상장 코드 마스터 구축 → {code: {'name':.., 'market':'KS'/'KQ'}}.
    KRX(주식) + KOSPI + KOSDAQ + ETF/KR 를 합치고, 디스크에 캐시(_MASTER_TTL_DAYS).
    """
    cached = _load_master_cache()
    if cached:
        return cached

    if verbose:
        print("[코드 마스터 다운로드 중... 최초 1회만, 이후 캐시]")
    codes = {}
    for src in ['KRX', 'KOSPI', 'KOSDAQ', 'ETF/KR']:
        try:
            df = fdr.StockListing(src)
            cc = 'Code' if 'Code' in df.columns else ('Symbol' if 'Symbol' in df.columns else None)
            if cc is None:
                continue
            for _, row in df.iterrows():
                code = str(row[cc]).strip().upper()
                if not code or code in codes:
                    continue
                name = str(row.get('Name', '')).strip()
                mkt_raw = str(row.get('Market', '')).strip().upper()
                # ETF/KR 등은 Market 비어있을 수 있음 → KS 기본
                suffix = 'KQ' if 'KOSDAQ' in mkt_raw else 'KS'
                codes[code] = {'name': name, 'market': suffix}
        except Exception as e:
            if verbose:
                print(f"  ({src} 스킵: {str(e)[:50]})")

    if codes:
        _save_master_cache(codes)
    if verbose:
        print(f"[마스터 코드 수: {len(codes)}]")
    return codes


# =========================================================================
# 2. 티커 입력 파서 (마스터 대조 + 신형코드 FDR fallback)
# =========================================================================
_CODE = re.compile(r'[0-9][0-9A-Z]{5}')     # 문자열 어디에 박혀 있든 코드 후보 추출
_US_TICKER = re.compile(r'^[A-Z]{1,5}$')


def _fdr_alive(code: str) -> bool:
    """마스터에 없는 코드(신형 ETF/ETN 등): FDR로 실제 데이터 받아보고 살아있으면 True."""
    try:
        end = datetime.now()
        start = end - timedelta(days=20)
        df = fdr.DataReader(code, start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d'))
        return df is not None and len(df) > 0
    except Exception:
        return False


def parse_ticker_input(raw: str, master: dict = None, verify_unknown: bool = True) -> tuple:
    """
    복붙 덩어리/쉼표입력 → (tickers, names). 중복 제거, 입력 순서 유지.

    처리 순서
    1) 문자열 전체에서 6자리 코드 후보를 모두 추출 (이름에 붙어 있어도 OK)
    2) 마스터에 있으면 채택 + 이름/시장(KS/KQ) 확정
    3) 마스터에 없으면:
       - verify_unknown=True : FDR로 실제 데이터 확인 → 살아있으면 채택(가짜 제거)
       - verify_unknown=False: 일단 채택(데이터 단계에서 자연 탈락)
    4) 미국 티커(영문 1~5자)는 그대로 US 처리
    """
    if master is None:
        master = {}

    s = raw.replace('\t', ' ').replace('\n', ' ').replace(',', ' ')

    # 코드 후보 (등장 순서, 중복 제거)
    seen_c, candidates = set(), []
    for m in _CODE.finditer(s):
        c = m.group()
        if c not in seen_c:
            seen_c.add(c)
            candidates.append(c)

    # 미국 티커 후보: 공백 토큰 중 순수 영문 1~5자
    us_tokens = [t.upper() for t in s.split() if _US_TICKER.match(t.upper())]

    tickers, names, rejected = [], {}, []
    for code in candidates:
        if code in master:
            info = master[code]
            full = f"{code} {info['market']}"
            tickers.append(full)
            names[full] = info['name'] or code
        else:
            if verify_unknown and not _fdr_alive(code):
                rejected.append(code)
                continue
            full = f"{code} KS"          # 마스터에 없지만 살아있는 신형코드
            tickers.append(full)
            names[full] = code

    # 미국 티커 추가
    us_seen = set()
    for tok in us_tokens:
        if tok in us_seen:
            continue
        us_seen.add(tok)
        full = f"{tok} US"
        if full not in names:
            tickers.append(full)
            names[full] = tok

    if rejected:
        print(f"[제외] 마스터에 없고 데이터도 없는 코드 {len(rejected)}개: {rejected}")

    return tickers, names


def parse_ticker(ticker: str) -> tuple:
    """'005930 KS' → ('005930', 'KS')"""
    parts = ticker.strip().split()
    if len(parts) >= 2:
        return parts[0].upper(), parts[1].upper()
    return parts[0].upper(), None


# =========================================================================
# 3. 종목명 캐시 (마스터에서 못 채운 종목만 pykrx 보강)
# =========================================================================
def _load_cache() -> dict:
    if not os.path.exists(_CACHE_FILE):
        return {}
    try:
        with open(_CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cache(cache: dict) -> None:
    try:
        with open(_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass


def resolve_kr_name(code: str, cache: dict) -> str:
    """한국 종목코드 → 종목명 (캐시 → pykrx 단건)."""
    code = code.strip().upper()
    if code in cache and cache[code]:
        return cache[code]
    try:
        from pykrx import stock as pykrx_stock
        nm = pykrx_stock.get_market_ticker_name(code)
        if nm and nm.strip():
            cache[code] = nm.strip()
            return nm.strip()
    except Exception:
        pass
    return code


def fill_missing_names(tickers: list, names: dict) -> dict:
    """이름이 비었거나 코드와 같은 종목만 캐시/pykrx로 보강. 미국 종목은 티커 그대로."""
    cache = _load_cache()
    for t in tickers:
        cur = names.get(t)
        code, market = parse_ticker(t)
        if cur and cur != code:          # 이미 이름이 있으면 패스
            continue
        if market in ('KS', 'KQ'):
            names[t] = resolve_kr_name(code, cache)
        else:
            names[t] = code
    _save_cache(cache)
    return names


# =========================================================================
# 4. 데이터 다운로드 (FDR)
# =========================================================================
def download_data(ticker: str, verbose: bool = False) -> pd.DataFrame:
    """FDR로 일봉 OHLCV 다운로드 → [Date, Open, High, Low, Close, Volume]."""
    code, market = parse_ticker(ticker)
    if market not in ('KS', 'KQ', 'US'):
        if verbose:
            print(f"  지원하지 않는 시장: {market}")
        return None

    start = (datetime.now() - timedelta(days=PERIOD_DAYS)).strftime('%Y-%m-%d')
    end = datetime.now().strftime('%Y-%m-%d')
    try:
        df = fdr.DataReader(code, start, end)   # 시장 접미사 없이 코드만 사용
    except Exception as e:
        if verbose:
            print(f"  다운로드 실패 {ticker}: {e}")
        return None
    if df is None or len(df) == 0:
        return None

    df = df.reset_index()
    for cand in ('Date', 'date', 'index'):
        if cand in df.columns:
            if cand != 'Date':
                df = df.rename(columns={cand: 'Date'})
            break

    required = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
    if any(c not in df.columns for c in required):
        if verbose:
            print(f"  컬럼 부족 {ticker}: {list(df.columns)}")
        return None

    df['Date'] = pd.to_datetime(df['Date'])
    if getattr(df['Date'].dt, 'tz', None) is not None:
        df['Date'] = df['Date'].dt.tz_localize(None)
    return df[required].copy()


# =========================================================================
# 5. 지표 계산 + 신호 생성 (구 screener.ExitSignalScreener)
# =========================================================================
def compute_indicators(df: pd.DataFrame) -> pd.DataFrame:
    """
    OHLCV → 지표/신호가 추가된 DataFrame.
    - MA10, RVOL: shift(1)로 '당일 제외' (오늘 종가 vs 어제까지의 평균)
    - 10일선 돌파/이탈일: 벡터화(mask + ffill, for-loop 제거)
    - Signal: BUY / WATCH / SELL / HOLD
    """
    d = df.copy()
    close, vol = d['Close'], d['Volume']

    d['MA10'] = close.shift(1).rolling(MA_PERIOD).mean()
    d['RVOL'] = vol / vol.shift(1).rolling(RVOL_PERIOD).mean()

    # RSI (Wilder smoothing)
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = (-delta).clip(lower=0)
    avg_gain = gain.ewm(alpha=1 / RSI_PERIOD, min_periods=RSI_PERIOD).mean()
    avg_loss = loss.ewm(alpha=1 / RSI_PERIOD, min_periods=RSI_PERIOD).mean()
    rs = avg_gain / avg_loss
    d['RSI'] = 100 - (100 / (1 + rs))

    # 10일선 돌파/이탈 추적 (벡터화)
    dates = d['Date'].dt.strftime('%Y-%m-%d')
    position = (close >= d['MA10']).astype('Int64')      # 1=위, 0=아래, MA NaN이면 <NA>
    prev = position.shift(1)
    break_below = (prev == 1) & (position == 0)          # 위→아래 = 이탈
    break_above = (prev == 0) & (position == 1)          # 아래→위 = 돌파

    d['Last_MA10_Break_Below'] = dates.where(break_below).ffill()
    d['Last_MA10_Break_Above'] = dates.where(break_above).ffill()

    below = (position == 0)
    grp = (position != position.shift(1)).cumsum()
    d['Days_Below_MA10'] = (below.groupby(grp).cumsum()).where(below, 0).fillna(0).astype(int)

    # 조건 + 신호
    c1 = close < d['MA10']                                # 10일선 아래
    c2 = d['RVOL'] >= RVOL_SURGE                          # 거래량 폭증
    c3 = (d['RVOL'] >= RVOL_WATCH) & (d['RVOL'] < RVOL_SURGE)
    d['Condition_1_Trend_Breakdown'] = c1
    d['Condition_2_Volume_Confirmation'] = c2

    d['Signal'] = np.select(
        [(~c1) & c2, c3, c1 & c2],
        ['BUY', 'WATCH', 'SELL'],
        default='HOLD'
    )
    d['Reasoning'] = np.select(
        [d['Signal'] == 'BUY', d['Signal'] == 'WATCH', d['Signal'] == 'SELL', c1],
        ['10일선 위 + 거래량 폭증 (매수)',
         '거래량 증가 (관심)',
         '10일선 하회 + 거래량 폭증 (매도)',
         '10일선 아래, 거래량 부족'],
        default='10일선 위, 정상 거래량'
    )
    return d


def analyze_ticker(ticker: str, mode: int) -> dict:
    """종목 1개 분석 → 결과 dict (구 analyzer.analyze_latest 역할)."""
    df = download_data(ticker)
    if df is None or len(df) == 0:
        return None

    if mode == 1 and datetime.now().time() < MARKET_CLOSE:
        today = datetime.now().date()
        df = df[df['Date'].dt.date != today]
    if len(df) < 2:
        return None

    d = compute_indicators(df)
    last = d.iloc[-1]
    prev_close = d.iloc[-2]['Close']

    ma_dist_pct = ((last['Close'] - last['MA10']) / last['MA10']) * 100 if pd.notna(last['MA10']) else np.nan
    pos = 'below' if last['Close'] < last['MA10'] else 'above'
    ba, bb = last['Last_MA10_Break_Above'], last['Last_MA10_Break_Below']
    bb_s = bb if pd.notna(bb) else '?'
    ba_s = ba if pd.notna(ba) else '?'
    if pos == 'below':
        trend_dir, trend_detail = '하락세', f"10일선 위({ba_s}) → 10일선 아래({bb_s})"
    else:
        trend_dir, trend_detail = '상승세', f"10일선 아래({bb_s}) → 10일선 위({ba_s})"

    return {
        'ticker': ticker,
        'close_price': last['Close'],
        'prev_close': prev_close,
        'price_change_percent': (last['Close'] - prev_close) / prev_close * 100,
        'ma10': last['MA10'],
        'ma_distance_percent': ma_dist_pct,
        'current_position': pos,
        'trend_direction': trend_dir,
        'trend_detail': trend_detail,
        'last_ma10_break_below': last['Last_MA10_Break_Below'],
        'last_ma10_break_above': last['Last_MA10_Break_Above'],
        'rvol': last['RVOL'],
        'rsi': float(last['RSI']) if pd.notna(last['RSI']) else None,
        'signal': last['Signal'],
        'condition_1_trend_breakdown': bool(last['Condition_1_Trend_Breakdown']),
        'condition_2_volume_confirmation': bool(last['Condition_2_Volume_Confirmation']),
    }


# =========================================================================
# 6. 출력 — 콘솔
# =========================================================================
def print_console(results: list, names: dict) -> pd.DataFrame:
    df = pd.DataFrame(results)

    print("\n" + "=" * 80)
    print("분석 결과 요약")
    print("=" * 80)
    rows = []
    for r in results:
        chg = f"{r['price_change_percent']:+.1f}%" if r.get('prev_close') else "-"
        rows.append({
            '종목': names.get(r['ticker'], r['ticker']),
            '현재가': f"{r['close_price']:.0f}",
            '전일비': chg,
            '10일선': f"{r['ma10']:.0f}" if pd.notna(r['ma10']) else "-",
            '괴리율': f"{r['ma_distance_percent']:+.1f}%" if pd.notna(r['ma_distance_percent']) else "-",
            '추세': r['trend_direction'],
            'RVOL': f"{r['rvol']:.1f}배" if pd.notna(r['rvol']) else "-",
            '신호': r['signal'],
        })
    print("\n" + tabulate(pd.DataFrame(rows), headers='keys', tablefmt='simple', showindex=False))

    print("\n" + "=" * 80)
    print("조건별 분류")
    print("=" * 80)

    def show(title, sub, subset):
        print(f"\n[{title}] {len(subset)}개 종목 ({sub}):")
        print("-" * 80)
        if len(subset) == 0:
            print("  없음")
            return
        for _, s in subset.iterrows():
            nm = names.get(s['ticker'], s['ticker'])
            print(f"  {nm:<24}  {s['trend_detail']}, RVOL {s['rvol']:.1f}배")

    sell = df[df['signal'] == 'SELL']
    caution = df[df['condition_1_trend_breakdown'] & ~df['condition_2_volume_confirmation']]
    upside = df[~df['condition_1_trend_breakdown'] & df['condition_2_volume_confirmation']]

    show('profit taking', '10일선 하회 + 거래량 폭증 (SELL)', sell)
    show('주의 필요', '10일선 하회 + 거래량 부족 (HOLD)', caution)
    show('upside', '10일선 위 + 거래량 폭증 (BUY)', upside)
    return df


# =========================================================================
# 7. 출력 — Excel (노이즈 필터 적용)
# =========================================================================
def _recent_breakdown(d):
    if pd.isna(d):
        return False
    try:
        return pd.to_datetime(d).date() >= date.today() - timedelta(days=EXCEL_RECENT_DAYS)
    except Exception:
        return False


def _valid_gap(ba, bb):
    if pd.isna(ba) or pd.isna(bb):
        return False
    try:
        return (pd.to_datetime(bb).date() - pd.to_datetime(ba).date()).days >= EXCEL_MIN_GAP
    except Exception:
        return False


def _row(category, r, names):
    return {
        '카테고리': category,
        '종목명': names.get(r['ticker'], r['ticker']),
        '티커': r['ticker'],
        'RVOL': round(r['rvol'], 1) if pd.notna(r['rvol']) else None,
        'RSI': round(r['rsi'], 1) if r['rsi'] is not None else None,
        '10일선돌파일': r['last_ma10_break_above'],
        '10일선이탈일': r['last_ma10_break_below'],
        '추세상세': r['trend_detail'],
        '현재가': r['close_price'],
        '전일종가': r['prev_close'],
        '전일비(%)': round(r['price_change_percent'], 1),
        '10일선': r['ma10'],
        '10일선괴리율(%)': round(r['ma_distance_percent'], 1) if pd.notna(r['ma_distance_percent']) else None,
    }


def save_excel(df: pd.DataFrame, names: dict, mode: int) -> str:
    sell = [_row('profit taking', r, names) for _, r in df[df['signal'] == 'SELL'].iterrows()]
    caution = [_row('주의 필요', r, names) for _, r in
               df[df['condition_1_trend_breakdown'] & ~df['condition_2_volume_confirmation']].iterrows()]
    upside = [_row('upside', r, names) for _, r in
              df[~df['condition_1_trend_breakdown'] & df['condition_2_volume_confirmation']].iterrows()]

    sell_f = [x for x in sell if _recent_breakdown(x['10일선이탈일']) and _valid_gap(x['10일선돌파일'], x['10일선이탈일'])]
    caution_f = [x for x in caution if _recent_breakdown(x['10일선이탈일']) and _valid_gap(x['10일선돌파일'], x['10일선이탈일'])]
    print(f"\n[필터링] 이탈일 최근 {EXCEL_RECENT_DAYS}일 이내 + 돌파-이탈 간격 {EXCEL_MIN_GAP}일 이상")
    print(f"  - profit taking: {len(sell)} → {len(sell_f)}")
    print(f"  - 주의 필요    : {len(caution)} → {len(caution_f)}")
    print(f"  - upside       : {len(upside)} (필터 제외)")

    def sort_part(rows):
        if not rows:
            return pd.DataFrame()
        p = pd.DataFrame(rows)
        return p.sort_values(by=['10일선이탈일', '10일선돌파일'], ascending=[False, True], na_position='last')

    order = {'profit taking': 0, '주의 필요': 1, 'upside': 2}
    df_all = pd.DataFrame(sell_f + caution_f + upside)
    if not df_all.empty:
        df_all['_o'] = df_all['카테고리'].map(order).fillna(99)
        df_all = df_all.sort_values(by=['_o', '10일선이탈일', '10일선돌파일'],
                                    ascending=[True, False, True], na_position='last').drop(columns='_o')

    out_dir = os.path.join(CURRENT_DIR, "outputs", "start_brief")
    os.makedirs(out_dir, exist_ok=True)
    suffix = "보고용" if mode == 1 else "실시간"
    path = os.path.join(out_dir, f"조건별_분류_{suffix}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

    with pd.ExcelWriter(path, engine='openpyxl') as w:
        (df_all if not df_all.empty else pd.DataFrame()).to_excel(w, sheet_name='전체', index=False)
        if sell_f:
            sort_part(sell_f).to_excel(w, sheet_name='profit taking', index=False)
        if caution_f:
            sort_part(caution_f).to_excel(w, sheet_name='주의 필요', index=False)
        if upside:
            sort_part(upside).to_excel(w, sheet_name='upside', index=False)
    return path


# =========================================================================
# 8. main
# =========================================================================
def main():
    print("=" * 80)
    print("TAKING PROFIT SCREENER (FinanceDataReader)")
    print("=" * 80)
    print("\n티커 입력: 'code MARKET' / 코드만 / 엑셀 복붙 덩어리(한 줄) 모두 가능")
    print("  예) 005930 KS, 000660 KS, AAPL US")
    print("  예) 000660<탭>SK하이닉스087010<탭>펩트론005930<탭>삼성전자...  (코드+이름 붙어도 자동 분리)")

    # 코드 마스터 준비 (엑셀 복붙 분리의 기준)
    master = build_code_master()

    raw = input("\n티커 입력: ").strip()
    if not raw:
        print("[에러] 티커를 입력해주세요")
        return

    tickers, names = parse_ticker_input(raw, master=master, verify_unknown=True)
    if not tickers:
        print("[에러] 인식된 티커가 없습니다")
        return
    print(f"\n인식된 고유 종목: {len(tickers)}개")
    for t in tickers:
        print(f"  {t:12s} -> {names[t]}")

    # 제외 티커
    exclude = input("\n제외할 티커 입력 (코드만, 없으면 엔터): ").strip()
    if exclude:
        ex_codes = {c.strip().upper() for c in exclude.replace(',', ' ').split()}
        before = len(tickers)
        tickers = [t for t in tickers if parse_ticker(t)[0] not in ex_codes]
        print(f"✓ 제외: {before} → {len(tickers)}")

    # 모드
    print("\n[1] 보고용 (장중=전일까지, 장마감 후=당일 포함)   [2] 실시간 (당일 포함)")
    while True:
        m = input("모드 선택 (1/2): ").strip()
        if m in ('1', '2'):
            mode = int(m)
            break
    print(f"✓ 모드: {'보고용' if mode == 1 else '실시간'}")

    # 종목명 보강 (마스터/붙여넣기에서 못 채운 것만)
    names = fill_missing_names(tickers, names)

    # 병렬 분석
    print(f"\n총 {len(tickers)}개 분석 시작 (병렬 {MAX_WORKERS})\n")
    results, failed, done = [], 0, 0
    start = datetime.now()
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(analyze_ticker, t, mode): t for t in tickers}
        for fut in as_completed(futures):
            try:
                r = fut.result()
            except Exception:
                r = None
            done += 1
            if r:
                results.append(r)
            else:
                failed += 1
            filled = int(40 * done / len(tickers))
            bar = '█' * filled + '░' * (40 - filled)
            print(f"\r[{bar}] {done}/{len(tickers)} | 성공 {len(results)} 실패 {failed}", end='', flush=True)
    print(f"\n✓ 완료 - 소요 {str(datetime.now() - start).split('.')[0]}")

    if failed:
        print(f"  (실패 {failed}개: 데이터 미수신 종목 — 신형코드/최근상장 등)")
    if not results:
        print("\n[에러] 분석 결과가 없습니다")
        return

    df = print_console(results, names)

    if input("\nExcel 저장? (y/n): ").strip().lower() == 'y':
        try:
            print(f"✓ Excel: {save_excel(df, names, mode)}")
        except Exception as e:
            print(f"✗ Excel 실패: {e}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n중단되었습니다.")