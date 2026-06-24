"""
20주선(20주선) 하회 종목 스크리닝

py -3.12 under_20w.py

FinanceDataReader로 전종목 데이터를 받아 20주선(≈20주선) 하회 종목을 스크리닝합니다.
"""
import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta, time
from tabulate import tabulate

# Windows 콘솔 인코딩 설정
if sys.platform == 'win32':
    import codecs
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# 현재 스크립트의 디렉토리를 Python 경로에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import FinanceDataReader as fdr

# ── 네트워크 무한 대기 방지 ──────────────────────────────────────────────
# FDR 내부가 requests 를 쓰는데, 응답 없는 종목에서 요청이 영원히 매달려
# 스레드를 점유하면 전체가 멈춘다. socket.setdefaulttimeout 만으로는
# 커넥션 재사용/명시적 타임아웃 때문에 안 먹는 경우가 있어,
# requests 어댑터의 send 를 패치해 '타임아웃 없는 모든 요청'에 강제 타임아웃을 박는다.
HTTP_TIMEOUT = 12  # 초. analyze_tickers_parallel 의 per_ticker_timeout 보다 짧게.

import socket as _socket
_socket.setdefaulttimeout(HTTP_TIMEOUT)

try:
    import requests.adapters as _rq_adapters

    _orig_send = _rq_adapters.HTTPAdapter.send

    def _send_with_timeout(self, request, **kwargs):
        if kwargs.get("timeout") is None:
            kwargs["timeout"] = HTTP_TIMEOUT
        return _orig_send(self, request, **kwargs)

    _rq_adapters.HTTPAdapter.send = _send_with_timeout
except Exception:
    pass

MAX_DAYS_BELOW_MA20W_INCLUDED = 3

# ── 경로 설정 (igis 통합) ───────────────────────────────────────────────
#   under_20w.py 위치: igis/under_20w/under_20w.py
#   REPO_ROOT = igis/ ,  출력 = igis/outputs/under_20w/
REPO_ROOT = os.path.dirname(current_dir)
INPUT_FILE = os.path.join(current_dir, "all_stock.xlsx")   # 단말 종목 리스트 (스크립트 폴더)
OUTPUT_DIR = os.path.join(REPO_ROOT, "outputs", "under_20w")


def parse_ticker(ticker: str) -> tuple:
    parts = ticker.strip().split()
    if len(parts) >= 2:
        return (parts[0].upper(), parts[1].upper())
    return (parts[0].upper(), None)


import re as _re
_CODE_A_RE = _re.compile(r"^A?\d{6}$")  # A000660 또는 000660


def load_universe_from_terminal(file_path: str) -> tuple:
    """all_stock.xlsx (데이터가이드 단말 형식)에서 유니버스 로드.

    형식: 상단 메타행 + '코드/코드명/시가총액/시장구분' 헤더행 + 데이터행(A######).
    반환: (tickers, names, market_caps)
      - tickers: ["000660 KS", "247540 KQ", ...]  (Bloomberg식 'code MKT')
      - names:   {ticker: 종목명}
      - market_caps: {ticker: 시가총액(억원)}  ← 파일의 '시가총액(억원)' 그대로
    """
    print(f"\n[종목 리스트 로드] {file_path}")
    if not os.path.exists(file_path):
        print(f"  파일 없음: {file_path}")
        return [], {}, {}

    raw = pd.read_excel(file_path, sheet_name=0, header=None)

    # 헤더 행 탐지: '코드'와 ('시장'구분 또는 '코드명')이 같이 있는 행
    header_row = None
    for i in range(min(20, len(raw))):
        rowvals = [str(raw.iloc[i, c]).strip() for c in range(raw.shape[1])]
        if "코드" in rowvals and (any("시장" in v for v in rowvals) or "코드명" in rowvals):
            header_row = i
            break
    if header_row is None:
        print("  헤더 행(코드/시장구분)을 찾지 못했습니다.")
        return [], {}, {}

    header = [str(raw.iloc[header_row, c]).strip() for c in range(raw.shape[1])]

    def _exact_col(name):
        return next((ci for ci, h in enumerate(header) if h == name), None)

    def _contains_col(*cands):
        for cand in cands:
            for ci, h in enumerate(header):
                if cand in h:
                    return ci
        return None

    c_code = _exact_col("코드")
    c_name = _exact_col("코드명")
    c_mkt = _contains_col("시장구분", "시장")
    c_cap = _contains_col("시가총액")
    if c_code is None:
        print("  '코드' 컬럼을 찾지 못했습니다.")
        return [], {}, {}

    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    tickers, names, market_caps = [], {}, {}
    for _, r in data.iterrows():
        code_raw = str(r.iloc[c_code]).strip()
        if not _CODE_A_RE.match(code_raw):
            continue
        code6 = code_raw[1:] if code_raw.upper().startswith("A") else code_raw

        mkt = str(r.iloc[c_mkt]).strip().upper() if c_mkt is not None else ""
        if mkt not in ("KS", "KQ", "US"):
            mkt = "KS"  # 시장구분 없으면 코스피로 가정
        ticker = f"{code6} {mkt}"

        if ticker in names:  # 중복 코드 방지
            continue
        tickers.append(ticker)

        nm = str(r.iloc[c_name]).strip() if c_name is not None else ""
        names[ticker] = nm if nm and nm != "nan" else ticker

        if c_cap is not None:
            cap = pd.to_numeric(
                str(r.iloc[c_cap]).replace(",", ""), errors="coerce"
            )
            # 파일 단위가 '억원' → 그대로 사용 (환산 없음)
            market_caps[ticker] = float(cap) if pd.notna(cap) else None
        else:
            market_caps[ticker] = None

    print(f"  로드 완료: {len(tickers):,}개 종목 (시가총액 포함)")
    return tickers, names, market_caps


def _period_to_start_date(period: str) -> datetime:
    period = period.upper().strip()
    end = datetime.now()
    if period == '1M':
        return end - timedelta(days=45)
    if period == '3M':
        return end - timedelta(days=120)
    if period == '6M':
        return end - timedelta(days=210)
    if period == '1Y':
        return end - timedelta(days=400)
    if period == '2Y':
        return end - timedelta(days=760)
    if period == '3Y':
        return end - timedelta(days=1130)
    return end - timedelta(days=120)


def _previous_business_day(day):
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day


def _expected_cache_date(mode: int = 1):
    now = datetime.now()
    today = now.date()
    expected = _previous_business_day(today)
    if mode == 1 and expected == today and now.time() < time(15, 30):
        expected = _previous_business_day(today - timedelta(days=1))
    return expected


def _normalize_ohlcv(df: pd.DataFrame, verbose: bool = True) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return None
    df = df.reset_index()
    date_col = None
    for cand in ('Date', 'date', 'index'):
        if cand in df.columns:
            date_col = cand
            break
    if date_col and date_col != 'Date':
        df = df.rename(columns={date_col: 'Date'})
    required = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
    missing = [c for c in required if c not in df.columns]
    if missing:
        if verbose:
            print(f"  컬럼 부족: {missing}, 사용 가능한 컬럼: {list(df.columns)}")
        return None
    date_series = pd.to_datetime(df['Date'])
    df['Date'] = date_series.dt.tz_localize(None) if date_series.dt.tz is not None else date_series
    return df[required].copy()


def _get_cache_path(ticker: str) -> str:
    code, market = parse_ticker(ticker)
    cache_dir = os.path.join(current_dir, 'database', 'fdr_cache')
    os.makedirs(cache_dir, exist_ok=True)
    return os.path.join(cache_dir, f"{code}_{market}.csv")


def download_data(ticker: str, period: str = '1Y', verbose: bool = True, mode: int = 1) -> pd.DataFrame:
    code, market = parse_ticker(ticker)
    if market not in ('KS', 'KQ', 'US'):
        if verbose:
            print(f"  지원하지 않는 시장: {market}")
        return None
    requested_start = _period_to_start_date(period)
    end = datetime.now()
    cache_path = _get_cache_path(ticker)
    cached_df = None
    if os.path.exists(cache_path):
        try:
            cached_df = pd.read_csv(cache_path, parse_dates=['Date'])
            cached_df = cached_df[['Date', 'Open', 'High', 'Low', 'Close', 'Volume']]
        except Exception:
            cached_df = None
    if cached_df is not None and not cached_df.empty:
        latest_cached = pd.to_datetime(cached_df['Date']).max()
        earliest_cached = pd.to_datetime(cached_df['Date']).min()
        cache_is_fresh = latest_cached.date() >= _expected_cache_date(mode)
        cache_covers_period = earliest_cached <= requested_start
        cache_has_20w_history = (latest_cached - earliest_cached).days >= 140
        if cache_is_fresh and (cache_covers_period or cache_has_20w_history):
            return cached_df[cached_df['Date'] >= requested_start][['Date', 'Open', 'High', 'Low', 'Close', 'Volume']].copy()
        if earliest_cached > requested_start:
            fetch_start = requested_start
        else:
            fetch_start = max(requested_start, latest_cached)
    else:
        fetch_start = requested_start
    try:
        new_df = fdr.DataReader(code, fetch_start.strftime('%Y-%m-%d'), end.strftime('%Y-%m-%d'))
    except Exception as e:
        if verbose:
            print(f"  데이터 다운로드 실패: {e}")
        new_df = None
    new_df = _normalize_ohlcv(new_df, verbose=verbose)
    frames = [df for df in (cached_df, new_df) if df is not None and not df.empty]
    if not frames:
        return None
    df = pd.concat(frames, ignore_index=True)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.drop_duplicates(subset=['Date'], keep='last').sort_values('Date')
    df.to_csv(cache_path, index=False, encoding='utf-8-sig')
    return df[df['Date'] >= requested_start][['Date', 'Open', 'High', 'Low', 'Close', 'Volume']].copy()


def analyze_single_ticker(ticker: str, period: str = '1Y', mode: int = 1) -> dict:
    try:
        df = download_data(ticker, period=period, verbose=False, mode=mode)
        if df is None or len(df) == 0:
            return None
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'])
            df['_date_only'] = df['Date'].dt.date
            from datetime import datetime as _dt, time as _time
            _now = _dt.now()
            _today = _now.date()
            _market_close = _time(15, 30)
            if mode == 1 and _now.time() < _market_close:
                is_today = df['_date_only'] == _today
                if is_today.any():
                    df = df[~is_today].copy()
            df = df.drop(columns=['_date_only'])
        df = df.copy()
        if 'Date' in df.columns:
            df = df.set_index('Date')
        if not isinstance(df.index, pd.DatetimeIndex):
            df.index = pd.to_datetime(df.index)
        date_span = (df.index[-1] - df.index[0]).days
        if date_span < 140:
            return None
        df['MA20W'] = df['Close'].shift(1).rolling('140D').mean()
        df['MA10'] = df['Close'].shift(1).rolling(window=10).mean()
        avg_vol = df['Volume'].shift(1).rolling(window=20).mean()
        df['RVOL'] = df['Volume'] / avg_vol
        delta = df['Close'].diff()
        gain = delta.where(delta > 0, 0.0)
        loss = (-delta).where(delta < 0, 0.0)
        avg_gain = gain.rolling(window=14).mean()
        avg_loss = loss.rolling(window=14).mean()
        rs = avg_gain / avg_loss
        df['RSI'] = 100 - (100 / (1 + rs))
        latest = df.iloc[-1]
        if pd.isna(latest['MA20W']):
            return None
        below_ma20w = latest['Close'] < latest['MA20W']
        if not below_ma20w:
            pass
        ma20w_distance_pct = ((latest['Close'] - latest['MA20W']) / latest['MA20W']) * 100
        ma10_distance_pct = None
        if pd.notna(latest.get('MA10')):
            ma10_distance_pct = ((latest['Close'] - latest['MA10']) / latest['MA10']) * 100
        position = (df['Close'] >= df['MA20W']).astype(int)
        position_shift = position.shift(1)
        break_below_mask = (position_shift == 1) & (position == 0)
        break_above_mask = (position_shift == 0) & (position == 1)
        dates = pd.Series(df.index.strftime('%Y-%m-%d'), index=df.index)
        last_break_below = None
        last_break_above = None
        for idx in df.index:
            if break_below_mask.loc[idx]:
                last_break_below = dates.loc[idx]
        for idx in df.index:
            if break_above_mask.loc[idx]:
                last_break_above = dates.loc[idx]
        days_below = 0
        for idx in reversed(df.index.tolist()):
            if pd.isna(df.loc[idx, 'MA20W']):
                break
            if df.loc[idx, 'Close'] < df.loc[idx, 'MA20W']:
                days_below += 1
            else:
                break
        prev_close = None
        price_change_pct = 0
        if len(df) >= 2:
            prev_close = df.iloc[-2]['Close']
            price_change_pct = ((latest['Close'] - prev_close) / prev_close) * 100
        current_position = 'below' if latest['Close'] < latest['MA20W'] else 'above'
        if current_position == 'below':
            trend_detail = f"20주선 위({last_break_above or '?'}) → 20주선 아래({last_break_below or '?'})"
        else:
            trend_detail = f"20주선 아래({last_break_below or '?'}) → 20주선 위({last_break_above or '?'})"
        result = {
            'ticker': ticker,
            'close_price': latest['Close'],
            'prev_close': prev_close,
            'price_change_percent': price_change_pct,
            'ma20w': latest['MA20W'],
            'ma20w_distance_percent': ma20w_distance_pct,
            'ma10': latest['MA10'] if pd.notna(latest.get('MA10')) else None,
            'ma10_distance_percent': ma10_distance_pct,
            'below_ma20w': below_ma20w,
            'last_ma20w_break_below': last_break_below,
            'last_ma20w_break_above': last_break_above,
            'days_below_ma20w': days_below,
            'trend_detail': trend_detail,
            'rvol': latest['RVOL'] if pd.notna(latest.get('RVOL')) else None,
            'rsi': float(latest['RSI']) if pd.notna(latest.get('RSI')) else None,
        }
        return result
    except Exception as e:
        return None


def analyze_tickers_parallel(tickers: list, period: str = '1Y', max_workers: int = 12,
                             mode: int = 1, per_ticker_timeout: int = 30,
                             max_rounds: int = 4) -> list:
    """전 종목 병렬 분석 (누락 방지: 타임아웃 + 라운드 재시도).

    - per_ticker_timeout: 종목 1개 분석이 이 시간(초)을 넘으면 그 라운드에선 포기 → 다음 라운드 재시도
    - max_rounds: 실패/타임아웃 종목을 최대 몇 라운드까지 재시도할지
    - 한 번 성공한 종목은 캐시에 남아 다음 라운드에서 즉시 처리됨
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError as FutureTimeout
    from threading import Lock

    mode_name = "보고용 (완성된 일봉)" if mode == 1 else "실시간 (현재 시점)"
    print(f"\n총 {len(tickers)}개 종목 분석 시작...")
    print(f"분석 모드: {mode_name}")
    print(f"병렬 처리: {max_workers}개 동시 실행  |  종목당 타임아웃: {per_ticker_timeout}초  |  최대 {max_rounds}라운드 재시도")
    print(f"데이터 기간: {period} (20주선 계산용)\n")

    results = []
    results_lock = Lock()
    done_tickers = set()
    pending = list(tickers)
    overall_start = datetime.now()

    def _analyze(ticker):
        try:
            return (ticker, analyze_single_ticker(ticker, period=period, mode=mode), None)
        except Exception as e:
            return (ticker, None, str(e))

    for round_no in range(1, max_rounds + 1):
        if not pending:
            break
        total_this_round = len(pending)
        if round_no == 1:
            print(f"[라운드 {round_no}] {total_this_round}개 분석")
        else:
            print(f"\n[라운드 {round_no}] 이전 라운드 미완료 {total_this_round}개 재시도")

        round_failed = []
        completed = 0
        round_start = datetime.now()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_ticker = {executor.submit(_analyze, t): t for t in pending}
            for future in as_completed(future_to_ticker):
                src_ticker = future_to_ticker[future]
                try:
                    ticker, result, error = future.result(timeout=per_ticker_timeout)
                except FutureTimeout:
                    round_failed.append(src_ticker)
                    ticker, result = src_ticker, None
                except Exception:
                    round_failed.append(src_ticker)
                    ticker, result = src_ticker, None
                else:
                    if result:
                        with results_lock:
                            if ticker not in done_tickers:
                                results.append(result)
                                done_tickers.add(ticker)
                    else:
                        round_failed.append(ticker)

                completed += 1
                elapsed = datetime.now() - round_start
                progress = completed / total_this_round * 100
                rate = completed / elapsed.total_seconds() if elapsed.total_seconds() > 0 else 0
                remaining = (total_this_round - completed) / rate if rate > 0 else 0
                print(f"\r[라운드 {round_no}] {completed}/{total_this_round} ({progress:.1f}%) "
                      f"| 경과: {str(elapsed).split('.')[0]} | 속도: {rate:.2f}종목/초 | "
                      f"남은시간: ~{int(remaining/60)}분 {int(remaining%60)}초", end='', flush=True)
            # 타임아웃 등으로 result()를 못 받은 future가 남아있지 않도록 정리
            executor.shutdown(wait=False, cancel_futures=True)

        print()
        # 이번 라운드에서 실패/타임아웃한 것만 다음 라운드로 (이미 성공한 건 제외)
        pending = [t for t in dict.fromkeys(round_failed) if t not in done_tickers]
        print(f"  라운드 {round_no} 완료 — 누적 성공 {len(results)}개, 남은 미완료 {len(pending)}개")

    total_time = datetime.now() - overall_start
    print(f"\n✓ 분석 완료 - 소요시간: {str(total_time).split('.')[0]}")
    print(f"  성공: {len(results)}개 / 전체 {len(tickers)}개")

    if pending:
        print(f"\n⚠️  {max_rounds}라운드 후에도 못 받은 종목 {len(pending)}개 (누락):")
        for ticker in pending[:30]:
            print(f"  - {ticker}")
        if len(pending) > 30:
            print(f"  ... 외 {len(pending) - 30}개")
        print("  → 다시 실행하면 캐시 덕분에 빠르게 받아 채울 수 있습니다.")
    else:
        print("  누락 없음 — 전 종목 분석 완료 ✓")

    return results


def filter_below_ma20w(results: list) -> pd.DataFrame:
    if not results:
        return pd.DataFrame()
    df = pd.DataFrame(results)
    print(f"\n[디버깅] 전체 분석 결과: {len(df)}개")
    below = df['below_ma20w'] == True
    print(f"[디버깅] 20주선 하회 종목: {below.sum()}개")
    days_below = df.get('days_below_ma20w', 0)
    recent_breakdown = days_below <= MAX_DAYS_BELOW_MA20W_INCLUDED
    print(f"[디버깅] 하회일수 4일 이상 제외: {(below & ~recent_breakdown).sum()}개")
    filtered = df[below & recent_breakdown].copy()
    if 'last_ma20w_break_below' in filtered.columns:
        filtered = filtered.sort_values('last_ma20w_break_below', ascending=False, na_position='last')
    return filtered


def main():
    print("="*80)
    print("20주선(20주선) 하회 종목 스크리닝")
    print("="*80)
    print("\n⚠️  주의사항:")
    print("  1. 인터넷 연결이 필요합니다")
    print("  2. FinanceDataReader로 데이터를 가져옵니다")
    print("  3. 전종목 분석은 약 15~20분 소요됩니다 (1년 데이터)")
    print("\n" + "="*80)
    print("분석 모드 선택")
    print("="*80)
    print("\n[1] 보고용 - 전일 또는 장마감 후 당일 (완성된 일봉)")
    print("    → 장중: 전일까지의 데이터 사용")
    print("    → 장마감 후(15:30 이후): 당일 포함")
    print("\n[2] 실시간 - 현재 시점의 미완성 데이터 포함")
    while True:
        mode_input = input("\n모드 선택 (1 또는 2): ").strip()
        if mode_input in ['1', '2']:
            mode = int(mode_input)
            break
        print("1 또는 2를 입력해주세요.")
    mode_name = "보고용 (완성된 일봉)" if mode == 1 else "실시간 (현재 시점)"
    print(f"\n✓ 선택된 모드: {mode_name}")
    print("\n" + "="*80)
    print("티커 리스트 로드")
    print("="*80)
    all_tickers, universe_names, universe_caps = load_universe_from_terminal(INPUT_FILE)
    if not all_tickers:
        print("\n[에러] 종목을 읽을 수 없습니다 (all_stock.xlsx 확인)")
        return
    print(f"\n총 {len(all_tickers)}개 종목을 분석합니다.")
    print("\n" + "="*80)
    print("전종목 분석 시작 (1년 데이터 - 20주선 계산용)")
    print("="*80)
    results = analyze_tickers_parallel(
        all_tickers, period='1Y', max_workers=12, mode=mode,
        per_ticker_timeout=20, max_rounds=4,
    )
    if not results:
        print("\n[에러] 분석 결과가 없습니다")
        return
    print("\n" + "="*80)
    print("스크리닝: 20주선(20주선) 하회 종목")
    print("="*80)
    filtered_df = filter_below_ma20w(results)
    if filtered_df.empty:
        print("\n조건을 만족하는 종목이 없습니다.")
        return
    print(f"\n✓ {len(filtered_df)}개 종목이 20주선 하회 조건을 만족합니다")
    print("\n[종목명/시가총액 매핑 — all_stock.xlsx 기준]")
    filtered_tickers = filtered_df['ticker'].tolist()
    # 종목명/시가총액은 단말 파일(all_stock.xlsx)에서 읽은 값을 그대로 사용
    ticker_names = {t: universe_names.get(t, t) for t in filtered_tickers}
    market_caps = {t: universe_caps.get(t) for t in filtered_tickers}
    if filtered_df.empty:
        print("\n시가총액 1000억 이상인 20주선 하회 종목이 없습니다.")
        return
    print("\n" + "="*80)
    print("스크리닝 결과 (20주선 괴리율 순)")
    print("="*80)
    summary_data = []
    for _, row in filtered_df.iterrows():
        ticker = row['ticker']
        security_name = ticker_names.get(ticker, ticker)
        if row.get('prev_close') and row['prev_close'] > 0:
            price_change_str = f"{row['price_change_percent']:+.1f}%"
        else:
            price_change_str = "-"
        rvol_str = f"{row['rvol']:.1f}배" if row.get('rvol') is not None and pd.notna(row['rvol']) else "-"
        summary_data.append({
            '종목': security_name[:30],
            '티커': ticker,
            '현재가': f"{row['close_price']:.0f}",
            '전일비': price_change_str,
            '20주선': f"{row['ma20w']:.0f}",
            '괴리율': f"{row['ma20w_distance_percent']:+.1f}%",
            'RVOL': rvol_str,
            '이탈일': row.get('last_ma20w_break_below', '?'),
        })
    summary_df_display = pd.DataFrame(summary_data)
    print("\n" + tabulate(summary_df_display, headers='keys', tablefmt='simple', showindex=False))
    print("\n" + "="*80)
    print("상세 정보 (20주선 하회 종목)")
    print("="*80)
    for _, row in filtered_df.iterrows():
        ticker = row['ticker']
        security_name = ticker_names.get(ticker, ticker)
        market_cap = market_caps.get(ticker)
        if any('\uac00' <= char <= '\ud7a3' for char in security_name):
            name_padded = f"{security_name:<15}"
        else:
            name_padded = f"{security_name:<30}"
        trend_info = row['trend_detail']
        rvol_str = f"RVOL {row['rvol']:.1f}배" if row.get('rvol') is not None and pd.notna(row['rvol']) else "RVOL N/A"
        if market_cap is not None:
            market_cap_str = f"시총 {market_cap:,.0f}억"
        else:
            market_cap_str = "시총 N/A"
        print(f"  {name_padded}  {trend_info}, {rvol_str}, {market_cap_str}")
    print("\n" + "="*80)
    print("TOP 5 종목 (20주선 이탈일 최근순)")
    print("="*80)
    top5_recent = filtered_df.sort_values('last_ma20w_break_below', ascending=False, na_position='last').head(5)
    for _, row in top5_recent.iterrows():
        ticker = row['ticker']
        name = ticker_names.get(ticker, ticker)
        cap = market_caps.get(ticker)
        cap_str = f"{cap:,.0f}억" if cap is not None and pd.notna(cap) else "N/A"
        rvol_str = f"{row['rvol']:.1f}배" if row.get('rvol') is not None and pd.notna(row['rvol']) else "N/A"
        print(f"  {name:<15} | 이탈일: {row['last_ma20w_break_below']} | RVOL: {rvol_str} | 시총: {cap_str}")
    print("\n" + "="*80)
    print("TOP 5 종목 (20주선 이탈일 오래된순)")
    print("="*80)
    top5_old = filtered_df.sort_values('last_ma20w_break_below', ascending=True, na_position='last').head(5)
    for _, row in top5_old.iterrows():
        ticker = row['ticker']
        name = ticker_names.get(ticker, ticker)
        cap = market_caps.get(ticker)
        cap_str = f"{cap:,.0f}억" if cap is not None and pd.notna(cap) else "N/A"
        rvol_str = f"{row['rvol']:.1f}배" if row.get('rvol') is not None and pd.notna(row['rvol']) else "N/A"
        print(f"  {name:<15} | 이탈일: {row['last_ma20w_break_below']} | RVOL: {rvol_str} | 시총: {cap_str}")
    print("\n" + "="*80)
    print("엑셀 파일 저장 중...")
    save_dir = OUTPUT_DIR
    os.makedirs(save_dir, exist_ok=True)
    output_date = datetime.now().strftime("%Y%m%d")
    output_filename = os.path.join(save_dir, f"under_20w_{output_date}.xlsx")
    base_columns = [
        'ticker', 'rvol',
        'last_ma20w_break_below', 'last_ma20w_break_above',
        'days_below_ma20w',
        'trend_detail',
        'close_price', 'prev_close', 'price_change_percent',
        'ma20w', 'ma20w_distance_percent',
        'ma10', 'ma10_distance_percent',
    ]
    if 'rsi' in filtered_df.columns:
        base_columns.insert(2, 'rsi')
    save_df = filtered_df[base_columns].copy()
    save_df.insert(0, '종목명', save_df['ticker'].map(ticker_names))
    save_df.insert(1, '티커', save_df['ticker'])
    save_df['종목명'] = save_df['종목명'].fillna(save_df['티커'])
    save_df = save_df.drop(columns=['ticker'])
    save_df.insert(2, '시가총액(억원)', save_df['티커'].map(market_caps))
    save_df['price_change_percent'] = save_df['price_change_percent'].round(1)
    save_df['ma20w_distance_percent'] = save_df['ma20w_distance_percent'].round(1)
    if 'ma10_distance_percent' in save_df.columns:
        save_df['ma10_distance_percent'] = save_df['ma10_distance_percent'].round(1)
    if 'rvol' in save_df.columns:
        save_df['rvol'] = save_df['rvol'].round(1)
    if 'rsi' in save_df.columns:
        save_df['rsi'] = save_df['rsi'].round(1)
    rename_dict = {
        'rvol': 'RVOL',
        'last_ma20w_break_below': '100일선이탈일',
        'last_ma20w_break_above': '100일선돌파일',
        'days_below_ma20w': '하회일수',
        'trend_detail': '추세상세',
        'close_price': '현재가',
        'prev_close': '전일종가',
        'price_change_percent': '전일비(%)',
        'ma20w': '100일선(20주선)',
        'ma20w_distance_percent': '100일선괴리율',
        'ma10': '10일선',
        'ma10_distance_percent': '10일선괴리율',
    }
    if 'rsi' in save_df.columns:
        rename_dict['rsi'] = 'RSI'
    save_df = save_df.rename(columns=rename_dict)
    from datetime import timedelta
    today = date.today()
    cutoff = today - timedelta(days=5)
    cutoff_ts = pd.Timestamp(cutoff)
    before_filter = len(save_df)
    save_df['이탈일_date'] = pd.to_datetime(save_df['100일선이탈일'], errors='coerce')
    save_df = save_df[save_df['이탈일_date'] >= cutoff_ts].copy()
    save_df = save_df.drop(columns=['이탈일_date'])
    after_filter = len(save_df)
    mode_str = "보고용" if mode == 1 else "실시간"
    print(f"\n[{mode_str}] 최근 5일 내({cutoff}~{today}) 20주선 이탈 종목: {before_filter}개 → {after_filter}개")
    before_gap = len(save_df)
    def _has_valid_gap(row):
        breakdown = row['100일선이탈일']
        breakout = row['100일선돌파일']
        if pd.isna(breakdown) or pd.isna(breakout):
            return False
        try:
            gap = (pd.to_datetime(breakdown) - pd.to_datetime(breakout)).days
            return gap >= 5
        except Exception:
            return False
    save_df = save_df[save_df.apply(_has_valid_gap, axis=1)].copy()
    after_gap = len(save_df)
    print(f"[갭 필터] 이탈일-돌파일 5일 이상: {before_gap}개 → {after_gap}개")
    if not save_df.empty and '시가총액(억원)' in save_df.columns:
        save_df['시가총액_num'] = pd.to_numeric(save_df['시가총액(억원)'], errors='coerce')
        save_df = save_df.sort_values(by=['시가총액_num'], ascending=[False], na_position='last')
        save_df = save_df.drop(columns=['시가총액_num'])
    final_columns = [
        '종목명', '티커', '시가총액(억원)', 'RVOL', 'RSI',
        '100일선이탈일', '100일선돌파일', '하회일수', '추세상세',
        '현재가', '전일종가', '전일비(%)',
        '100일선(20주선)', '100일선괴리율',
        '10일선', '10일선괴리율'
    ]
    for col in final_columns:
        if col not in save_df.columns:
            save_df[col] = pd.NA
    save_df = save_df[final_columns]
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        save_df.to_excel(writer, sheet_name='스크리닝결과', index=False)
    print(f"\n[저장 완료] {output_filename}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n프로그램이 중단되었습니다.")
    except Exception as e:
        print(f"\n[에러] 분석 실패: {e}")
        import traceback
        traceback.print_exc()