# -*- coding: utf-8 -*-
"""
CISS Risk Scoring - Multi-Source Data Loader (No Bloomberg)
============================================================
무료 API + 프록시 계산으로 13개 시계열 수집

Sources:
  - ECOS (한국은행)  : CD, 국고채(3M/3Y/10Y), USD/KRW, CRS 1Y   [ECOS_API_KEY]
  - FRED             : US 10Y Treasury (CDS 프록시용)           [FRED_API_KEY]
  - yfinance         : MOVE Index, KOSPI(^KS11), 금융업종(139270.KS)
  - Proxy (계산)     : FX 내재변동성, CDS, VKOSPI

출력 컬럼명은 기존 BloombergDataLoader와 100% 동일 → transforms.py 수정 불필요.
API 키는 igis 루트의 .env에서 자동 로드(env_loader, 의존성 없음).
"""

import os
import sys
import time
import warnings
from datetime import datetime
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd
import requests

# igis 공용 .env 로더 (python-dotenv 대체, 의존성 없음)
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..'))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)
try:
    from env_loader import load_dotenv
    load_dotenv()
except ImportError:
    pass   # env_loader 없어도 OS 환경변수로 동작


def _http_get_with_retry(url: str, params: dict = None,
                         retries: int = 3, backoff: float = 1.5, timeout: int = 30):
    """HTTP GET with retry on 5xx / connection errors."""
    last_exc = None
    for attempt in range(retries):
        try:
            resp = requests.get(url, params=params, timeout=timeout)
            if resp.status_code < 500:
                resp.raise_for_status()
                return resp
            last_exc = requests.HTTPError(f"{resp.status_code} server error")
        except (requests.ConnectionError, requests.Timeout, requests.HTTPError) as exc:
            last_exc = exc
        time.sleep(backoff ** attempt)
    raise last_exc


# -----------------------------------------------------------------------------
# Optional dependencies
# -----------------------------------------------------------------------------
try:
    import yfinance as yf
    HAS_YFINANCE = True
except ImportError:
    HAS_YFINANCE = False
    warnings.warn("[data_loader_v2] yfinance not installed. `pip install yfinance`")


# -----------------------------------------------------------------------------
# API Keys (.env에서 로드됨)
# -----------------------------------------------------------------------------
ECOS_API_KEY = os.getenv('ECOS_API_KEY', 'sample')
FRED_API_KEY = os.getenv('FRED_API_KEY', '')


# =============================================================================
# 1. ECOS (한국은행) Loader
# =============================================================================
class ECOSLoader:
    """한국은행 ECOS REST API loader.

    URL pattern:
      /api/StatisticSearch/{KEY}/json/kr/{start}/{end}/{stat}/{cycle}/{from}/{to}/{item}
    """

    BASE_URL = "https://ecos.bok.or.kr/api/StatisticSearch"

    # 통계표 817Y002 = 시장금리(일별), 731Y001 = 환율(일별)
    SERIES_MAP: Dict[str, Tuple[str, str, str]] = {
        'CD_91':   ('817Y002', '010150000', 'D'),  # CD(91일)
        'KTB_3M':  ('817Y002', '010195000', 'D'),  # 통안증권 91일 (KTB 3M 대용)
        'KTB_3Y':  ('817Y002', '010200000', 'D'),  # 국고채 3년
        'KTB_10Y': ('817Y002', '010210000', 'D'),  # 국고채 10년
        'USDKRW':  ('731Y001', '0000001', 'D'),    # 원/달러 매매기준율
        'CRS_1Y':  ('817Y002', '010260000', 'D'),  # 통화스왑 1Y
    }

    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv('ECOS_API_KEY', 'sample')

    def fetch(self, series_key: str, start_date: str, end_date: str) -> pd.Series:
        if series_key not in self.SERIES_MAP:
            raise KeyError(f"Unknown ECOS series: {series_key}")

        stat_code, item_code, cycle = self.SERIES_MAP[series_key]
        start_str = pd.to_datetime(start_date).strftime('%Y%m%d')
        end_str = pd.to_datetime(end_date).strftime('%Y%m%d')

        url = (
            f"{self.BASE_URL}/{self.api_key}/json/kr/1/100000/"
            f"{stat_code}/{cycle}/{start_str}/{end_str}/{item_code}"
        )

        try:
            resp = _http_get_with_retry(url)
            payload = resp.json()
        except Exception as exc:
            print(f"  [ECOS] {series_key}: request failed ({exc})")
            return pd.Series(dtype=float, name=series_key)

        rows = payload.get('StatisticSearch', {}).get('row', [])
        if not rows:
            err = payload.get('RESULT', {}).get('MESSAGE', 'no rows')
            print(f"  [ECOS] {series_key}: empty response ({err})")
            return pd.Series(dtype=float, name=series_key)

        df = pd.DataFrame(rows)
        df['date'] = pd.to_datetime(df['TIME'], format='%Y%m%d')
        df['value'] = pd.to_numeric(df['DATA_VALUE'], errors='coerce')
        s = df.set_index('date')['value'].sort_index()
        s.name = series_key
        return s


# =============================================================================
# 2. FRED Loader
# =============================================================================
class FREDLoader:
    """FRED REST API loader."""

    BASE_URL = "https://api.stlouisfed.org/fred/series/observations"

    SERIES_MAP: Dict[str, str] = {
        'US_10Y':   'DGS10',     # US 10Y Treasury yield
        'USDKRW':   'DEXKOUS',   # USD/KRW (ECOS 백업)
    }

    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.getenv('FRED_API_KEY', '')

    def fetch(self, series_key: str, start_date: str, end_date: str) -> pd.Series:
        if not self.api_key:
            print(f"  [FRED] {series_key}: FRED_API_KEY not set, skipping")
            return pd.Series(dtype=float, name=series_key)

        if series_key not in self.SERIES_MAP:
            raise KeyError(f"Unknown FRED series: {series_key}")

        params = {
            'series_id': self.SERIES_MAP[series_key],
            'api_key': self.api_key,
            'file_type': 'json',
            'observation_start': pd.to_datetime(start_date).strftime('%Y-%m-%d'),
            'observation_end': pd.to_datetime(end_date).strftime('%Y-%m-%d'),
        }

        try:
            resp = _http_get_with_retry(self.BASE_URL, params=params)
            obs = resp.json().get('observations', [])
        except Exception as exc:
            print(f"  [FRED] {series_key}: request failed ({exc})")
            return pd.Series(dtype=float, name=series_key)

        if not obs:
            return pd.Series(dtype=float, name=series_key)

        df = pd.DataFrame(obs)
        df['date'] = pd.to_datetime(df['date'])
        df['value'] = pd.to_numeric(df['value'], errors='coerce')
        s = df.set_index('date')['value'].sort_index()
        s.name = series_key
        return s


# =============================================================================
# 3. yfinance helper
# =============================================================================
def fetch_yfinance(ticker: str, start_date: str, end_date: str,
                   field: str = 'Close') -> pd.Series:
    """yfinance OHLCV 시리즈 가져오기. 실패 시 빈 Series."""
    if not HAS_YFINANCE:
        return pd.Series(dtype=float)
    try:
        df = yf.download(
            ticker, start=start_date, end=end_date,
            progress=False, auto_adjust=False, group_by='column',
        )
        if df.empty:
            return pd.Series(dtype=float)
        # multiindex 처리
        if isinstance(df.columns, pd.MultiIndex):
            if field in df.columns.get_level_values(0):
                s = df[field].iloc[:, 0]
            else:
                s = df.iloc[:, 0]
        else:
            s = df[field] if field in df.columns else df.iloc[:, 0]
        s.index = pd.to_datetime(s.index)
        s.name = ticker
        return s
    except Exception as exc:
        print(f"  [yfinance] {ticker}: {exc}")
        return pd.Series(dtype=float)


# =============================================================================
# 4. Multi-Source Loader (main entry)
# =============================================================================
class MultiSourceDataLoader:
    """Bloomberg를 사용하지 않는 무료 데이터 통합 로더.

    출력 컬럼은 기존 BloombergDataLoader와 동일하므로 transforms.py 그대로 사용.
    """

    def __init__(self, start_date: str = '2024-01-01',
                 end_date: Optional[str] = None):
        self.start_date = start_date
        self.end_date = end_date or datetime.now().strftime('%Y-%m-%d')
        self.ecos = ECOSLoader()
        self.fred = FREDLoader()

    # ---------- public API ----------------------------------------------------
    def fetch_all_data(self) -> pd.DataFrame:
        print(f"[INFO] Multi-source fetch: {self.start_date} ~ {self.end_date}")
        cols: Dict[str, pd.Series] = {}

        # --- ECOS: 금리, 환율, CRS ---
        print("[INFO] Fetching ECOS...")
        cd     = self._safe(self.ecos.fetch, 'CD_91',   self.start_date, self.end_date)
        ktb3m  = self._safe(self.ecos.fetch, 'KTB_3M',  self.start_date, self.end_date)
        ktb3y  = self._safe(self.ecos.fetch, 'KTB_3Y',  self.start_date, self.end_date)
        ktb10y = self._safe(self.ecos.fetch, 'KTB_10Y', self.start_date, self.end_date)
        usdkrw = self._safe(self.ecos.fetch, 'USDKRW',  self.start_date, self.end_date)
        crs1y  = self._safe(self.ecos.fetch, 'CRS_1Y',  self.start_date, self.end_date)

        cols['KWCDC_Curncy']    = cd
        cols['GVSK3M_Index']    = ktb3m if not ktb3m.empty else ktb3y
        cols['GVSK3YR_Index']   = ktb3y
        cols['GVSK10YR_Index']  = ktb10y
        cols['USDKRW_Curncy']   = usdkrw if not usdkrw.empty else \
                                  self._safe(self.fred.fetch, 'USDKRW', self.start_date, self.end_date)
        cols['KWSWNI1_Curncy']  = crs1y

        # --- FRED: US 10Y ---
        print("[INFO] Fetching FRED (US 10Y)...")
        us10y = self._safe(self.fred.fetch, 'US_10Y', self.start_date, self.end_date)
        if us10y.empty:
            print("  [FRED] US_10Y unavailable; trying yfinance ^TNX fallback")
            us10y = fetch_yfinance('^TNX', self.start_date, self.end_date, 'Close')

        # --- yfinance: MOVE, KOSPI, 금융업종 ---
        print("[INFO] Fetching yfinance (MOVE, KOSPI, 금융 ETF)...")
        cols['MOVE_Index']  = fetch_yfinance('^MOVE',     self.start_date, self.end_date, 'Close')

        kospi_close  = fetch_yfinance('^KS11',     self.start_date, self.end_date, 'Close')
        kospi_volume = fetch_yfinance('^KS11',     self.start_date, self.end_date, 'Volume')
        cols['KOSPI_Index']        = kospi_close
        cols['KOSPI_Index_VOLUME'] = kospi_volume

        # TIGER 200 금융 ETF (KOSPI 200 금융지수 추종) - 금융업종 대용
        cols['KOSPFIN_Index'] = fetch_yfinance('139270.KS', self.start_date, self.end_date, 'Close')

        # --- Proxy 1: FX 내재변동성 = USD/KRW 20일 실현변동성 (annualized, %) ---
        print("[INFO] Computing proxies (FX vol, CDS, VKOSPI)...")
        if not cols['USDKRW_Curncy'].empty:
            ret = cols['USDKRW_Curncy'].pct_change()
            cols['USDKRWV1M_BGN_Curncy'] = ret.rolling(20).std() * np.sqrt(252) * 100
        else:
            cols['USDKRWV1M_BGN_Curncy'] = pd.Series(dtype=float)

        # --- Proxy 2: 한국 CDS 5Y ≈ (KTB 10Y − US 10Y) × 100bp ---
        if not ktb10y.empty and not us10y.empty:
            joined = pd.concat([ktb10y, us10y], axis=1).ffill()
            joined.columns = ['kr', 'us']
            cols['CKREA1U5_CBGN_Curncy'] = (joined['kr'] - joined['us']) * 100
        elif not ktb10y.empty:
            print("  [Proxy] US 10Y unavailable; using KTB 10Y level as FI1 fallback")
            cols['CKREA1U5_CBGN_Curncy'] = ktb10y * 100
        else:
            cols['CKREA1U5_CBGN_Curncy'] = pd.Series(dtype=float)

        # --- Proxy 3: VKOSPI = KOSPI 20일 실현변동성 (annualized, %) ---
        if not kospi_close.empty:
            ret = kospi_close.pct_change()
            cols['VKOSPI_Index'] = ret.rolling(20).std() * np.sqrt(252) * 100
        else:
            cols['VKOSPI_Index'] = pd.Series(dtype=float)

        # --- 병합 ---
        df = pd.concat(cols, axis=1)
        df.index = pd.to_datetime(df.index)
        df = df.sort_index().dropna(how='all')

        print(f"[INFO] Loaded {len(df)} rows, {df.shape[1]} columns")
        self._report_coverage(df)
        return df

    def resample_weekly(self, df: pd.DataFrame, rule: str = 'W-FRI') -> pd.DataFrame:
        return df.resample(rule).last().dropna(how='all')

    # ---------- helpers -------------------------------------------------------
    @staticmethod
    def _safe(fn, *args, **kwargs) -> pd.Series:
        try:
            s = fn(*args, **kwargs)
            return s if s is not None else pd.Series(dtype=float)
        except Exception as exc:
            print(f"  [ERR] {fn.__qualname__}: {exc}")
            return pd.Series(dtype=float)

    @staticmethod
    def _report_coverage(df: pd.DataFrame) -> None:
        expected = [
            'KWCDC_Curncy', 'GVSK3M_Index', 'GVSK3YR_Index', 'GVSK10YR_Index',
            'MOVE_Index', 'KOSPI_Index', 'KOSPI_Index_VOLUME', 'VKOSPI_Index',
            'USDKRW_Curncy', 'USDKRWV1M_BGN_Curncy', 'KWSWNI1_Curncy',
            'CKREA1U5_CBGN_Curncy', 'KOSPFIN_Index',
        ]
        print("  Column coverage:")
        for c in expected:
            n = df[c].notna().sum() if c in df.columns else 0
            mark = 'OK ' if n > 0 else '-- '
            print(f"    [{mark}] {c:30s} ({n} non-null)")


# =============================================================================
# Public API (mirrors data_loader.load_raw_data signature)
# =============================================================================
def load_raw_data_v2(
    start_date: str = '2024-01-01',
    end_date: Optional[str] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """원본 데이터 로드 + 주간 리샘플링."""
    loader = MultiSourceDataLoader(start_date, end_date)
    daily = loader.fetch_all_data()
    weekly = loader.resample_weekly(daily)
    return daily, weekly


if __name__ == '__main__':
    daily, weekly = load_raw_data_v2('2024-01-01')
    print("\n[Daily tail]")
    print(daily.tail())
    print("\n[Weekly tail]")
    print(weekly.tail())
