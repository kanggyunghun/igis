# -*- coding: utf-8 -*-
"""
CISS Risk Scoring - Transforms Module
======================================
원본 데이터 → 15개 지표 변환 → ECDF 정규화

[2026-07-03 수정]
  - _eq3_illiquidity: 0-volume 방어(0→NaN→ffill), inf→NaN 치환,
    rolling min_periods=15 완화 (NaN 1~2개가 20일 윈도우를 오염시키지 않도록)
  - transform_all: dropna 직전에 inf→NaN 치환 + 꼬리 결측 제한적 ffill(limit=5)
    → 특정 지표의 단기 NaN이 전체 날짜를 절단하는 문제 방지
"""

import pandas as pd
import numpy as np
from typing import Dict, Tuple


class IndicatorTransformer:
    """원본 데이터를 15개 CISS 지표로 변환"""

    def __init__(self, vol_window: int = 20):
        self.vol_window = vol_window

    def transform_all(self, df: pd.DataFrame) -> pd.DataFrame:
        """모든 지표 변환 수행"""
        # 결측치 forward-fill 처리
        df = df.ffill()
        indicators = pd.DataFrame(index=df.index)

        # Money Market (MM1, MM2, MM3)
        indicators['MM1'] = self._mm1_cd_level(df)
        indicators['MM2'] = self._mm2_cd_ktb_spread(df)
        indicators['MM3'] = self._mm3_cd_volatility(df)

        # Bond Market (BD1, BD2, BD3)
        indicators['BD1'] = self._bd1_term_spread(df)
        indicators['BD2'] = self._bd2_credit_spread(df)
        indicators['BD3'] = self._bd3_bond_volatility(df)

        # Equity Market (EQ1, EQ2, EQ3)
        indicators['EQ1'] = self._eq1_kospi_return(df)
        indicators['EQ2'] = self._eq2_vkospi(df)
        indicators['EQ3'] = self._eq3_illiquidity(df)

        # FX Market (FX1, FX2, FX3)
        indicators['FX1'] = self._fx1_usdkrw_return(df)
        indicators['FX2'] = self._fx2_fx_volatility(df)
        indicators['FX3'] = self._fx3_crs_basis(df)

        # Financial Intermediaries (FI1, FI2, FI3)
        indicators['FI1'] = self._fi1_cds(df)
        indicators['FI2'] = self._fi2_fin_volatility(df)
        indicators['FI3'] = self._fi3_fin_relative_return(df)

        # ---------------------------------------------------------------
        # [수정] inf 방어 + 꼬리 결측 제한적 보정 (5영업일 한도)
        #   - 어떤 지표든 연산 과정에서 생긴 inf 를 NaN 으로 치환
        #   - 최근 5영업일 한도로만 ffill → 단기 결측은 carry,
        #     장기 결측(시리즈 사망)은 기존처럼 dropna 에 걸려 절단됨
        #     (그건 오히려 발견해야 할 신호이므로 의도된 동작)
        # ---------------------------------------------------------------
        indicators = indicators.replace([np.inf, -np.inf], np.nan)
        indicators = indicators.ffill(limit=5)

        cleaned = indicators.dropna()
        if cleaned.empty:
            non_null = indicators.notna().sum()
            empty_indicators = non_null[non_null == 0].index.tolist()
            first_valid = indicators.apply(lambda s: s.first_valid_index())
            last_valid = indicators.apply(lambda s: s.last_valid_index())
            details = [
                f"raw rows={len(df)}",
                f"raw date range={df.index.min()} ~ {df.index.max()}",
                f"indicator non-null counts={non_null.to_dict()}",
                f"all-NaN indicators={empty_indicators}",
                f"first valid={first_valid.to_dict()}",
                f"last valid={last_valid.to_dict()}",
            ]
            raise ValueError(
                "No complete indicator observations after dropna(). "
                "At least one required raw series is missing or the date overlap "
                "between series is empty. " + "; ".join(details)
            )

        return cleaned

    # =========================================================================
    # Money Market Indicators
    # =========================================================================
    def _mm1_cd_level(self, df: pd.DataFrame) -> pd.Series:
        """MM1: CD금리 수준 (level, stress=up)"""
        return df['KWCDC_Curncy']

    def _mm2_cd_ktb_spread(self, df: pd.DataFrame) -> pd.Series:
        """MM2: CD-국고채3M 스프레드 (spread, stress=up)"""
        return df['KWCDC_Curncy'] - df['GVSK3M_Index']

    def _mm3_cd_volatility(self, df: pd.DataFrame) -> pd.Series:
        """MM3: CD금리 변동성 (realized vol, stress=up)"""
        returns = df['KWCDC_Curncy'].diff()
        vol = returns.rolling(self.vol_window).std() * np.sqrt(252)
        return vol

    # =========================================================================
    # Bond Market Indicators
    # =========================================================================
    def _bd1_term_spread(self, df: pd.DataFrame) -> pd.Series:
        """BD1: 기간스프레드 10Y-3Y (spread, stress=down → invert)"""
        spread = df['GVSK10YR_Index'] - df['GVSK3YR_Index']
        # 역전(하락) = 스트레스 → 반전하여 양수가 스트레스
        return -spread

    def _bd2_credit_spread(self, df: pd.DataFrame) -> pd.Series:
        """BD2: CD-국고채3Y 스프레드 (spread, stress=up)"""
        return df['KWCDC_Curncy'] - df['GVSK3YR_Index']

    def _bd3_bond_volatility(self, df: pd.DataFrame) -> pd.Series:
        """BD3: MOVE Index (level, stress=up)"""
        return df['MOVE_Index']

    # =========================================================================
    # Equity Market Indicators
    # =========================================================================
    def _eq1_kospi_return(self, df: pd.DataFrame) -> pd.Series:
        """EQ1: KOSPI 주간수익률 (return, stress=down → invert)"""
        returns = df['KOSPI_Index'].pct_change()
        # 하락 = 스트레스 → 반전하여 양수가 스트레스
        return -returns

    def _eq2_vkospi(self, df: pd.DataFrame) -> pd.Series:
        """EQ2: VKOSPI (level, stress=up)"""
        return df['VKOSPI_Index']

    def _eq3_illiquidity(self, df: pd.DataFrame) -> pd.Series:
        """EQ3: Amihud 비유동성 (illiquidity, stress=up)

        [수정] 0-volume 방어 + inf 치환 + min_periods 완화
        """
        returns = df['KOSPI_Index'].pct_change().abs()
        # 0-volume(비거래일/결측 대체값) 방어: 0 → NaN → ffill
        volume = df['KOSPI_Index_VOLUME'].replace(0, np.nan).ffill()
        # Amihud = |return| / volume (scaled)
        illiq = (returns / volume) * 1e12  # 스케일 조정
        illiq = illiq.replace([np.inf, -np.inf], np.nan)
        # min_periods 완화: NaN 1~2개가 20일 윈도우를 오염시키지 않도록
        return illiq.rolling(self.vol_window, min_periods=15).mean()

    # =========================================================================
    # FX Market Indicators
    # =========================================================================
    def _fx1_usdkrw_return(self, df: pd.DataFrame) -> pd.Series:
        """FX1: USD/KRW 수익률 (return, stress=up, 원화약세)"""
        return df['USDKRW_Curncy'].pct_change()

    def _fx2_fx_volatility(self, df: pd.DataFrame) -> pd.Series:
        """FX2: FX 내재변동성 (level, stress=up)"""
        return df['USDKRWV1M_BGN_Curncy']

    def _fx3_crs_basis(self, df: pd.DataFrame) -> pd.Series:
        """FX3: CRS 베이시스 (level, stress=down → invert)"""
        # CRS가 낮을수록 달러 조달 어려움 (스트레스)
        # 여기서는 CRS 수준 자체를 사용 (낮은 값 = 스트레스)
        crs = df['KWSWNI1_Curncy']
        # 반전: 낮을수록 스트레스가 높음
        return -crs

    # =========================================================================
    # Financial Intermediaries Indicators
    # =========================================================================
    def _fi1_cds(self, df: pd.DataFrame) -> pd.Series:
        """FI1: 국가 CDS (level, stress=up)"""
        return df['CKREA1U5_CBGN_Curncy']

    def _fi2_fin_volatility(self, df: pd.DataFrame) -> pd.Series:
        """FI2: 금융업종 변동성 (realized vol, stress=up)"""
        returns = df['KOSPFIN_Index'].pct_change()
        vol = returns.rolling(self.vol_window).std() * np.sqrt(252) * 100
        return vol

    def _fi3_fin_relative_return(self, df: pd.DataFrame) -> pd.Series:
        """FI3: 금융업종 상대수익률 (relative return, stress=down → invert)"""
        fin_ret = df['KOSPFIN_Index'].pct_change()
        kospi_ret = df['KOSPI_Index'].pct_change()
        relative = fin_ret - kospi_ret
        # 금융업종 언더퍼폼 = 스트레스 → 반전
        return -relative


class ECDFTransformer:
    """ECDF (Empirical Cumulative Distribution Function) 변환"""

    def __init__(self):
        self.distributions = {}

    def fit(self, df: pd.DataFrame) -> 'ECDFTransformer':
        """각 지표별 분포 학습"""
        for col in df.columns:
            values = df[col].dropna().values
            self.distributions[col] = values
        return self

    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """ECDF 변환 적용 → [0, 1] 범위"""
        result = pd.DataFrame(index=df.index)

        for col in df.columns:
            if col in self.distributions:
                result[col] = df[col].apply(
                    lambda x: self._ecdf(x, self.distributions[col])
                )
            else:
                result[col] = df[col].rank(pct=True)

        return result

    def fit_transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """학습 + 변환"""
        return self.fit(df).transform(df)

    def _ecdf(self, x: float, sample: np.ndarray) -> float:
        """단일 값의 ECDF 계산"""
        if pd.isna(x):
            return np.nan
        return np.mean(sample <= x)


def compute_indicators(raw_data: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    원본 데이터 → 지표 변환 → ECDF 정규화

    Args:
        raw_data: 원본 데이터 (data_loader_v2 기준)

    Returns:
        indicators: 변환된 15개 지표 (원본 스케일)
        ecdf_indicators: ECDF 정규화된 지표 [0, 1]
    """
    # 지표 변환
    transformer = IndicatorTransformer()
    indicators = transformer.transform_all(raw_data)

    # ECDF 변환
    ecdf = ECDFTransformer()
    ecdf_indicators = ecdf.fit_transform(indicators)

    return indicators, ecdf_indicators


if __name__ == '__main__':
    from data_loader_v2 import load_raw_data

    daily, weekly = load_raw_data()
    indicators, ecdf_indicators = compute_indicators(weekly)

    print("\n[Raw Indicators]")
    print(indicators.tail())
    print("\n[ECDF Indicators (0-1)]")
    print(ecdf_indicators.tail())
    print("\n[Statistics]")
    print(ecdf_indicators.describe())