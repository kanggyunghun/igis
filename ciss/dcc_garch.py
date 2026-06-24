# -*- coding: utf-8 -*-
"""
CISS Risk Scoring - DCC-GARCH Module
=====================================
Dynamic Conditional Correlation 추정
"""

import pandas as pd
import numpy as np
from typing import Tuple, Optional
import warnings
warnings.filterwarnings('ignore')

# arch 라이브러리 사용 시도
try:
    from arch import arch_model
    from arch.univariate import ConstantMean, GARCH
    HAS_ARCH = True
except ImportError:
    HAS_ARCH = False
    print("[WARN] arch library not installed. Using EWMA fallback.")


class DCCEstimator:
    """
    DCC-GARCH 동적 상관행렬 추정

    CISS 원본 방법론:
    1. 각 지표에 GARCH(1,1)로 조건부 분산 추정
    2. 표준화 잔차로 DCC 상관행렬 추정
    3. 시변 상관행렬을 CISS 계산에 사용
    """

    def __init__(self, method: str = 'auto'):
        """
        Args:
            method: 'dcc' (full DCC-GARCH), 'ewma' (exponential weighted), 'auto'
        """
        valid_methods = {'auto', 'dcc', 'ewma'}
        if method not in valid_methods:
            raise ValueError(
                f"Invalid method '{method}'. Choose one of: {sorted(valid_methods)}."
            )

        if method == 'auto':
            self.method = 'dcc' if HAS_ARCH else 'ewma'
        elif method == 'dcc' and not HAS_ARCH:
            raise ImportError(
                "method='dcc' requested but 'arch' library is not installed. "
                "Install it with `pip install arch` or set method='ewma'."
            )
        else:
            self.method = method

        self.garch_models = {}
        self.conditional_vols = None
        self.standardized_residuals = None

    def fit(self, df: pd.DataFrame) -> 'DCCEstimator':
        """
        DCC-GARCH 모델 적합

        Args:
            df: ECDF 변환된 지표 데이터
        """
        if df is None or df.empty:
            raise ValueError(
                "ecdf_data is empty; cannot estimate dynamic correlations. "
                "Check indicator generation and raw data coverage first."
            )

        if self.method == 'dcc':
            if not HAS_ARCH:
                raise RuntimeError(
                    "method='dcc' cannot run because 'arch' is unavailable."
                )
            return self._fit_dcc(df)

        return self._fit_ewma(df)

    def get_correlation_matrix(self, t: int = -1) -> pd.DataFrame:
        """
        특정 시점의 상관행렬 반환

        Args:
            t: 시점 인덱스 (-1 = 최신)
        """
        if self.correlations is None:
            raise ValueError("Model not fitted. Call fit() first.")

        return self.correlations[t]

    def get_all_correlations(self) -> np.ndarray:
        """모든 시점의 상관행렬 반환 (T x N x N)"""
        return self.correlations

    def _fit_dcc(self, df: pd.DataFrame) -> 'DCCEstimator':
        """Full DCC-GARCH 추정 (arch 라이브러리 사용)"""
        print("[INFO] Fitting DCC-GARCH model...")

        n_assets = df.shape[1]
        T = df.shape[0]
        cols = df.columns.tolist()

        # Step 1: 각 시계열에 GARCH(1,1) 적합
        standardized = pd.DataFrame(index=df.index, columns=cols)
        conditional_vols = pd.DataFrame(index=df.index, columns=cols)

        for col in cols:
            series = df[col].dropna() * 100  # 스케일 조정

            try:
                model = arch_model(series, vol='Garch', p=1, q=1, mean='Constant')
                result = model.fit(disp='off')

                # 조건부 분산 및 표준화 잔차
                cond_vol = result.conditional_volatility
                resid = result.resid / cond_vol

                conditional_vols[col] = cond_vol
                standardized[col] = resid
                self.garch_models[col] = result

            except Exception as e:
                print(f"  [WARN] GARCH failed for {col}: {e}")
                # 폴백: 단순 표준화
                standardized[col] = (series - series.mean()) / series.std()
                conditional_vols[col] = series.rolling(20).std()

        self.conditional_vols = conditional_vols
        self.standardized_residuals = standardized.dropna()

        # Step 2: DCC 상관행렬 추정
        self.correlations = self._estimate_dcc_correlations(
            self.standardized_residuals.values
        )

        print(f"[INFO] DCC-GARCH fitted: {T} periods, {n_assets} assets")
        return self

    def _fit_ewma(self, df: pd.DataFrame, lambda_: float = 0.94) -> 'DCCEstimator':
        """EWMA 기반 동적 상관행렬 (간단한 대안)"""
        print(f"[INFO] Fitting EWMA correlation (lambda={lambda_})...")

        data = df.dropna().values
        T, n = data.shape
        cols = df.columns.tolist()
        if T < 2:
            raise ValueError(
                f"Need at least 2 complete ECDF observations for EWMA correlation; got {T}."
            )

        # EWMA 공분산 행렬 계산
        correlations = np.zeros((T, n, n))

        # 초기 상관행렬 (무조건부)
        cov_init = np.cov(data.T)
        std_init = np.sqrt(np.diag(cov_init))
        corr_init = cov_init / np.outer(std_init, std_init)
        correlations[0] = corr_init

        # EWMA 업데이트
        cov_t = cov_init.copy()
        for t in range(1, T):
            outer_t = np.outer(data[t-1], data[t-1])
            cov_t = lambda_ * cov_t + (1 - lambda_) * outer_t

            # 상관행렬로 변환
            std_t = np.sqrt(np.diag(cov_t))
            std_t[std_t == 0] = 1e-10  # 0 방지
            corr_t = cov_t / np.outer(std_t, std_t)

            # 범위 클리핑
            np.fill_diagonal(corr_t, 1.0)
            corr_t = np.clip(corr_t, -1, 1)

            correlations[t] = corr_t

        self.correlations = correlations
        self.conditional_vols = df.rolling(20).std()
        self.standardized_residuals = (df - df.mean()) / df.std()

        print(f"[INFO] EWMA fitted: {T} periods, {n} assets")
        return self

    def _estimate_dcc_correlations(self, std_resid: np.ndarray,
                                   alpha: float = 0.05,
                                   beta: float = 0.93) -> np.ndarray:
        """
        DCC 상관행렬 추정

        Q_t = (1-alpha-beta)*Q_bar + alpha*(e_{t-1}*e_{t-1}') + beta*Q_{t-1}
        R_t = diag(Q_t)^{-1/2} * Q_t * diag(Q_t)^{-1/2}
        """
        T, n = std_resid.shape

        # 무조건부 상관행렬
        Q_bar = np.corrcoef(std_resid.T)

        # DCC 업데이트
        correlations = np.zeros((T, n, n))
        Q_t = Q_bar.copy()

        for t in range(T):
            if t > 0:
                e_t = std_resid[t-1].reshape(-1, 1)
                Q_t = (1 - alpha - beta) * Q_bar + alpha * (e_t @ e_t.T) + beta * Q_t

            # Q_t → R_t (정규화)
            Q_diag = np.sqrt(np.diag(Q_t))
            Q_diag[Q_diag == 0] = 1e-10
            R_t = Q_t / np.outer(Q_diag, Q_diag)

            # 범위 클리핑
            np.fill_diagonal(R_t, 1.0)
            R_t = np.clip(R_t, -1, 1)

            correlations[t] = R_t

        return correlations


def compute_dynamic_correlations(ecdf_data: pd.DataFrame,
                                 method: str = 'auto') -> Tuple[np.ndarray, DCCEstimator]:
    """
    동적 상관행렬 계산

    Args:
        ecdf_data: ECDF 변환된 지표 데이터
        method: 'dcc', 'ewma', 'auto'

    Returns:
        correlations: (T, N, N) 상관행렬 배열
        estimator: 적합된 DCCEstimator 객체
    """
    estimator = DCCEstimator(method=method)
    estimator.fit(ecdf_data)

    return estimator.get_all_correlations(), estimator


if __name__ == '__main__':
    from data_loader import load_raw_data
    from transforms import compute_indicators

    daily, weekly = load_raw_data()
    indicators, ecdf_data = compute_indicators(weekly)

    correlations, estimator = compute_dynamic_correlations(ecdf_data)

    print(f"\nCorrelation matrix shape: {correlations.shape}")
    print("\n[Latest Correlation Matrix]")
    print(pd.DataFrame(
        correlations[-1],
        index=ecdf_data.columns,
        columns=ecdf_data.columns
    ).round(3))
