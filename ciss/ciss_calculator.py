# -*- coding: utf-8 -*-
"""
CISS Risk Scoring - CISS Calculator Module
=========================================
Compute Composite Indicator of Systemic Stress (CISS).
"""

import pandas as pd
import numpy as np
from typing import Dict, Optional, List, Tuple


class CISSCalculator:
    """
    CISS (Composite Indicator of Systemic Stress) calculator.

    Formula:
        CISS_t = (w' * s_t) * sqrt(w' * C_t * w)

    where:
        - s_t: ECDF-normalized indicator vector
        - w: indicator weight vector
        - C_t: dynamic correlation matrix
    """

    SECTOR_INDICATORS = {
        "Money_Market": ["MM1", "MM2", "MM3"],
        "Bond_Market": ["BD1", "BD2", "BD3"],
        "Equity_Market": ["EQ1", "EQ2", "EQ3"],
        "FX_Market": ["FX1", "FX2", "FX3"],
        "Financial_Intermediaries": ["FI1", "FI2", "FI3"],
    }

    SECTOR_WEIGHTS = {
        "Money_Market": 0.20,
        "Bond_Market": 0.20,
        "Equity_Market": 0.20,
        "FX_Market": 0.20,
        "Financial_Intermediaries": 0.20,
    }

    def __init__(self, sector_weights: Optional[Dict[str, float]] = None):
        self.sector_weights = (sector_weights or self.SECTOR_WEIGHTS).copy()
        self._validate_weights()

    def _validate_weights(self) -> None:
        """Validate and normalize sector weights."""
        expected = set(self.SECTOR_INDICATORS.keys())
        actual = set(self.sector_weights.keys())
        if actual != expected:
            missing = sorted(expected - actual)
            extra = sorted(actual - expected)
            details = []
            if missing:
                details.append(f"missing keys: {missing}")
            if extra:
                details.append(f"unknown keys: {extra}")
            raise ValueError("Invalid sector_weights keys: " + ", ".join(details))

        total = float(sum(self.sector_weights.values()))
        if not np.isclose(total, 1.0):
            print(f"[WARN] Sector weights sum to {total}, normalizing to 1.0")
            for key in self.sector_weights:
                self.sector_weights[key] /= total

    def _build_indicator_layout(
        self, ecdf_data: pd.DataFrame
    ) -> Tuple[List[str], np.ndarray, Dict[str, np.ndarray]]:
        """
        Build indicator order, normalized indicator weights, and sector masks.
        Fails fast if required indicators are missing.
        """
        if ecdf_data.empty:
            raise ValueError("ecdf_data is empty; cannot compute CISS.")

        missing = []
        for sector, indicators in self.SECTOR_INDICATORS.items():
            for ind in indicators:
                if ind not in ecdf_data.columns:
                    missing.append(f"{sector}:{ind}")

        if missing:
            raise ValueError(
                "Missing required indicators for CISS calculation: "
                + ", ".join(missing)
            )

        indicator_order: List[str] = []
        indicator_weights: List[float] = []
        indicator_sectors: List[str] = []

        for sector, indicators in self.SECTOR_INDICATORS.items():
            weight_per_indicator = self.sector_weights[sector] / len(indicators)
            for ind in indicators:
                indicator_order.append(ind)
                indicator_weights.append(weight_per_indicator)
                indicator_sectors.append(sector)

        w = np.array(indicator_weights, dtype=float)
        w = w / w.sum()

        sector_masks: Dict[str, np.ndarray] = {}
        for sector in self.SECTOR_INDICATORS:
            sector_masks[sector] = np.array(
                [ind_sector == sector for ind_sector in indicator_sectors],
                dtype=bool,
            )

        return indicator_order, w, sector_masks

    def compute_sector_indices(self, ecdf_data: pd.DataFrame) -> pd.DataFrame:
        """
        Compute sector-level stress indices as simple means of indicators.
        Strictly requires all indicators to exist.
        """
        self._build_indicator_layout(ecdf_data)

        sector_indices = pd.DataFrame(index=ecdf_data.index)
        for sector, indicators in self.SECTOR_INDICATORS.items():
            sector_indices[sector] = ecdf_data[indicators].mean(axis=1)

        return sector_indices

    def compute_ciss(
        self, ecdf_data: pd.DataFrame, correlations: np.ndarray
    ) -> pd.DataFrame:
        """
        Compute full CISS using dynamic correlation matrices.

        Args:
            ecdf_data: ECDF-normalized indicators, shape (T, 15)
            correlations: Dynamic correlation matrices, shape (T, 15, 15)

        Returns:
            DataFrame with CISS, Correlation_Effect, and sector contributions.
        """
        indicator_order, w, sector_masks = self._build_indicator_layout(ecdf_data)
        ecdf_ordered = ecdf_data[indicator_order]

        T = len(ecdf_ordered)
        n = len(w)

        if correlations is None:
            raise ValueError("correlations cannot be None in compute_ciss().")
        if correlations.ndim != 3:
            raise ValueError(
                f"correlations must be 3D (T x N x N), got ndim={correlations.ndim}"
            )
        if correlations.shape[0] != T:
            raise ValueError(
                "Time dimension mismatch between ecdf_data and correlations: "
                f"T_ecdf={T}, T_corr={correlations.shape[0]}"
            )
        if correlations.shape[1] != n or correlations.shape[2] != n:
            raise ValueError(
                "Correlation matrix dimension mismatch: "
                f"expected ({n}, {n}), got "
                f"({correlations.shape[1]}, {correlations.shape[2]})."
            )

        ciss_values = np.zeros(T)
        correlation_effect = np.zeros(T)
        sector_contributions = {
            sector: np.zeros(T) for sector in self.SECTOR_INDICATORS
        }

        for t in range(T):
            s_t = ecdf_ordered.iloc[t].values
            C_t = correlations[t]

            weighted_stress = s_t * w
            portfolio_var = float(w @ C_t @ w)
            correlation_effect[t] = np.sqrt(max(portfolio_var, 0.0))

            ciss_values[t] = float(np.sum(weighted_stress) * correlation_effect[t])

            for sector, mask in sector_masks.items():
                sector_contributions[sector][t] = float(
                    np.sum(weighted_stress[mask]) * correlation_effect[t]
                )

        result = pd.DataFrame(index=ecdf_data.index)
        result["CISS"] = ciss_values
        result["Correlation_Effect"] = correlation_effect
        for sector in self.SECTOR_INDICATORS:
            result[f"{sector}_Contribution"] = sector_contributions[sector]

        return result

    def compute_simple_ciss(self, ecdf_data: pd.DataFrame) -> pd.DataFrame:
        """
        Compute simple CISS without correlation effects.
        """
        indicator_order, _, _ = self._build_indicator_layout(ecdf_data)
        ecdf_ordered = ecdf_data[indicator_order]

        sector_contributions = pd.DataFrame(index=ecdf_data.index)
        for sector, indicators in self.SECTOR_INDICATORS.items():
            sector_contributions[sector] = (
                ecdf_ordered[indicators].mean(axis=1) * self.sector_weights[sector]
            )

        result = pd.DataFrame(index=ecdf_data.index)
        result["CISS_Simple"] = sector_contributions.sum(axis=1)
        for sector in self.SECTOR_INDICATORS:
            result[f"{sector}_Contribution"] = sector_contributions[sector]

        return result


def compute_ciss_score(
    ecdf_data: pd.DataFrame,
    correlations: np.ndarray,
    use_correlation: bool = True,
) -> pd.DataFrame:
    """
    Convenience function for CISS score calculation.

    Args:
        ecdf_data: ECDF-normalized indicators
        correlations: Dynamic correlation matrices
        use_correlation: Use correlation amplification effect if True

    Returns:
        CISS result DataFrame
    """
    calculator = CISSCalculator()

    if use_correlation and correlations is not None:
        return calculator.compute_ciss(ecdf_data, correlations)
    return calculator.compute_simple_ciss(ecdf_data)


if __name__ == "__main__":
    from data_loader import load_raw_data
    from transforms import compute_indicators
    from dcc_garch import compute_dynamic_correlations

    daily, weekly = load_raw_data()
    _, ecdf_data = compute_indicators(weekly)
    correlations, _ = compute_dynamic_correlations(ecdf_data)

    ciss_result = compute_ciss_score(ecdf_data, correlations)
    print("\n[CISS Results]")
    print(ciss_result.tail(10))

    print("\n[Statistics]")
    print(ciss_result["CISS"].describe())
