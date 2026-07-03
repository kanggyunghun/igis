# -*- coding: utf-8 -*-
"""
diag_date_cutoff.py — CISS 파이프라인 날짜 절단 원인 진단
=========================================================
ciss/ 폴더에 넣고 실행:  py -3.12 diag_date_cutoff.py

3단계 진단:
  [A] raw data 단계  : 소스별/컬럼별 마지막 유효 날짜 → 범인 소스 식별
  [B] indicator 단계 : 변환 후 지표별 마지막 유효 날짜 → transforms에서 잘리는지 확인
  [C] 캐시 파일 검사 : loader가 stale cache를 읽고 있는지 확인
"""
import os
import sys
import glob
from datetime import datetime, timedelta

import pandas as pd

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..'))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from env_loader import load_dotenv
load_dotenv()

from data_loader_v2 import load_raw_data_v2
from transforms import compute_indicators

START = '2024-01-01'
END = datetime.now().strftime('%Y-%m-%d')

SEP = "=" * 64


def last_valid_report(df: pd.DataFrame, title: str) -> pd.Series:
    """컬럼별 마지막 유효 날짜를 오래된 순으로 출력하고 반환."""
    print("\n" + SEP)
    print(f" {title}")
    print(SEP)
    lv = df.apply(lambda s: s.last_valid_index()).sort_values()
    overall_max = lv.max()
    for col, d in lv.items():
        if d is None or pd.isna(d):
            flag = "  ✗✗ 전체 NaN"
            dstr = "-"
        else:
            lag = (overall_max - d).days
            dstr = f"{pd.Timestamp(d):%Y-%m-%d}"
            flag = f"  ← ★ 범인 후보 ({lag}일 뒤처짐)" if lag >= 2 else ""
        print(f"  {str(col):42s} {dstr}{flag}")
    print(f"\n  >> 전체 최신 날짜   : {pd.Timestamp(overall_max):%Y-%m-%d}")
    print(f"  >> 가장 뒤처진 날짜 : {pd.Timestamp(lv.min()):%Y-%m-%d}")
    print(f"  >> dropna(how='any') 시 데이터는 '{pd.Timestamp(lv.min()):%Y-%m-%d}' 에서 끊김")
    return lv


def check_caches():
    """ciss/ 및 repo 루트에서 캐시로 의심되는 파일의 수정시각 검사."""
    print("\n" + SEP)
    print(" [C] 캐시 파일 검사 (수정시각 3일 이상 지난 csv/pkl/parquet)")
    print(SEP)
    patterns = ['*.csv', '*.pkl', '*.pickle', '*.parquet', '*.feather']
    found_stale = False
    for base in (_THIS_DIR, os.path.join(_THIS_DIR, 'cache'),
                 os.path.join(_THIS_DIR, 'data'),
                 os.path.join(_REPO_ROOT, 'outputs', 'ciss')):
        if not os.path.isdir(base):
            continue
        for pat in patterns:
            for f in glob.glob(os.path.join(base, pat)):
                mtime = datetime.fromtimestamp(os.path.getmtime(f))
                age = (datetime.now() - mtime).days
                mark = "  ← ★ stale 의심" if age >= 3 else ""
                if age >= 3:
                    found_stale = True
                print(f"  {os.path.relpath(f, _REPO_ROOT):55s} "
                      f"수정: {mtime:%Y-%m-%d %H:%M}{mark}")
    if not found_stale:
        print("  (3일 이상 지난 캐시성 파일 없음 — 캐시 문제 아님)")


def main():
    print(SEP)
    print(" CISS 날짜 절단 진단")
    print(f" 조회 기간: {START} ~ {END}")
    print(SEP)

    # [A] raw data
    print("\n[A] raw data 로드 중... (소스 API 호출, 1~2분 소요 가능)")
    daily, weekly = load_raw_data_v2(START, END)
    lv_raw = last_valid_report(daily, "[A] RAW DATA — 컬럼별 마지막 유효 날짜 (daily)")

    # [B] indicators
    print("\n[B] 지표 변환 중...")
    raw_ind, ecdf_ind = compute_indicators(daily)
    lv_ind = last_valid_report(raw_ind, "[B] INDICATORS — 변환 후 지표별 마지막 유효 날짜")

    # ecdf 최종 (파이프라인이 실제로 쓰는 것)
    print("\n" + SEP)
    print(" [B-2] ECDF 최종 인덱스 (DCC/CISS에 실제 투입되는 날짜 범위)")
    print(SEP)
    print(f"  first: {ecdf_ind.index.min():%Y-%m-%d}")
    print(f"  last : {ecdf_ind.index.max():%Y-%m-%d}")
    print(f"  rows : {len(ecdf_ind)}")

    # [C] cache
    check_caches()

    # 판정
    print("\n" + SEP)
    print(" 판정 가이드")
    print(SEP)
    raw_min, raw_max = lv_raw.min(), lv_raw.max()
    ind_max = lv_ind.max()
    ecdf_max = ecdf_ind.index.max()

    if (raw_max - raw_min).days >= 2:
        laggards = lv_raw[lv_raw < raw_max - timedelta(days=1)].index.tolist()
        print(f"  → 원인 1 (소스 lag): raw 단계에서 이미 컬럼 간 날짜 차이 존재.")
        print(f"    뒤처진 컬럼: {laggards}")
        print(f"    해결: data_loader_v2 에서 해당 시리즈 ffill, 또는")
        print(f"          compute_indicators 진입 전 daily.ffill() 적용.")
    if pd.Timestamp(ind_max) < pd.Timestamp(raw_max):
        print(f"  → 원인 2 (transforms 절단): raw 최신은 {pd.Timestamp(raw_max):%m-%d}인데")
        print(f"    indicator 최신이 {pd.Timestamp(ind_max):%m-%d}. transforms 내부의")
        print(f"    dropna/정렬/join 로직 확인 필요.")
    if pd.Timestamp(ecdf_max) < pd.Timestamp(ind_max):
        print(f"  → 원인 3 (ECDF/정합 절단): indicator 최신 {pd.Timestamp(ind_max):%m-%d} 대비")
        print(f"    ecdf 최신 {pd.Timestamp(ecdf_max):%m-%d}. compute_indicators 말미의")
        print(f"    dropna(how='any') 가 유력.")
    if (raw_max - raw_min).days < 2 and ecdf_max == ind_max == raw_max:
        print("  → raw/indicator/ecdf 모두 같은 날짜까지 존재.")
        print("    소스 API 자체가 그 날짜까지만 제공 중이거나, 캐시 문제([C] 확인).")

    print(SEP)


if __name__ == '__main__':
    main()