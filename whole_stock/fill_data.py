#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fill_data.py — 거래일 공백 보강 도구

whole_stock.py 가 "거래일 공백"을 감지해 중단했을 때, 빠진 날짜의 데이터를
수동으로 채워 전종목_수급.xlsx 의 stack 시트에 직접 보강한다.

준비물 (whole_stock.py 와 같은 폴더):
    fill_data.xlsx
        - 공백 날짜 개수만큼 시트를 만든다.
        - 시트명 = 날짜(YYYYMMDD).  예: 20260624, 20260625
        - 각 시트의 구성은 전종목_수급.xlsx 의 'update' 시트와 동일
          (상단 메타/아이템 라벨 행 + 종목코드(A######) 행부터 데이터).
          ※ update 시트를 그대로 복사해 값만 해당 날짜 것으로 바꾸면 가장 안전하다.

동작:
    1) fill_data.xlsx 의 각 시트를 읽어 시트명(YYYYMMDD)을 기준일로 파싱
    2) update 시트와 동일 포맷으로 해석(같은 로더 재사용) → tidy long
    3) 전종목_수급.xlsx 의 기존 stack 시트와 합치고 (날짜,코드,항목) 중복 제거
    4) 합쳐진 누적을 전종목_수급.xlsx 의 stack 시트에 직접 다시 기록
       (보강분도 일반 데이터로 합쳐짐 — 별도 표시/마커 없음)

실행:
    py -3.12 fill_data.py

주의:
    - 전종목_수급.xlsx 가 Excel에서 '닫혀' 있어야 stack 시트 기록이 가능하다.
    - 보강 후 whole_stock.py 를 다시 실행하면 공백이 메워져 정상 진행된다.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

# whole_stock.py 의 파싱/기록 로직을 그대로 재사용 (같은 폴더에 있어야 함)
from whole_stock import (
    INPUT_FILE,
    STACK_SHEET,
    KEEP_TRADING_DAYS,
    load_update_sheet,
    load_stack_sheet,
    accumulate_and_persist,
    write_stack_into_workbook,
    STACK_FILE,
)

BASE_DIR = Path(__file__).resolve().parent
FILL_FILE = BASE_DIR / "fill_data.xlsx"

_SHEET_DATE_RE = re.compile(r"^\s*(\d{8})\s*$")  # 20260624 형식


def _parse_sheet_date(sheet_name: str) -> pd.Timestamp | None:
    """시트명(YYYYMMDD)을 날짜로 파싱. 형식이 안 맞으면 None."""
    m = _SHEET_DATE_RE.match(str(sheet_name))
    if not m:
        return None
    dt = pd.to_datetime(m.group(1), format="%Y%m%d", errors="coerce")
    return None if pd.isna(dt) else pd.Timestamp(dt).normalize()


def main() -> None:
    print("=" * 60)
    print("fill_data — 거래일 공백 보강")
    print("=" * 60)

    if not FILL_FILE.exists():
        raise FileNotFoundError(
            f"{FILL_FILE} 가 없습니다. whole_stock.py 와 같은 폴더에 "
            f"fill_data.xlsx 를 만들고 공백 날짜 시트를 채워주세요."
        )
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"{INPUT_FILE} 를 찾을 수 없습니다.")

    # --- 1) fill_data.xlsx 의 시트 목록 → 날짜 파싱 ---
    xls = pd.ExcelFile(FILL_FILE)
    dated_sheets: list[tuple[pd.Timestamp, str]] = []
    skipped = []
    for sheet in xls.sheet_names:
        d = _parse_sheet_date(sheet)
        if d is None:
            skipped.append(sheet)
        else:
            dated_sheets.append((d, sheet))

    if skipped:
        print(f"  ⚠️ 날짜 형식(YYYYMMDD)이 아닌 시트는 건너뜁니다: {skipped}")
    if not dated_sheets:
        raise RuntimeError(
            "fill_data.xlsx 에 YYYYMMDD 형식의 시트가 없습니다. "
            "예: 20260624 시트를 만들어 주세요."
        )

    dated_sheets.sort(key=lambda x: x[0])  # 날짜 오름차순
    print(f"\n[1] 보강 대상 시트 {len(dated_sheets)}개")
    for d, sheet in dated_sheets:
        print(f"    - {sheet}  → {d:%Y-%m-%d}")

    # --- 2) 각 시트를 update 포맷으로 파싱 (시트명 날짜를 기준일로 강제) ---
    print("\n[2] 시트 파싱 (update 포맷)")
    fill_parts = []
    for d, sheet in dated_sheets:
        try:
            long_df, used_date = load_update_sheet(FILL_FILE, sheet, date_override=d)
        except Exception as e:
            print(f"    ✗ [{sheet}] 파싱 실패: {e}")
            continue
        n_codes = long_df["code"].nunique()
        if n_codes == 0:
            print(f"    ✗ [{sheet}] 유효 종목 0개 → 건너뜀")
            continue
        fill_parts.append(long_df)
        print(f"    ✓ [{sheet}] {used_date:%Y-%m-%d}  {n_codes:,}종목")

    if not fill_parts:
        raise RuntimeError("파싱된 보강 데이터가 없습니다. 시트 구성을 확인하세요.")

    fill_long = pd.concat(fill_parts, ignore_index=True)

    # --- 3) 기존 stack 읽기 (워크북 stack 시트 ∪ 별도 누적파일) ---
    print("\n[3] 기존 누적 로드")
    stack_parts = []
    try:
        stack_parts.append(load_stack_sheet(INPUT_FILE, STACK_SHEET))
        print(f"    워크북 stack 시트 로드")
    except Exception as e:
        print(f"    워크북 stack 시트 없음/실패({e}) → 무시")
    if STACK_FILE.exists():
        try:
            stack_parts.append(load_stack_sheet(STACK_FILE, "stack"))
            print(f"    별도 누적파일 로드: {STACK_FILE.name}")
        except Exception as e:
            print(f"    별도 누적파일 로드 실패({e}) → 무시")
    stack_long = (pd.concat(stack_parts, ignore_index=True)
                  if stack_parts else
                  pd.DataFrame(columns=["date", "code", "name", "field", "value"]))

    # --- 4) 합치기(중복제거+trim) + 별도 누적파일 저장 ---
    print("\n[4] 보강분 병합")
    merged_long = accumulate_and_persist(stack_long, fill_long, KEEP_TRADING_DAYS, STACK_FILE)

    # --- 5) 전종목_수급.xlsx 의 stack 시트에 직접 기록 ---
    print("\n[5] 전종목_수급.xlsx stack 시트 보강 기록")
    ok = write_stack_into_workbook(merged_long, INPUT_FILE, STACK_SHEET)
    if ok:
        filled_dates = ", ".join(f"{d:%Y-%m-%d}" for d, _ in dated_sheets)
        print("\n" + "=" * 60)
        print(f"✓ 보강 완료: {filled_dates}")
        print("  이제 whole_stock.py 를 다시 실행하면 정상 진행됩니다.")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("⚠️ 워크북 stack 시트 기록에 실패했습니다.")
        print("  전종목_수급.xlsx 가 Excel에서 열려 있지 않은지 확인 후 다시 실행하세요.")
        print(f"  (별도 누적파일에는 저장됨: {STACK_FILE.name})")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    main()