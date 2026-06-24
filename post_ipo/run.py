#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Post IPO Monitor - Main Entry Point
====================================
2년 이내 신규상장 종목 모니터링

비교일(전일자)은 수급.xlsx에서 값이 채워진 마지막 두 날짜로 자동 결정됩니다.
별도의 전일 결과 파일이나 입력이 필요 없습니다.

실행:
    python run.py             → 버전 A
    python run.py --b         → 버전 B (__post ipo univ.xlsx 사용)
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from screen_ipo import IPOMonitor
from post_ipo_daily import Config, setup_logging


def main():
    """메인 실행 함수
    실행 예시:
        python run.py        → 버전 A
        python run.py --b    → 버전 B
    """
    setup_logging()

    parser = argparse.ArgumentParser(description='Post IPO Monitor')
    parser.add_argument('--b', action='store_true', help='버전 B: __post ipo univ.xlsx 사용')
    args = parser.parse_args()

    source = 'B' if args.b else 'A'

    print("=" * 60)
    print("Post IPO Monitor")
    print(f"Universe 소스: {'__post ipo univ.xlsx (B)' if source == 'B' else '최초상장일.xlsx (A)'}")
    print("기준일/비교일: 수급.xlsx의 채워진 마지막 두 날짜로 자동 결정")
    print("=" * 60)

    config = Config()
    monitor = IPOMonitor(config)
    result = monitor.run(source=source)

    if result is not None and not result.empty:
        print(f"\n모니터링 완료: {len(result)}개 종목")
    else:
        print("\n모니터링 결과가 없습니다.")


if __name__ == "__main__":
    main()