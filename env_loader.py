# -*- coding: utf-8 -*-
"""
env_loader.py — igis 공용 .env 로더 (의존성 없음)
=================================================
python-dotenv 없이 .env 파일을 읽어 os.environ에 주입한다.
igis 루트(또는 상위)에 .env 하나만 두면 모든 프로젝트가 공유한다.

사용법
------
    from env_loader import load_dotenv
    load_dotenv()                 # 자동 탐색 (호출한 파일 폴더 → 상위로)
    # 또는 모듈 import만 해도 자동 실행됨

.env 형식
---------
    FRED_API_KEY=xxxxxxxx
    ECOS_API_KEY=yyyyyyyy
    # 주석 가능
"""
import os


def load_dotenv(path: str | None = None, start_dir: str | None = None) -> str | None:
    """
    .env를 찾아 os.environ에 주입(이미 설정된 환경변수는 덮어쓰지 않음).

    Parameters
    ----------
    path : str | None
        명시적 .env 경로. 주면 그 파일만 사용.
    start_dir : str | None
        탐색 시작 폴더. 미지정 시 이 모듈 폴더에서 시작해 상위로 거슬러 올라감.

    Returns
    -------
    적용된 .env 경로(없으면 None).
    """
    candidates = []
    if path:
        candidates.append(path)
    else:
        d = start_dir or os.path.dirname(os.path.abspath(__file__))
        for _ in range(6):                       # 현재~최대 5단계 상위
            candidates.append(os.path.join(d, ".env"))
            parent = os.path.dirname(d)
            if parent == d:
                break
            d = parent

    for env_path in candidates:
        if not os.path.exists(env_path):
            continue
        try:
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, _, val = line.partition("=")
                    key, val = key.strip(), val.strip().strip('"').strip("'")
                    if key and key not in os.environ:   # 실제 환경변수 우선
                        os.environ[key] = val
        except Exception:
            pass
        return env_path
    return None


# import만 해도 자동 적용
load_dotenv()
