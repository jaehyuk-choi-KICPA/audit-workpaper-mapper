# -*- coding: utf-8 -*-
"""기말환율 메모(참고자료) 파서.

외화 계좌의 ENGAGEMENT 회신금액 = 기말환율 × 조회서 외화값 으로 환산하기 위해,
사용자가 [참고자료] 폴더에 적어둔 기말환율 메모(텍스트)를 읽는다.

메모 형식 (줄당 하나, '통화=환율'. '#' 주석/빈 줄 무시):
    USD=1434.90
    EUR=1521.30
환율은 '1 외화당 원화'. 통화코드는 대문자로 정규화한다.

키 없으면(메모 없음/빈 값) FX 환산은 비활성 — 외화 계좌도 조회서 KRW 금액을 그대로 쓴다.
"""

import glob
import re
from pathlib import Path


def parse_fx_memo(path: str) -> dict:
    """환율 메모 → {통화코드: 환율(float)}. 파싱 실패 줄은 건너뜀(예외 없음)."""
    rates: dict[str, float] = {}
    try:
        text = Path(path).read_text(encoding="utf-8")
    except Exception:
        try:
            text = Path(path).read_text(encoding="cp949")
        except Exception:
            return {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"([A-Za-z]{2,4})\s*[=:\t ]\s*([\d,]+\.?\d*)", line)
        if not m:
            continue
        try:
            rates[m.group(1).upper()] = float(m.group(2).replace(",", ""))
        except ValueError:
            pass
    return rates


def find_fx_memo(ref_dir: str) -> "str | None":
    """[참고자료] 폴더에서 환율 메모(.txt/.csv) 첫 파일을 찾는다. 없으면 None."""
    for pat in ("*환율*.txt", "*환율*.csv", "*fx*.txt", "*FX*.txt"):
        hits = [h for h in glob.glob(str(Path(ref_dir) / "**" / pat), recursive=True)
                if "~$" not in h]
        if hits:
            return sorted(hits)[0]
    return None


def load_fx_rates(ref_dir: str) -> dict:
    """[참고자료]에서 환율 메모를 찾아 {통화: 환율} 반환. 없으면 빈 dict(=FX 비활성)."""
    memo = find_fx_memo(ref_dir)
    return parse_fx_memo(memo) if memo else {}
