# -*- coding: utf-8 -*-
"""거래처원장 이상적양식 → 차입금 계정별 거래처(은행) 잔액 추출 (BBDD RECAP·200 좌측용).

이상적 분개/원장 중간다리(`ledger.read_ideal_ledger`, _IDEAL_HEADER)에서 차입금 계정
(단기/유동성장기/장기차입금)만 골라 **시설(거래처)별 기초(전기이월)·기말(잔액)** 을 뽑는다.
- 거래처명은 "기업은행(중소기업시설자금대출)"처럼 기관명+괄호(대출종류/계좌) → 기관명만 정규화
  (이미 '우리은행'처럼 깨끗하면 그대로). 괄호 내용은 대출종류로 부수 추출.
- 기초·기말이 둘 다 0인 행(한도-only)은 제외.
"""

import re

# 계정명 키워드 → RECAP 계정분류. '유동성장기'를 '장기차입금'보다 먼저(부분일치 오선점 방지).
_CLASS = [
    ("유동성장기", "유동성장기부채"),
    ("단기차입금", "단기차입금"),
    ("장기차입금", "장기차입금"),
]


def _clean_inst(s):
    """기관명 정규화: 괄호(계좌/종류) 제거 + 공백 제거. 이미 깨끗하면 그대로."""
    if s is None:
        return None
    s = re.sub(r"\(.*?\)", "", str(s))
    s = re.sub(r"\s+", "", s).strip()
    return s or None


def _paren(s):
    m = re.search(r"\((.*?)\)", str(s or ""))
    return m.group(1).strip() if m else None


def _amt(v):
    if v is None or isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def _classify(acct):
    a = acct or ""
    for kw, cls in _CLASS:
        if kw in a:
            return cls
    return None


def borrowings_by_inst(ideal_rows: list) -> list:
    """이상적 원장 행 → [{금융기관, 대출종류, 계정분류, 기초, 기말}] (차입금 시설별)."""
    out = []
    for r in ideal_rows:
        cls = _classify(r.get("계정과목"))
        if not cls:
            continue
        기초, 기말 = _amt(r.get("전기이월")), _amt(r.get("잔액"))
        if 기초 == 0 and 기말 == 0:
            continue
        out.append({
            "금융기관": _clean_inst(r.get("거래처명")),
            "대출종류": _paren(r.get("거래처명")),
            "계정분류": cls,
            "기초": 기초,
            "기말": 기말,
        })
    return out
