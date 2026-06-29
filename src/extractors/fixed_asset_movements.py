# -*- coding: utf-8 -*-
"""이상적 분개장 → 유무형자산 취득·처분 거래 추출.

⚠️ **프로젝트 원칙(이상적 양식 중간다리)**: 분개장은 raw로 바로 다루지 않는다. 회사제시 분개장을
`extractors/journal.parse_journal`로 **이상적 분개장 양식**({날짜,번호,계정과목,거래처,적요,차변,대변})
으로 먼저 변환(정산표 첫 매핑/조서생성 시 변환자료에 캐시)하고, **개별 조서(G 등)는 그 이상적
분개장에서 꺼내 쓴다.** 이 모듈은 그 이상적 라인을 받아 도메인 추출만 한다(헤더감지·포맷처리 X).

부호 규칙(처분 시 감가상각누계액이 −부호로 빠지는 것 포함):
  · 자산 본계정: 차변>0 = **취득**, 대변>0 = **처분(실제 처분원가)**.
  · 감가상각누계액 계정: 처분 시 반대분개로 **차변(감소)** 발생 → '누계감소'에 집계.
계정과목은 `[코드]계정명` 형식 → 코드 제거 후 '/' 앞부분이 자산계정명과 일치하면 채택.
"""

import re
from datetime import date, datetime


def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def _amt(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.replace(",", "").strip().strip("()")
        try:
            return float(t)
        except ValueError:
            return None
    return None


def _acct_name(raw) -> str:
    """'[20800]차량운반구' → '차량운반구', '[64606]합사운영비/지급임차료/..' → '합사운영비'."""
    t = re.sub(r"^\s*\[[^\]]*\]\s*", "", str(raw or ""))
    return t.split("/")[0].strip()


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        m = re.search(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", v)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
    return None


def extract_fixed_asset_movements(journal_lines, asset_names, *, accum_names=None) -> dict:
    """**이상적 분개장 라인**에서 asset_names 계정의 취득/처분 거래 추출.

    Args:
        journal_lines: parse_journal 결과 [{날짜,번호,계정과목,거래처,적요,차변,대변}, ...].
        asset_names: 자산 본계정명 집합(예: {'건물','차량운반구','소프트웨어'}).
        accum_names: 감가상각누계액 계정명 집합(처분 시 차변 감소 포착용, 선택).
    Returns:
        {계정명: {"취득":[line], "처분":[line], "취득합", "처분합", "누계감소"}}  (line={날짜,전표,적요,거래처,금액})
    """
    wants = {_norm(a): str(a) for a in asset_names}
    accum = {_norm(a): str(a) for a in (accum_names or [])}
    out: dict[str, dict] = {}

    def rec(name):
        return out.setdefault(name, {"취득": [], "처분": [], "취득합": 0.0,
                                     "처분합": 0.0, "누계감소": 0.0})

    for ln in journal_lines or []:
        name = _acct_name(ln.get("계정과목"))
        nn = _norm(name)
        is_asset, is_accum = nn in wants, nn in accum
        if not (is_asset or is_accum):
            continue
        dr = _amt(ln.get("차변")) or 0
        cr = _amt(ln.get("대변")) or 0
        if dr == 0 and cr == 0:
            continue
        if is_accum:
            rec(accum[nn])["누계감소"] += dr - cr
            continue
        line = {"날짜": _to_date(ln.get("날짜")), "전표": ln.get("번호"),
                "적요": ln.get("적요"), "거래처": ln.get("거래처")}
        r = rec(wants[nn])
        if dr > 0:
            r["취득"].append({**line, "금액": dr})
            r["취득합"] += dr
        if cr > 0:
            r["처분"].append({**line, "금액": cr})
            r["처분합"] += cr
    return out


def parse_fixed_asset_movements(source, asset_names, *, accum_names=None) -> dict:
    """편의 래퍼: source(분개장 또는 분개장 시트 가진 정산표)를 parse_journal로 이상적 변환 후 추출.

    원칙상 변환은 정산표 EXE가 미리 하고 캐시를 공유하는 게 정석(파이프라인은 load_with_cache로 분리 호출).
    이 래퍼는 단위 테스트·독립 호출 편의용.
    """
    from .journal import parse_journal
    return extract_fixed_asset_movements(parse_journal(source), asset_names, accum_names=accum_names)
