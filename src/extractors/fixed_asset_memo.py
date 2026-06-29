# -*- coding: utf-8 -*-
"""유무형자산 세부계정·내용연수·상각방법 메모(참고자료) 파서.

감가상각비 재계산(고정자산 관리대장_25)·회계정책 박스·주석은 자산별 **내용연수와
상각방법**을 알아야 한다. 회사 시스템에서 자동 추출이 어려운 판단 정보라, 인차지가
[참고자료] 폴더에 적어두는 메모(TXT 파이프 표)로 받는다.

메모 형식 (fx.py 메모 규약과 동일: '#'/빈 줄 무시, 줄당 자산 하나):
    # 구분 | 계정과목 | 내용연수 | 상각방법
    유형 | 건물 | 40 | 정액
    유형 | 차량운반구 | 5 | 정률
    유형 | 비품 | 5 | 정률,정액      # 병기 → 회사제시 금액 보고 추론
    무형 | 소프트웨어 | 5 | 정액

- 구분: '유형'/'무형'(앞글자만 맞으면 됨). 내용연수: 정수(년). '5년'처럼 단위 붙어도 됨.
- 상각방법: 정액/정률(여러 개면 쉼표/슬래시 구분 — 추론 대상). '정액법'처럼 '법' 붙어도 됨.
- 메모 없거나 빈 값이면 빈 리스트 → 재계산은 관리대장 상각율/회사제시 금액으로 폴백(생성기 책임).
"""

import glob
import re
from pathlib import Path


def _norm_method(token: str) -> "str | None":
    """'정액법'/'정률' 등 → '정액'/'정률'. 못 알아보면 None."""
    t = re.sub(r"\s+", "", token)
    if "정액" in t or "직선" in t or "straight" in t.lower():
        return "정액"
    if "정률" in t or "정율" in t or "체감" in t or "declin" in t.lower():
        return "정률"
    return None


def _norm_kind(token: str) -> str:
    """'유형자산'/'무형' → '유형'/'무형'. 기본 '유형'."""
    t = str(token)
    return "무형" if "무" in t else "유형"


def parse_memo(path: str) -> list[dict]:
    """메모 → [{구분, 계정과목, 내용연수:int|None, 상각방법:[정액|정률]}]. 실패 줄은 건너뜀."""
    try:
        text = Path(path).read_text(encoding="utf-8")
    except Exception:
        try:
            text = Path(path).read_text(encoding="cp949")
        except Exception:
            return []
    out: list[dict] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        # 구분자: 파이프 우선, 없으면 탭/다중공백
        parts = [p.strip() for p in (line.split("|") if "|" in line else re.split(r"\t|\s{2,}", line))]
        parts = [p for p in parts if p != ""]
        if len(parts) < 3:
            continue
        # 구분 컬럼이 생략된 줄(계정과목 | 내용연수 | 상각방법)도 허용
        if re.search(r"유형|무형", parts[0]) and len(parts) >= 4:
            kind, acct, years_tok, method_tok = parts[0], parts[1], parts[2], parts[3]
        else:
            kind, acct, years_tok, method_tok = "유형", parts[0], parts[1], parts[2]
        ym = re.search(r"\d+", years_tok)
        years = int(ym.group()) if ym else None
        methods = [m for m in (_norm_method(t) for t in re.split(r"[,/·]", method_tok)) if m]
        if not acct or (years is None and not methods):
            continue
        out.append({
            "구분": _norm_kind(kind), "계정과목": acct,
            "내용연수": years, "상각방법": methods or ["정액"],
        })
    return out


def find_memo(ref_dir: str) -> "str | None":
    """[참고자료]에서 유무형자산 메모(.txt) 첫 파일을 찾는다. 없으면 None."""
    for pat in ("*내용연수*.txt", "*유무형*.txt", "*유형자산*.txt", "*감가상각*.txt"):
        hits = [h for h in glob.glob(str(Path(ref_dir) / "**" / pat), recursive=True)
                if "~$" not in h]
        if hits:
            return sorted(hits)[0]
    return None


def load_memo(ref_dir: str) -> list[dict]:
    """[참고자료]에서 메모를 찾아 파싱. 없으면 빈 리스트."""
    memo = find_memo(ref_dir)
    return parse_memo(memo) if memo else []
