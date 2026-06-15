"""취합엑셀 공통 — 번호 섹션 탐지 헬퍼.

BANK·INSURANCE·INVESTMENT 시트는 모두 같은 패턴을 쓴다:
  - 섹션 마커가 **col index 1**에만 위치, 텍스트가 `숫자(-숫자). 설명` 형태
    (예: "1. 조회기준일 현재...", "2-2. [대출거래의 내용]...", "9. ...")
  - 헤더 행은 마커 다음 몇 행 안에 위치, 데이터는 그 다음 ~ 다음 마커 전까지
"""

import re

# 섹션 마커는 col index 1 에만 존재
MARKER_COL = 1

# `숫자(-숫자). ` 뒤에 공백+텍스트가 와야 마커로 인정.
# (연이자율 '0.0000' 같은 순수 숫자를 마커로 오인하지 않도록)
_RE_SECTION = re.compile(r"^\s*\d+(-\d+)?\.\s+\S")


def section_marker(row: tuple) -> str | None:
    """행의 마커 열(col index 1)에 섹션 마커가 있으면 반환, 없으면 None."""
    if len(row) <= MARKER_COL:
        return None
    v = row[MARKER_COL]
    if isinstance(v, str) and _RE_SECTION.match(v):
        return v.strip()
    return None


def find_section_bounds(rows: list[tuple], prefix: str) -> tuple[int, int] | None:
    """`prefix`로 시작하는 섹션의 (마커행 인덱스, 끝 인덱스)를 반환.

    끝 인덱스 = 다음 섹션 마커 행 (없으면 len(rows)). 반환 구간은 [start, end).
    찾지 못하면 None.

    prefix 예: "1." → "1. ..." 매칭 ("10."은 매칭 안 됨).
              "2-2." → "2-2. ..." 매칭.
    """
    start = None
    for i, r in enumerate(rows):
        m = section_marker(r)
        if m and m.startswith(prefix):
            start = i
            break
    if start is None:
        return None

    end = len(rows)
    for i in range(start + 1, len(rows)):
        if section_marker(rows[i]):
            end = i
            break
    return start, end
