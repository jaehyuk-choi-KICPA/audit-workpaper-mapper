"""금융기관조회서 취합엑셀 BANK 시트 파서.

은행 회신 1번 항목(금융상품의 내용)에서 계좌별 잔액을 추출한다.
금융상품의 종류로 현금성자산 / 퇴직연금을 분류한다:
  - 퇴직연금 → A300 (기타조회내역)
  - 그 외     → A200 (현금성자산)

시트 구조 (회사 무관, 표준 양식):
  - 섹션 마커: A열(또는 인접) 텍스트가 `숫자.` / `숫자-숫자.` 로 시작
    1번 = 금융상품, 2-1/2-2 = 대출, 9번 = 담보 ...
  - 1번 섹션: [마커행] → [+2행 헤더] → [+3행~ 데이터] → [다음 마커 전까지]
"""

import re
from pathlib import Path

import openpyxl

from ._sections import find_section_bounds

# 헤더 동의어 → 표준 컬럼명
_SYNONYMS: dict[str, str] = {
    "조서번호":      "조서번호",
    "금융기관명":    "금융기관명",
    "금융상품의 종류": "금융상품종류",
    "금융상품의종류": "금융상품종류",
    "계좌번호":      "계좌번호",
    "금액":          "금액",
    "연이자율":      "연이자율",
    "만기일":        "만기일",
    "금액_통화":     "통화",
    "금액_금액":     "외화금액",
}

# 1번 섹션 헤더로 인정하기 위한 최소 필수 컬럼
_REQUIRED = {"금융기관명", "금융상품종류", "계좌번호", "금액"}

# 만기일 sentinel — 날짜가 아닌 값은 수시입출금으로 간주
_MATURITY_NONE = {"00000000", "00010101", "0", ""}

# 퇴직연금 분류 키워드 (금융상품종류 컬럼값)
_PENSION_KEYWORD = "퇴직연금"

_DEFAULT_SHEET = "BANK"


def _map_headers(row: tuple) -> dict[str, int]:
    """행에서 {표준컬럼명: 열인덱스} 매핑. 필수 컬럼 없으면 빈 dict."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(row):
        if cell is None:
            continue
        std = _SYNONYMS.get(str(cell).strip())
        if std and std not in mapping:
            mapping[std] = idx
    return mapping if _REQUIRED.issubset(mapping.keys()) else {}


def _format_maturity(value) -> str:
    """만기일 셀값을 표준화. 무만기 sentinel → '수시입출금'."""
    if value is None:
        return "수시입출금"
    s = str(value).strip()
    if s in _MATURITY_NONE:
        return "수시입출금"
    # YYYYMMDD 형태면 YYYY-MM-DD 로
    if re.fullmatch(r"\d{8}", s):
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s


def _classify(product_kind) -> str:
    """금융상품종류 → 계정분류."""
    if product_kind and _PENSION_KEYWORD in str(product_kind):
        return "퇴직연금"
    return "현금성자산"


def _to_record(row: tuple, cm: dict[str, int]) -> dict:
    def g(col: str):
        idx = cm.get(col)
        return row[idx] if idx is not None and idx < len(row) else None

    kind = g("금융상품종류")
    return {
        "조서번호":     g("조서번호"),
        "금융기관명":   g("금융기관명"),
        "금융상품종류": kind,
        "계좌번호":     g("계좌번호"),
        "금액":         g("금액"),
        "연이자율":     g("연이자율"),
        "만기":         _format_maturity(g("만기일")),
        "통화":         g("통화"),
        "외화금액":     g("외화금액"),
        "계정분류":     _classify(kind),
    }


def parse_bank(path: str, sheet_name: str = _DEFAULT_SHEET) -> list[dict]:
    """BANK 시트 1번 항목을 파싱해 계좌별 표준 행 리스트를 반환한다.

    Returns:
        [{"조서번호","금융기관명","금융상품종류","계좌번호","금액",
          "연이자율","만기","통화","외화금액","계정분류"}, ...]
        계정분류 = "현금성자산" | "퇴직연금"

    Raises:
        ValueError: 시트 없음 / 1번 섹션·헤더 미발견 / 데이터 없음
    """
    p = Path(path)
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"[BANK 파서] 시트 '{sheet_name}' 없음. 존재: {wb.sheetnames}")

    rows = list(wb[sheet_name].iter_rows(values_only=True))
    wb.close()

    # 1번 섹션 경계 [start, end) 탐색
    bounds = find_section_bounds(rows, "1.")
    if bounds is None:
        raise ValueError("[BANK 파서] 1번 섹션(금융상품) 마커를 찾을 수 없습니다.")
    sec1_idx, sec1_end = bounds

    # 마커 다음 행들 중 헤더 행 탐색 (필수 컬럼 모두 포함하는 첫 행)
    header_idx, col_map = None, None
    for i in range(sec1_idx + 1, min(sec1_idx + 6, sec1_end)):
        cm = _map_headers(rows[i])
        if cm:
            header_idx, col_map = i, cm
            break
    if col_map is None:
        raise ValueError("[BANK 파서] 1번 섹션 헤더를 찾을 수 없습니다.")

    # 데이터: 헤더 다음 ~ 섹션 끝(다음 마커 전)
    result = []
    for i in range(header_idx + 1, sec1_end):
        row = rows[i]
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        # 금융기관명 없는 행(주석·합계 등) skip
        inst_idx = col_map["금융기관명"]
        inst = row[inst_idx] if inst_idx < len(row) else None
        if inst is None or str(inst).strip() == "":
            continue
        result.append(_to_record(row, col_map))

    if not result:
        raise ValueError(f"[BANK 파서] 1번 섹션 데이터 행 없음: {path}")

    return result
