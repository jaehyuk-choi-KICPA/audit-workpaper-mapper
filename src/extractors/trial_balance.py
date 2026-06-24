"""별도정산표(trial balance) 파서.

정산표 파일의 **별도정산표 시트만** 사용한다. 이 시트는 계정과목별 1행으로
기초(전기말 수정후) · 기말(당기 회사제시) · 당기수정(차변/대변) · 수정후 잔액을 담는다.
(당기수정 차변/대변은 별도정산표가 이미 '수정사항집계' 시트를 SUMIF로 물고 있어 부호가
정산표 규약대로 계산돼 있다 → 코드는 그 값을 그대로 읽는다.)

출력(이상적 양식 records):
  계정명 | 대분류 | 기초 | 기말 | 수정차변 | 수정대변 | 수정사항 | 수정후
  - 계정명  = 회사사용 계정과목(별도정산표 B열) → 총괄표 '회사사용 계정과목'
  - 대분류  = 공시용 계정과목(별도정산표 C열) → 총괄표 '대분류'(계정→조서 라우팅 키)
  - 수정사항 = 수정차변 - 수정대변 (편의 필드; 부호는 정산표 규약 그대로)

헤더 라벨이 회사/연도마다 조금씩 달라(날짜가 헤더에 들어감) 위치 탐지는 **안정적 구조 라벨**
('공시용'·'차변'·'대변'·'Audited'·'Unaudited')에 앵커한다. 셀값이 없거나 시트가 없으면
예외 대신 빈 리스트를 반환한다(한 파서의 빈 결과가 전체 생성을 막지 않도록).
"""

import warnings
from pathlib import Path

import openpyxl

from ._headers import normalize
from .settlement import _is_subtotal

_DEFAULT_SHEET = "별도정산표"


def _is_amount(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _find_header(rows: list[tuple]) -> "tuple[int, dict] | None":
    """헤더 행 인덱스와 {필드: 열인덱스} 매핑을 찾는다(0-indexed).

    별도정산표는 전기/당기 두 블록이 있어 '차변'/'대변'과 'Audited'가 각각 2개씩 나온다.
    **당기** 블록을 '기말(회사제시)=Unaudited' 열에 앵커해 잡는다:
      기초 = Unaudited 좌측 마지막 Audited / 당기수정 차변·대변 = Unaudited 우측 첫 차변·대변 /
      수정후 = 당기 대변 우측 첫 Audited / 대분류 = '공시용' / 계정명 = 대분류 바로 왼쪽.
    """
    def has(n, key):
        return key.lower() in n.lower()

    for i, row in enumerate(rows):
        norms = [normalize(c) for c in row]
        col_대분류 = next((j for j, n in enumerate(norms) if "공시용" in n), None)
        col_기말 = next((j for j, n in enumerate(norms) if has(n, "unaudited")), None)
        if col_대분류 is None or col_기말 is None:
            continue
        col_차변 = next((j for j, n in enumerate(norms) if n == "차변" and j > col_기말), None)
        col_대변 = next((j for j, n in enumerate(norms) if n == "대변" and j > col_기말), None)
        audited = [j for j, n in enumerate(norms)
                   if has(n, "audited") and not has(n, "unaudited")]
        col_기초 = max((j for j in audited if j < col_기말), default=None)
        col_수정후 = next((j for j in audited if col_대변 is not None and j > col_대변), None)
        col_계정명 = col_대분류 - 1 if col_대분류 > 0 else None
        cm = {
            "계정명": col_계정명, "대분류": col_대분류,
            "기초": col_기초, "기말": col_기말,
            "수정차변": col_차변, "수정대변": col_대변, "수정후": col_수정후,
        }
        if all(v is not None for v in cm.values()):
            return i, cm
    return None


def parse_trial_balance(path: str, sheet_name: str = _DEFAULT_SHEET) -> list[dict]:
    """별도정산표에서 계정과목별 행을 추출한다.

    Returns:
        [{"계정명","대분류","기초","기말","수정차변","수정대변","수정사항","수정후"}, ...]
        (시트/헤더 없으면 빈 리스트)
    """
    p = Path(path)
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        warnings.warn(f"[별도정산표 파서] 시트 '{sheet_name}' 없음. 존재: {wb.sheetnames}")
        return []
    rows = list(wb[sheet_name].iter_rows(values_only=True))
    wb.close()

    found = _find_header(rows)
    if found is None:
        warnings.warn("[별도정산표 파서] 헤더 행(공시용/차변/대변)을 찾지 못함.")
        return []
    header_idx, cm = found

    def g(row, key):
        j = cm[key]
        return row[j] if j is not None and j < len(row) else None

    result = []
    # 판매비와관리비 섹션 추적: 그 안의 상세행에 판관비=True 플래그. R(판관비) 총괄표가 회사마다
    # 다른 계정명(하자보수비·잡비·접대비·세금과공과 등)을 명시목록 없이 섹션 통째로 렌더하게 한다.
    # (제조원가명세서 플래그와 동일 패턴.) 섹션은 'Ⅳ.판매비와관리비'~'영업손익/영업외'에서 끝.
    판관비_section = False
    for row in rows[header_idx + 1:]:
        대분류 = g(row, "대분류")
        계정명 = g(row, "계정명")
        _nm = str(계정명 or "")
        if "판매비와관리비" in _nm or "판매비및관리비" in _nm or "판매비와 관리비" in _nm:
            판관비_section = True
            continue
        if 판관비_section and any(k in _nm for k in ("영업손익", "영업이익", "영업외", "영업비용")):
            판관비_section = False
        # 상세행만: 대분류(공시용) 있고, 소계/합계/섹션머리 아님
        if 대분류 is None or str(대분류).strip() == "":
            continue
        if _is_subtotal(대분류) or _is_subtotal(계정명):
            continue
        기초 = g(row, "기초"); 기말 = g(row, "기말")
        수정차변 = g(row, "수정차변"); 수정대변 = g(row, "수정대변")
        수정후 = g(row, "수정후")
        dr = 수정차변 if _is_amount(수정차변) else 0
        cr = 수정대변 if _is_amount(수정대변) else 0
        # 수정사항 = 수정후(M) - 기말(I). 별도정산표는 자산은 M=I+K-L, 부채/자본은 M=I+L-K로
        # 부호를 반영해 M을 계산하므로, M-I가 계정유형 무관하게 부호까지 정확하다(K-L은 자산에만 맞음).
        adj = None
        if _is_amount(수정후) and _is_amount(기말):
            adj = round(수정후) - round(기말)
        else:
            adj = dr - cr
        result.append({
            "계정명": str(계정명).strip() if 계정명 is not None else None,
            "대분류": str(대분류).strip(),
            "판관비": 판관비_section,
            "기초": 기초 if _is_amount(기초) else None,
            "기말": 기말 if _is_amount(기말) else None,
            "수정차변": dr,
            "수정대변": cr,
            "수정사항": adj,
            "수정후": 수정후 if _is_amount(수정후) else None,
        })

    # ---- 제조원가명세서 섹션(별도정산표 하단) ----
    # 대분류(C열) 없이 계정명(B)만 있고 컬럼은 동일(H전기/I당기/M수정후). Q 매출원가·R-2 인건비 노무비
    # 소스. 대분류=계정명으로 부여해 동일 엔진/소유맵으로 라우팅(예: 당기총공사비용, 급여_제조).
    mfg_idx = None
    name_col = cm["계정명"]
    for i, row in enumerate(rows):
        v = row[name_col] if name_col < len(row) else None
        if v is not None and "제조원가명세서" in str(v):
            mfg_idx = i
            break
    if mfg_idx is not None:
        for row in rows[mfg_idx + 1:]:
            계정명 = g(row, "계정명")
            if 계정명 is None or str(계정명).strip() == "" or _is_subtotal(계정명):
                continue
            기말 = g(row, "기말"); 수정후 = g(row, "수정후"); 기초 = g(row, "기초")
            if not (_is_amount(기말) or _is_amount(수정후)):
                continue
            adj = (round(수정후) - round(기말)) if (_is_amount(수정후) and _is_amount(기말)) else 0
            nm = str(계정명).strip()
            result.append({
                "계정명": nm, "대분류": nm, "제조원가": True,
                "기초": 기초 if _is_amount(기초) else None,
                "기말": 기말 if _is_amount(기말) else None,
                "수정차변": 0, "수정대변": 0, "수정사항": adj,
                "수정후": 수정후 if _is_amount(수정후) else None,
            })
    return result
