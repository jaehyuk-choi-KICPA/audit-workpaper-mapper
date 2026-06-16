"""A-300 (기타조회내역) 생성기.

3개 이질 테이블을 채운다 (소계 없음):
  1. 퇴직연금자산   ← BANK 1번 중 계정분류='퇴직연금'
  2. 대출 거래      ← BANK 2-2번 (실대출만; 금액은 파란 숫자열)
  3. 담보제공 내용  ← BANK 9번 (실담보만; 금융기관별 그룹)

FN·월보 시트는 감사인 수행영역이므로 건드리지 않는다.

양식 보존: 테이블마다 컬럼이 다르고 소계가 없으므로 A200의 소계 엔진 대신
base의 하위 도구(capture_row/stamp_row=제목·헤더 재배치, render_rows=데이터 행,
스타일 복사)를 조합한다. 대출 행수가 템플릿과 달라도 아래 담보 섹션이
밀리지 않도록 위치를 동적 계산한다.
"""

from openpyxl.utils import column_index_from_string

from .base import WorkpaperGenerator


def _amt(v) -> int:
    import re
    if v is None or isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        return round(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return round(float(s)) if s else 0
    except ValueError:
        return 0


# ---------------------------------------------------------------------------
# 변환: 파서 결과 → A-300 테이블 행 dict (컬럼 키에 맞춤)
# ---------------------------------------------------------------------------

def build_a300_data(bank_rows: list[dict], loan_rows: list[dict],
                    collateral_rows: list[dict]) -> dict:
    pension = [{
        "ref":          b.get("조서번호"),
        "구분":          b.get("금융기관명"),
        "금융상품종류":   b.get("금융상품종류"),
        "계좌번호":       b.get("계좌번호"),
        "조회금액":       _amt(b.get("금액")),
        # 퇴직연금은 만기 개념이 없음 → 수시입출금 sentinel은 '-'로 표기
        "만기일":         ("-" if b.get("만기") in (None, "", "수시입출금") else b.get("만기")),
        "최종이자지급일":  "-",
    } for b in bank_rows if b.get("계정분류") == "퇴직연금"]

    loans = [{
        "ref":        l.get("조서번호"),
        "구분":        l.get("금융기관명"),
        "대출종류":     l.get("대출종류"),
        "약정한도액":   _amt(l.get("약정한도액")),
        "대출금액":     _amt(l.get("대출금액")),
        "대출일":       l.get("대출일"),
        "최종만기일":   l.get("최종만기일"),
        "연이율":       l.get("연이율"),
        "최종이자지급일": l.get("최종이자지급일"),
        "상환방법":     l.get("상환방법"),
        "담보보증":     l.get("담보보증"),
    } for l in loan_rows]

    collateral = [{
        "금융기관명":     c.get("금융기관명"),   # 그룹핑용 (표에는 미기재)
        "구분":          c.get("구분"),
        "담보보증내용":   c.get("담보보증내용"),
        "소유자":         c.get("소유자"),
        "감정금액":       _amt(c.get("감정금액")),
        "설정금액":       _amt(c.get("설정금액")),
        "설정순위":       c.get("설정순위"),
        "선순위설정금액":  _amt(c.get("선순위설정금액")),
    } for c in collateral_rows]

    return {"pension": pension, "loans": loans, "collateral": collateral}


# ---------------------------------------------------------------------------
# 엑셀 출력
# ---------------------------------------------------------------------------

class A300Generator(WorkpaperGenerator):

    def fill(self, data: dict, params: dict = None) -> None:
        self.open_template()
        cfg = self.config
        gap = cfg.get("gap", 1)

        pen, loan, coll = cfg["pension"], cfg["loan"], cfg["collateral"]

        # 1) 도너(제목·헤더·데이터 스타일) 캡처 — 초기화 전
        cap = {
            "pen_title": self.capture_row(pen["title_row"]),
            "pen_hdr":   self.capture_row(pen["header_row"]),
            "pen_dstyle": self.capture_row_style(pen["data_donor_row"]),
            "loan_title": self.capture_row(loan["title_row"]),
            "loan_hdr":   self.capture_row(loan["header_row"]),
            "loan_dstyle": self.capture_row_style(loan["data_donor_row"]),
            "coll_title": self.capture_row(coll["title_row"]),
            "coll_inst":  self.capture_row(coll["inst_subheader_row"]),
            "coll_hdr":   self.capture_row(coll["header_row"]),
            "coll_dstyle": self.capture_row_style(coll["data_donor_row"]),
        }

        # 2) 영역 초기화
        self.clear_region(cfg["clear_from"], cfg["clear_until"])

        r = cfg["clear_from"]

        # 3) 퇴직연금
        self.stamp_row(r, cap["pen_title"]); r += 2
        self.stamp_row(r, cap["pen_hdr"]); r += 1
        r, _ = self.render_rows(r, pen["columns"], data["pension"], cap["pen_dstyle"])
        r += gap

        # 4) 대출
        self.stamp_row(r, cap["loan_title"]); r += 2
        self.stamp_row(r, cap["loan_hdr"]); r += 1
        r, _ = self.render_rows(r, loan["columns"], data["loans"], cap["loan_dstyle"])
        r += gap

        # 5) 담보 (금융기관별 그룹)
        self.stamp_row(r, cap["coll_title"]); r += 2
        inst_col_idx = column_index_from_string(coll["inst_col"])
        groups = self._group_by(data["collateral"], "금융기관명")
        if not groups:
            # 담보 없음: 안내만
            self.ws[f'{coll["inst_col"]}{r}'] = "해당사항 없음"
            r += 1
        for inst, rows in groups:
            self.stamp_row(r, cap["coll_inst"], overrides={inst_col_idx: inst}); r += 2
            self.stamp_row(r, cap["coll_hdr"]); r += 1
            r, _ = self.render_rows(r, coll["columns"], rows, cap["coll_dstyle"])
            r += gap

    @staticmethod
    def _group_by(rows, key):
        order, buckets = [], {}
        for row in rows:
            k = row.get(key)
            if k not in buckets:
                buckets[k] = []
                order.append(k)
            buckets[k].append(row)
        return [(k, buckets[k]) for k in order]
