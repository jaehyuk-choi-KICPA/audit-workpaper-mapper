"""A100 (완전성 체크) 시트 생성기.

CS Control Sheet 파싱 결과를 받아 A100의 두 섹션을 채운다:
  - TEST 3번 테이블: 구분·일련번호·금융기관·각종 체크 컬럼
  - TEST 4번 서술: 발송 기관 수 자동 생성

config YAML 기대 구조 (_internal/config/ 에 위치):
  sheet: A100
  header_cells:
    회사명:   B2      # 예시 — 실제 셀 주소는 양식에 맞게 지정
    preparer: F3
    reviewer: G3
    기준일:   H3
  test3_table:
    start_row: 12
    columns:
      A: 구분
      B: 일련번호
      C: 금융기관
      D: 당기회사제시CS
      E: 은행연합회자료
      F: 금융자산부채원장
      G: 관련손익원장
      H: 발송여부
  test4_narrative_cell: B30
"""

from .base import WorkpaperGenerator

_V = "V"
_BLANK = ""

# A-100 TEST3 체크 컬럼 채우기 규칙 (완성본 V로 확정, 2026-06-15)
#  - 당기회사제시CS / 금융자산부채원장 / 관련손익원장 : 무조건 V
#  - 은행연합회자료 : 공란. 외부 은행연합회 데이터 대조 결과라 입력 3개 파일로
#    판단 불가 → 시스템은 비워두고 감사인이 채운다 (환각 방지).
#    실제 완성본에서도 하나·농협은행 등이 공란이라 기관 유형으로 추정 불가.


class A100Generator(WorkpaperGenerator):
    """A100 완전성 체크 테이블 생성기."""

    def fill(self, cs_rows: list[dict], params: dict) -> None:
        """A100 시트를 채운다.

        Args:
            cs_rows: parse_cs() 반환값
                     [{"구분", "일련번호", "금융기관명"}, ...]
            params:  공통 헤더 파라미터
                     {"회사명", "preparer", "reviewer", "기준일"}
        """
        self.open_template()
        cfg = self.config

        # 공통 헤더 셀
        for key, cell_addr in cfg.get("header_cells", {}).items():
            self.write_cell(cell_addr, params.get(key))

        # TEST 3번 테이블 행 구성
        table_rows = []
        for r in cs_rows:
            table_rows.append({
                "구분":            r.get("구분"),
                "일련번호":        r.get("일련번호"),
                "금융기관":        r.get("금융기관명") or "",
                "당기회사제시CS":  _V,
                "은행연합회자료":  _BLANK,   # 감사인 영역 — 비워둠
                "금융자산부채원장": _V,
                "관련손익원장":    _V,
                "발송여부":        _V,
            })

        self.insert_data_rows(cfg["test3_table"], table_rows)

        # TEST 4번 서술
        narrative_cell = cfg.get("test4_narrative_cell")
        if narrative_cell:
            n = len(cs_rows)
            self.write_cell(
                narrative_cell,
                f"회사는 당기 금융기관 CS에 금융기관을 {n}개 제시하였음. "
                f"총 {n}개 금융기관에 조회서를 발송함.",
            )
