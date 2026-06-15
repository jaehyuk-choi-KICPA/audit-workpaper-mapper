"""조서 생성기 공통 기반 클래스."""

import shutil
from pathlib import Path

import openpyxl
import yaml


class WorkpaperGenerator:
    """양식 템플릿을 복사해 데이터를 채우는 기반 클래스.

    하위 클래스는 fill() 메서드를 구현해 계정별 로직을 정의한다.
    셀 주소는 모두 config YAML에서 읽어 소스코드에 하드코딩하지 않는다.
    """

    def __init__(self, template_path: str, config_path: str):
        self.template_path = Path(template_path)
        self.config = self._load_config(config_path)
        self.wb = None
        self.ws = None

    def _load_config(self, config_path: str) -> dict:
        with open(config_path, encoding="utf-8") as f:
            return yaml.safe_load(f)

    def open_template(self):
        self.wb = openpyxl.load_workbook(self.template_path)
        sheet_name = self.config.get("sheet")
        self.ws = self.wb[sheet_name] if sheet_name else self.wb.active

    def write_cell(self, cell_addr: str, value):
        """고정 셀에 값을 쓴다."""
        self.ws[cell_addr] = value

    def insert_data_rows(self, section: dict, rows: list[dict]):
        """데이터 단락에 행을 삽입하고 채운다.

        section: config에서 읽은 단락 정의 (start_row, columns)
        rows: 삽입할 데이터 목록
        """
        start_row = section["start_row"]
        columns = section["columns"]
        n = len(rows)

        if n == 0:
            return

        # 템플릿의 시작 행 아래에 (n-1)개 행 삽입 (첫 행은 기존 행 활용)
        if n > 1:
            self.ws.insert_rows(start_row + 1, n - 1)

        for i, row_data in enumerate(rows):
            r = start_row + i
            for col_letter, key in columns.items():
                self.ws[f"{col_letter}{r}"] = row_data.get(key)

    def inject_narrative(self, narrative: dict):
        """서술(목적·방법·결론)을 지정 셀에 쓴다."""
        narr_cfg = self.config.get("narrative", {})
        for key, cell_addr in narr_cfg.items():
            if key in narrative:
                self.write_cell(cell_addr, narrative[key])

    def save(self, output_path: str):
        """결과를 새 파일로 저장한다. 템플릿 원본은 보존된다."""
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(out)

    def fill(self, *args, **kwargs):
        raise NotImplementedError("하위 클래스에서 fill()을 구현해야 합니다.")
