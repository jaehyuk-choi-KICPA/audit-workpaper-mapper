# -*- coding: utf-8 -*-
"""2단계(별도 EXE) — 기타자산(D200)·기타부채(CC200) 상세 시트 일괄 생성 파이프라인.

1단계 산출물(출력조서/{회사}/D_4000_…·CC_4000_…)과 거래처원장(이상적 잔액명세서)을 입력받아
각 조서의 두번째 시트(상세)를 총괄표 계정 기반으로 동적 생성하고 무손실 이식(graft)한다.

흐름(조서별, 전부 _safe 래핑 — 부분 실패 허용, 항상 출력):
  ① trim_blank_tail   완성본 상세 시트의 빈 꼬리행 제거(회계법인 bloat 방어, 클린엔 무영향)
  ② scan_lead_accounts 총괄표(D100/CC100) read-only → 계정·행·기초·기말
  ③ extract_light      상세 시트만 가벼운 임시본으로 추출
  ④ render_detail      계정 기반 동적 렌더(거래처원장 매핑·합계·tie-out·check)
  ⑤ graft_sheet        완성본에 상세 시트만 무손실 이식 → os.replace 제자리 갱신
  ⑥ validate_detail_blocks  양식 무결성 게이트

완성본은 절대 openpyxl로 열어 저장하지 않는다(총괄표는 read-only 읽기, 상세는 extract/ graft).
"""

import os
import sys
import time
import warnings
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import openpyxl
import yaml

from cache import cached_ideal_ledger
from generators import sheet_surgery, lead_detail
from pipeline import RunReport
import validator


# 조서 레지스트리: code / config / 산출물 glob(1단계 출력) / 입력자료 토큰.
DETAIL_REGISTRY = [
    {"code": "D",  "config": "d200.yaml",  "glob": ("D_4000_*기타자산*.xls*", "D_*기타자산*.xls*", "D_4000_*.xls*")},
    {"code": "CC", "config": "cc200.yaml", "glob": ("CC_4000_*기타부채*.xls*", "CC_*기타부채*.xls*", "CC_4000_*.xls*")},
]


def _replace_robust(src: str, dst: str, tries: int = 6) -> None:
    """os.replace의 Windows 일시 파일락(WinError 5/32)을 재시도+copy 폴백으로 견고화."""
    import shutil
    last = None
    for i in range(tries):
        try:
            os.replace(src, dst)
            return
        except PermissionError as e:
            last = e
            time.sleep(0.4 * (i + 1))
    # 최후: 복사 후 원본 삭제 시도
    shutil.copyfile(src, dst)
    try:
        os.remove(src)
    except OSError:
        pass


def _find_one(globs, root: Path):
    import glob as _g
    for pat in globs:
        hits = [h for h in _g.glob(str(root / "**" / pat), recursive=True) if "~$" not in h]
        if hits:
            return sorted(hits)[0]
    return None


def build_detail_all(*, output_root, ledger_path, config_dir, parsed_dir=None,
                     params=None, progress=None, registry=None):
    """출력조서 폴더의 D·CC 완성본 상세 시트를 일괄 생성. Returns {code: RunReport}."""
    output_root = Path(output_root)
    cfgdir = Path(config_dir)
    parsed = Path(parsed_dir or (output_root.parent / "변환자료" / output_root.name))
    parsed.mkdir(parents=True, exist_ok=True)
    tmp = parsed / "_detail_tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    reg = registry or DETAIL_REGISTRY
    reports = {}

    # 거래처원장 → 이상적 잔액명세서(해시 캐시, 1회) — 없으면 빈 매핑(부분출력).
    by_ledger = {}
    ledger_report = RunReport()
    if ledger_path and Path(ledger_path).exists():
        def _ld():
            recs = cached_ideal_ledger(str(ledger_path), str(parsed), str(cfgdir))
            return lead_detail.group_ledger(recs)
        try:
            by_ledger = _ld()
            ledger_report.add("입력", "ok", f"거래처원장 {len(by_ledger)}계정 매핑 로드")
        except Exception as e:
            ledger_report.add("입력", "warn", f"거래처원장 파싱 실패({type(e).__name__}: {e}) — 잔액명세 생략")
    else:
        ledger_report.add("입력", "warn", "거래처원장 없음 — 거래처별 명세 생략(tie/check만)")

    for item in reg:
        code = item["code"]
        report = RunReport()
        for st, lv, m in ledger_report.entries:
            report.add(st, lv, m)
        reports[code] = report
        full = _find_one(item["glob"], output_root)
        if not full:
            report.add("입력", "error", f"{code}: 1단계 산출물 미발견({item['glob'][0]})")
            if progress:
                progress(code, False)
            continue
        cfg = yaml.safe_load((cfgdir / item["config"]).read_text(encoding="utf-8"))

        def _run():
            sheet = lead_detail.resolve_sheet(full, cfg["sheet_prefix"])
            lead = lead_detail.resolve_sheet(full, cfg["lead"]["sheet_prefix"])
            if not sheet or not lead:
                raise ValueError(f"시트 미발견(상세={sheet}, 총괄표={lead})")
            cfg["sheet"] = sheet
            cfg["lead"]["sheet"] = lead
            # ① debloat
            deb = str(tmp / f"{code}_deb.xlsx")
            removed = sheet_surgery.trim_blank_tail(full, sheet, deb)
            if removed:
                report.add("생성", "ok", f"{code}: 빈 꼬리행 {removed}행 정리")
            # ② scan
            accounts = lead_detail.scan_lead_accounts(deb, cfg["lead"])
            report.add("입력", "ok", f"{code}: 총괄표 계정 {len(accounts)}개")
            # ③ extract → ④ render
            light = str(tmp / f"{code}_light.xlsx")
            sheet_surgery.extract_light(deb, sheet, light)
            wb = openpyxl.load_workbook(light)
            ws = wb[sheet]
            info = lead_detail.render_detail(ws, cfg, accounts, by_ledger)
            wb.save(light)
            report.add("생성", "ok", f"{code}: 상세 블록 {info['blocks']}개 생성")
            if info.get("unmatched"):
                uniq = list(dict.fromkeys(info["unmatched"]))   # 순서 유지 dedup
                report.add("확인필요", "warn",
                           f"{code}: 원장 미발견 계정 {len(uniq)}개 — {', '.join(uniq[:8])}")
            # ⑤ graft → 제자리 갱신(Windows 파일락 대비 재시도+copy 폴백)
            out_tmp = str(tmp / f"{code}_out.xlsx")
            sheet_surgery.graft_sheet(deb, light, sheet, out_tmp)
            _replace_robust(out_tmp, full)
            report.add("출력", "ok", f"{code}: 완성본 이식 → {Path(full).name}")
            # ⑥ validate
            try:
                issues = validator.validate_detail_blocks(
                    full, sheet, expected_blocks=len(accounts), lead_sheet=lead)
                if issues:
                    for m in issues[:8]:
                        report.add("확인필요", "warn", m)
                else:
                    report.add("출력", "ok", f"{code}: 양식 무결성 0 issue")
            except Exception as e:
                report.add("출력", "warn", f"{code}: 검증 실패({type(e).__name__}: {e})")
            return True

        try:
            _run()
        except Exception as e:
            report.add("생성", "error", f"{code} 생성 실패({type(e).__name__}: {e})")
        if progress:
            try:
                progress(code, not report.has_error)
            except Exception:
                pass

    return reports
