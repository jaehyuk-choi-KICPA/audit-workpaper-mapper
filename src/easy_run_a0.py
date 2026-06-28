# -*- coding: utf-8 -*-
"""A-0(현금및현금성자산, 장단기금융상품) 개별시트 채움 — 쉬운 실행기 (비전문가용 대화형).

★ 총괄표는 만들지 않는다 — **조서생성.exe가 이미 만든 A-0(총괄표 채워진 것)** 를 읽어,
  그 총괄표에 **연동**해 A-0 전용 개별시트만 동적 구성한다(중복 생성 방지):
  · A050 주석    : 총괄표 계정 수정후/기초를 =ROUND(!/1000,0) 참조(이름매칭, 없는 계정 skip)
  · A020 실사    : 기말환율표(참고자료) + 외화 원금×환율 가이드
  · 보험가입현황 : 조회서 INSURANCE 해약환급금 per-policy 매핑

사용 순서:
  1) 먼저 **조서생성.exe**로 총괄표를 만든다 → `출력조서/{회사명}/A-0_4000_현금및현금성자산.xlsx`
  2) `입력자료/`에 금융기관조회서를 둔다(보험가입현황용 — 없으면 빈 표).
  3) `A-0생성.bat` 더블클릭 (또는 python src/easy_run_a0.py)
  4) 회사명 입력 → 위 A-0 파일에 개별시트가 채워진다(완성본 보존 graft).
"""
import glob
import io
import os
import re
import sys
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore", message=".*[Hh]eader or footer.*")
sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass

import openpyxl
from extractors import load_fx_rates
from generators.sheet_surgery import graft_sheet
from generators import a0_detail

TOTAL_SHEET = "4000_A000 총괄표"
A050_SHEET = "A050_주석검토"
A020_SHEET = "A020_금융자산 실사"
HOLD_SHEET = "장기금융상품_보험가입현황"


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE = _base_dir()
INPUT_DIR = Path(os.environ.get("A1_INPUT_DIR", BASE / "입력자료"))
PARSED_DIR = Path(os.environ.get("A1_PARSED_DIR", BASE / "변환자료"))
TEMPLATE_DIR = Path(os.environ.get("A1_TEMPLATE_DIR", BASE / "양식자료"))
OUTPUT_DIR = Path(os.environ.get("A1_OUTPUT_DIR", BASE / "출력조서"))
REF_DIR = Path(os.environ.get("A1_REF_DIR", BASE / "참고자료"))
CONFIG_DIR = Path(getattr(sys, "_MEIPASS", str(BASE))) / "_internal" / "config"
if not getattr(sys, "frozen", False) and not list(TEMPLATE_DIR.glob("*.xls*")):
    _dev = BASE / "_internal" / "양식"
    if _dev.exists():
        TEMPLATE_DIR = _dev


def _line(m=""):
    print(m, flush=True)


def _ask(label, required=True, example=""):
    hint = f" (예: {example})" if example else ""
    while True:
        v = input(f"  {label}{hint}: ").strip()
        if v or not required:
            return v
        _line("  ▷ 값을 입력해 주세요.")


def _find(pattern):
    c = [h for h in glob.glob(str(INPUT_DIR / "**" / pattern), recursive=True) if "~$" not in h]
    return sorted(c)[0] if c else None


def _extract_one(src, sheet, out):
    """완성본에서 단일 시트만 가벼운 임시본으로 추출(도형·외부링크 제거)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in zin.namelist():
            if n.startswith(("xl/drawings/", "xl/media/")):
                continue
            b = zin.read(n)
            if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                b = re.sub(rb"<drawing\b[^>]*/>", b"", b); b = re.sub(rb"<legacyDrawing\b[^>]*/>", b"", b)
            elif n.endswith(".rels"):
                b = re.sub(rb'<Relationship[^>]*(?:drawing|image|vmlDrawing)"[^>]*/>', b"", b)
            elif n == "[Content_Types].xml":
                b = re.sub(rb"<Override[^>]*(?:drawing|vmlDrawing)[^>]*/>", b"", b)
            zo.writestr(n, b)
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    for s in list(wb.sheetnames):
        if s != sheet:
            del wb[s]
    wb[sheet].sheet_state = "visible"
    wb.save(out)


def _edit_sheet(path, sheet, fn, tmp):
    """extract_one → fn(ws) → graft_sheet(완성본 보존)로 한 시트 적용."""
    light = str(Path(tmp) / "light.xlsx")
    _extract_one(path, sheet, light)
    wb = openpyxl.load_workbook(light)
    fn(wb[sheet])
    wb.save(light)
    out = path + ".graft"
    graft_sheet(path, light, sheet, out)
    os.replace(out, path)


def _inject_cols(path, sheet, widths):
    """graft가 빈 템플릿 cols를 제거하므로, 시트에 <cols> 너비를 XML로 재주입."""
    import xml.etree.ElementTree as ET
    M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    z = zipfile.ZipFile(path)
    wbx = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    r2t = {r.get("Id"): r.get("Target") for r in rels}
    target = None
    for s in wbx.iter(M + "sheet"):
        if s.get("name") == sheet:
            target = "xl/" + r2t[s.get(RID)].lstrip("/")
    z.close()
    if not target:
        return
    cols = "<cols>" + "".join(
        '<col min="%d" max="%d" width="%d" customWidth="1"/>' % (c, c, w) for c, w in sorted(widths.items())
    ) + "</cols>"
    tmp = path + ".w"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zin.infolist():
            b = zin.read(it.filename)
            if it.filename == target:
                b = re.sub(rb"<cols>.*?</cols>", b"", b, flags=re.DOTALL)
                b = b.replace(b"<sheetData", cols.encode() + b"<sheetData", 1)
            zo.writestr(it, b)
    os.replace(tmp, path)


def _find_a0_outputs():
    """출력조서/**/에서 조서생성이 만든 A-0 파일을 찾는다(총괄표 시트 보유 확인)."""
    cands = [h for h in glob.glob(str(OUTPUT_DIR / "**" / "*.xls*"), recursive=True)
             if "~$" not in h and re.search(r"A-?0[_\b]", Path(h).name) and "현금" in Path(h).name]
    good = []
    for h in cands:
        try:
            z = zipfile.ZipFile(h)
            wbx = z.read("xl/workbook.xml").decode("utf-8", "ignore"); z.close()
            if "총괄표" in wbx:
                good.append(h)
        except Exception:
            pass
    return sorted(set(good))


def main():
    _line("=" * 56)
    _line("  A-0 개별시트 채움 (총괄표 연동 — 보험가입현황·A020·A050)")
    _line("=" * 56)
    for d in (INPUT_DIR, PARSED_DIR, OUTPUT_DIR, REF_DIR):
        d.mkdir(parents=True, exist_ok=True)

    a0s = _find_a0_outputs()
    if not a0s:
        _line("\n[안내] 조서생성으로 만든 A-0 파일을 찾지 못했습니다.")
        _line("       먼저 '조서생성'으로 총괄표를 만든 뒤 이 프로그램을 실행하세요.")
        _line(f"       (확인 위치: {OUTPUT_DIR}\\{{회사명}}\\A-0_4000_현금및현금성자산.xlsx)")
        return 2
    if len(a0s) == 1:
        a0_path = a0s[0]
    else:
        _line("\n[A-0 파일이 여러 개] 회사명을 입력해 선택하세요:")
        for h in a0s:
            _line("   - " + str(Path(h).parent.name))
        company = _ask("회사명")
        sel = [h for h in a0s if Path(h).parent.name == company]
        if not sel:
            _line("  해당 회사 A-0를 못 찾았습니다."); return 2
        a0_path = sel[0]
    out_dir = Path(a0_path).parent
    _line(f"\n[A-0 대상] {a0_path}")

    confirm = _find("*조회서*기말감사*.xls*") or _find("필수자료2*") or _find("*금융기관조회서*.xls*")
    ledger = _find("필수자료3*") or _find("*잔액*.xls*") or _find("*거래처원장*.xls*")
    _line(f"[조회서]   {Path(confirm).name if confirm else '없음 — 장기금융/보험가입현황 생략'}")
    _line(f"[잔액원장] {Path(ledger).name if ledger else '없음 — 장기금융 잔액 매칭 생략'}")

    tmp = str(out_dir / "_tmp_a0")
    Path(tmp).mkdir(parents=True, exist_ok=True)
    fx = load_fx_rates(str(REF_DIR)) or {}
    policies = a0_detail.parse_insurance_holdings(confirm) if confirm else []
    recon = a0_detail.build_longterm_recon(confirm, ledger) if (confirm and ledger) else []
    recon_st = a0_detail.build_shortterm_recon(confirm, ledger) if (confirm and ledger) else []
    notes = []

    # 총괄표 연동값 읽기용(A050이 참조) — 같은 파일의 총괄표 ws
    try:
        twb = openpyxl.load_workbook(a0_path, data_only=False)
        total_ws = twb[TOTAL_SHEET]

        st_placed = {"n": 0}

        def _do_total(ws):
            a0_detail.cleanup_total_marker(ws)
            a0_detail.fill_total_longterm(ws, recon)        # 5) 장기금융 평가대사(A-1/조회서 기반)
            st_placed["n"] = a0_detail.fill_total_shortterm(ws, recon_st)  # 4) 단기금융(존재+표 있을 때)
        _edit_sheet(a0_path, TOTAL_SHEET, _do_total, tmp)
        notes.append("총괄표 5)장기금융: 조회서 보험사 %d건 평가대사" % len(recon))
        if not recon_st:
            notes.append("총괄표 4)단기금융: 없음(조회서 INVESTMENT 평가액 0 → prune)")
        elif st_placed["n"]:
            notes.append("총괄표 4)단기금융: %d건 평가대사" % st_placed["n"])
        else:
            notes.append("총괄표 4)단기금융: %d건 감지(%s 등) — 단 양식에 4)평가대사 표가 없어 미배치(수동/표추가 필요)"
                         % (len(recon_st), str(recon_st[0]["거래처"])[:12]))

        _edit_sheet(a0_path, A050_SHEET, lambda ws: a0_detail.fill_a050_notes(ws, total_ws, TOTAL_SHEET), tmp)
        notes.append("A050 주석: 총괄표 ROUND 연동")
        _edit_sheet(a0_path, A020_SHEET, lambda ws: a0_detail.fill_a020_fx(ws, fx), tmp)
        notes.append("A020 실사: 기말환율표(%d통화)" % len(fx))
        _edit_sheet(a0_path, HOLD_SHEET, lambda ws: a0_detail.fill_insurance_holdings(ws, policies), tmp)
        _inject_cols(a0_path, HOLD_SHEET, a0_detail.HOLDINGS_COL_WIDTHS)
        notes.append("보험가입현황: 조회서 %d건 매핑" % len(policies))
    except Exception as e:
        notes.append(f"개별시트 일부 실패: {type(e).__name__}: {e}")

    # 정리
    try:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception:
        pass

    lines = ["=" * 60, "  A-0 생성 완료", "=" * 60, f"  결과: {a0_path}"]
    for n in notes:
        lines.append("   - " + n)
    lines.append("  ※ 초안입니다. 보험가입현황은 조회서 회신분 — 회사 보험명세로 보완하세요.")
    summary = "\n".join(lines)
    _line("\n" + summary)
    try:
        (out_dir / f"{out_dir.name}_A0_실행리포트.txt").write_text(summary, encoding="utf-8")
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)
