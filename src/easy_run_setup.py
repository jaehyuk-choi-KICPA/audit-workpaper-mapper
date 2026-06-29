# -*- coding: utf-8 -*-
"""작업생성기 — 회사 작업폴더 만들기.

회사명을 입력하면  작업/{회사명}/  폴더를 만들고, 그 안에 들어가야 할
4폴더(입력자료·변환자료·참고자료·출력조서)를 쏙 넣어 준다. 여러 회사를
연달아 만들 수 있다. 만든 뒤 [입력자료]에 원본 엑셀을 넣고 조서생성기를 돌리면 된다.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass

import workspace as W


def _line(m=""):
    print(m, flush=True)


def main():
    _line("=" * 56)
    _line("  작업생성기 — 회사 작업폴더 만들기")
    _line("=" * 56)
    root = W.work_root()
    _line(f"\n  작업폴더 위치: {root}")
    existing = sorted(p.name for p in root.iterdir() if p.is_dir()) if root.exists() else []
    if existing:
        _line("  이미 있는 회사: " + ", ".join(existing))

    _line("\n  회사명을 입력하면 그 이름의 폴더가 작업/ 안에 만들어집니다.")
    _line("  (여러 회사를 연달아 입력 가능 · 그만하려면 빈 칸으로 Enter)")

    made = []
    while True:
        name = input("\n  회사명 (예: ○○_2512): ").strip()
        if not name:
            break
        # 폴더명으로 못 쓰는 문자 방어
        bad = set('\\/:*?"<>|')
        if any(c in bad for c in name):
            _line('  ▷ 폴더명에 \\ / : * ? " < > | 는 쓸 수 없습니다. 다시 입력해 주세요.')
            continue
        exists = (root / name).exists()
        ws = W.workspace_for(name)
        tag = "이미 있어 그대로 사용" if exists else "새로 만듦"
        _line(f"  ✓ {name}  ({tag})")
        _line(f"      └ {ws.input_dir.parent}")
        _line(f"        ├ 입력자료   ← 여기에 원본 엑셀 ('{W.INPUT_README_NAME}' 안내 동봉)")
        _line("        ├ 변환자료   (프로그램 자동 사용)")
        _line("        ├ 참고자료   (환율·유무형상각 템플릿 동봉 · 전기조서도 여기)")
        _line("        └ 출력조서   (결과가 여기에 생성)")
        if name not in made:
            made.append(name)

    _line("\n" + "=" * 56)
    if made:
        _line(f"  완료 — {len(made)}개 회사 폴더 준비: {', '.join(made)}")
        _line("  이제 각 회사의 [입력자료]에 파일을 넣고 '조서생성기'를 실행하세요.")
    else:
        _line("  만든 폴더가 없습니다.")
    _line("=" * 56)
    return 0


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)
