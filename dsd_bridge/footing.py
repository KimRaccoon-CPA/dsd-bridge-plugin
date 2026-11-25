"""`dsd-bridge check --footing 풋팅.xlsx` — 한 장 요약 + 틱마크 풋팅 사본 (0.6.0).

검사한 보고서를 보고서 모양 그대로 엑셀로 펼친 뒤, 검사한 행의 과목 칸에 판정 색과 메모(틱마크)를 단다.

| 색 | 뜻 |
|---|---|
| 초록 ✓ | 합계검증 OK |
| 노랑 ≈ | 단수차 |
| 주황 ± | 부호주의 |
| 빨강 ✗ | 불일치 |
| 파랑 ☑ | 판정 기록으로 확인 처리 |
| 회색 – | 검사 제외 표 (사유는 메모) |

메모에는 기간별 `입력 / 계산 / 차이 / 판정`, 주석 상호대사(`--xref`)를 했으면 `CR 주석 n` 을 적는다.
맨 앞 `요약` 시트는 시트마다 검사 건수·판정 수·검사 제외 표·**검사하지 않은 시트**(주석처럼 합계검증 대상
표가 없는 시트)를 보여 준다. 숨김 시트를 지운 **검토용 사본**이므로 되돌리기(import)에는 쓰지 않는다.
"""

from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

from .check import MISMATCH, OK, ROUNDING, SIGN_NOTE
from .errors import OutputError
from .text import DEFAULT_PATTERN, format_number

ACCEPTED = "확인 처리"
EXCLUDED = "검사 제외"
COLORS = {OK: "C6EFCE", ROUNDING: "FFEB9C", SIGN_NOTE: "FCE4D6", MISMATCH: "FFC7CE",
          ACCEPTED: "DDEBF7", EXCLUDED: "D9D9D9"}
MARKS = {OK: "✓", ROUNDING: "≈", SIGN_NOTE: "±", MISMATCH: "✗", ACCEPTED: "☑", EXCLUDED: "–"}
#: 한 행의 판정이 여럿이면 이 순서로 가장 나쁜 것을 색으로 쓴다.
SEVERITY = (MISMATCH, SIGN_NOTE, ROUNDING, ACCEPTED, OK)


def _fmt(v) -> str:
    return format_number(v, DEFAULT_PATTERN)


def write_footing(source, out, report) -> Path:
    """`source`(.dsd 또는 작업 .xlsx)와 그 `CheckReport` 로 풋팅 사본을 `out` 에 쓴다."""
    from openpyxl import load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Font, PatternFill

    from .export_xlsx import export
    from .sheetstate import is_internal_sheet, load_workbook_state

    src, dst = Path(source), Path(out)
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp) / "footing.xlsx"
        if src.suffix.lower() == ".dsd":
            export(src, work)
        else:
            shutil.copyfile(src, work)
        state = load_workbook_state(work)
        where: dict = {}             # 행 요소 ID → (시트, Excel 행)
        table_start: dict = {}       # 표 요소 ID → (시트, Excel 행)
        for sheet in state.sheets:
            for r in sheet.rows:
                if r.id:
                    where.setdefault(r.id, (sheet.name, r.row))
                if r.table and r.table not in table_start:
                    table_start[r.table] = (sheet.name, r.row)
        wb = load_workbook(work)

        notes: dict = {}             # (시트, 행) → [메모 줄]
        verdicts: dict = {}          # (시트, 행) → {판정}
        per_sheet: dict = {}
        for it in report.items:
            loc = where.get(it.row)
            verdict = ACCEPTED if report._accepted(it) else it.verdict
            stats = per_sheet.setdefault(it.sheet, {k: 0 for k in (OK, ROUNDING, SIGN_NOTE, MISMATCH, ACCEPTED)})
            stats[verdict] += 1
            if loc is None:
                continue
            verdicts.setdefault(loc, set()).add(verdict)
            line = (f"{MARKS[verdict]} {it.period}: 입력 {_fmt(it.expected)} / 계산 {_fmt(it.actual)} / "
                    f"차이 {_fmt(it.diff)} → {verdict}")
            if verdict == ACCEPTED and report.decisions is not None:
                line += f" ({report.decisions.reason_ko(it.key)})"
            notes.setdefault(loc, []).append(line)
        if report.xref is not None:
            for x in report.xref.items:
                loc = where.get(x.fs.row)
                if loc is None or x.match is None:
                    continue
                notes.setdefault(loc, []).append(f"CR {x.fs.period}: {x.match.where_ko}"
                                                 + (f" ({x.ref_verdict})" if x.ref_verdict else ""))
        for sheet_name, eid, why in report.excluded:
            loc = table_start.get(eid)
            if loc is not None:
                verdicts.setdefault(loc, set()).add(EXCLUDED)
                notes.setdefault(loc, []).append(f"{MARKS[EXCLUDED]} {EXCLUDED}: {why}")

        for (sheet_name, row), lines in notes.items():
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            cell = ws.cell(row=row, column=6)
            vs = verdicts.get((sheet_name, row), set())
            worst = next((v for v in SEVERITY + (EXCLUDED,) if v in vs), None)
            if worst:
                cell.fill = PatternFill("solid", fgColor=COLORS[worst])
            cell.comment = Comment("\n".join(lines), "DSD브릿지", width=420, height=40 + 18 * len(lines))

        for name in list(wb.sheetnames):
            if is_internal_sheet(name):
                del wb[name]

        ws = wb.create_sheet("요약", 0)
        ws.append(["DSD브릿지 풋팅 사본 — 검토용(되돌리기에 쓰지 마세요)", src.name])
        ws.append([])
        ws.append(["시트", "OK", "단수차", "부호주의", "불일치", "확인 처리", "검사 제외 표", "비고"])
        header_row = ws.max_row
        excluded_by_sheet: dict = {}
        for sheet_name, _eid, why in report.excluded:
            excluded_by_sheet.setdefault(sheet_name, []).append(why)
        for sheet in state.sheets:
            stats = per_sheet.get(sheet.name)
            ex = excluded_by_sheet.get(sheet.name, [])
            if stats is None and not ex:
                ws.append([sheet.name, "", "", "", "", "", "", "검사하지 않음(합계검증 대상 표 없음)"])
                continue
            stats = stats or {k: 0 for k in (OK, ROUNDING, SIGN_NOTE, MISMATCH, ACCEPTED)}
            ws.append([sheet.name, stats[OK], stats[ROUNDING], stats[SIGN_NOTE], stats[MISMATCH],
                       stats[ACCEPTED], len(ex), "; ".join(sorted(set(ex)))[:200]])
            if stats[MISMATCH]:
                ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor=COLORS[MISMATCH])
        for c in ws[header_row]:
            c.font = Font(bold=True)
        ws.append([])
        ws.append(["합계검증", report.summary_ko()])
        if report.xref is not None:
            ws.append(["주석 상호대사", report.xref.summary_ko()])
        for r in report.review:
            ws.append(["확인 필요", r])
        ws.append([])
        ws.append(["표시", "과목 칸 색과 메모(마우스를 올리면 보임)"])
        for k in (OK, ROUNDING, SIGN_NOTE, MISMATCH, ACCEPTED, EXCLUDED):
            ws.append([f"{MARKS[k]} {k}", ""])
            ws.cell(ws.max_row, 1).fill = PatternFill("solid", fgColor=COLORS[k])
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["H"].width = 60
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            wb.save(dst)
        except OSError as exc:
            raise OutputError(f"오류: 풋팅 사본을 쓸 수 없습니다: {dst}\n  사유: {exc}") from exc
    return dst


__all__ = ["write_footing", "COLORS", "MARKS", "ACCEPTED", "EXCLUDED"]
