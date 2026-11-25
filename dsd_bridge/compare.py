"""`dsd-bridge compare` — 보고서 두 판 비교 보고서 (0.6.0).

회사에 보낸 판 ↔ 회신받은 판, 정정 전 ↔ 후, DART 편집기 저장 전 ↔ 후처럼 **같은 보고서의 두 판**을
한 장으로 견준다. `.dsd`·작업 `.xlsx` 어느 쪽이든 받는다.

- **문단(문구)**: 시트마다 문단을 차례로 맞대어(순서 정렬) 추가·삭제·수정을 찾고, 수정된 문단은 낱말 단위로
  지운 글(빨간 취소선)과 새 글(초록 밑줄)을 표시한다
- **표**: `dsd-bridge diff` 의 짝짓기를 그대로 써서 숫자·글 칸 변경과 행 추가·삭제를 모은다
- `--report 비교.xlsx` 로 요약·문단 변경·표 변경·행 추가삭제 시트가 든 엑셀을 쓴다

읽기 전용이며 종료 코드는 언제나 0 이다.
"""

from __future__ import annotations

import difflib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from .diff import diff
from .docmap import MapBlock
from .errors import OutputError

#: 구분.
EDITED = "수정"
ADDED = "추가"
REMOVED = "삭제"
#: 기본 출력에서 보여 줄 문단 수.
SHOW = 30
_TOKEN = re.compile(r"\s+|[^\s]+")


def _norm(text) -> str:
    return " ".join(str(text or "").split())


def word_ops(old: str, new: str) -> list:
    """낱말 단위 비교 `[(구분, 글)]` — 구분은 `equal`/`delete`/`insert`."""
    a, b = _TOKEN.findall(old), _TOKEN.findall(new)
    out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op == "equal":
            out.append(("equal", "".join(a[i1:i2])))
        else:
            if i2 > i1:
                out.append(("delete", "".join(a[i1:i2])))
            if j2 > j1:
                out.append(("insert", "".join(b[j1:j2])))
    return out


def inline_ko(ops) -> str:
    """콘솔용 — `[-지운 글-]{+새 글+}`."""
    parts = []
    for op, text in ops:
        if op == "equal":
            parts.append(text)
        elif op == "delete":
            parts.append(f"[-{text}-]")
        else:
            parts.append(f"{{+{text}+}}")
    return "".join(parts)


@dataclass(frozen=True)
class TextChange:
    sheet: str
    old_row: int
    new_row: int
    kind: str
    old: str
    new: str

    @property
    def ops(self) -> list:
        return word_ops(self.old, self.new)

    def line_ko(self) -> str:
        where = f"{self.sheet} " + (f"{self.new_row}행" if self.new_row else f"(전 {self.old_row}행)")
        if self.kind == ADDED:
            return f"[문단 {ADDED}] {where} | {self.new[:120]}"
        if self.kind == REMOVED:
            return f"[문단 {REMOVED}] {where} | {self.old[:120]}"
        return f"[문단 {EDITED}] {where} | {inline_ko(self.ops)[:200]}"

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "old_row": self.old_row, "new_row": self.new_row,
                "kind": self.kind, "old": self.old, "new": self.new}


@dataclass
class CompareReport:
    old: Path
    new: Path
    texts: list = field(default_factory=list)
    table: object = None              # DiffReport
    notes: list = field(default_factory=list)

    @property
    def exit_code(self) -> int:
        return 0

    def counts(self) -> dict:
        out = {EDITED: 0, ADDED: 0, REMOVED: 0}
        for t in self.texts:
            out[t.kind] += 1
        return out

    def summary_ko(self) -> str:
        c = self.counts()
        d = self.table
        return (f"두 판 비교: 문단 {EDITED} {c[EDITED]} / {ADDED} {c[ADDED]} / {REMOVED} {c[REMOVED]} | "
                f"표 칸 변경 {len(d.changes)} / 행 추가·삭제 {len(d.gaps)}")

    def lines_ko(self, verbose: bool = False) -> list:
        lines = [f"비교: {self.old.name} → {self.new.name}"]
        lines.extend(f"안내: {n}" for n in self.notes + list(self.table.structure))
        if self.texts:
            lines.append(f"── 문단 변경 {len(self.texts)}건 ──")
            shown = self.texts if verbose else self.texts[:SHOW]
            lines.extend(t.line_ko() for t in shown)
            if len(shown) < len(self.texts):
                lines.append(f"… 그 밖 {len(self.texts) - len(shown)}건 (-v 또는 --report 로 전부 봅니다)")
        if self.table.changes:
            lines.append(f"── 표 칸 변경 {len(self.table.changes)}건 ──")
            shown = self.table.changes if verbose else self.table.changes[:SHOW]
            lines.extend(c.line_ko() for c in shown)
            if len(shown) < len(self.table.changes):
                lines.append(f"… 그 밖 {len(self.table.changes) - len(shown)}건 (-v 또는 --report 로 전부 봅니다)")
        if self.table.gaps:
            lines.append(f"── 행 추가·삭제 {len(self.table.gaps)}건 ──")
            lines.extend(g.line_ko() for g in self.table.gaps)
        lines.append(self.summary_ko())
        return lines

    def json_obj(self) -> dict:
        return {"old": str(self.old), "new": str(self.new),
                "texts": [t.as_dict() for t in self.texts],
                "table": self.table.json_obj(), "notes": list(self.notes)}

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 비교 결과를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc

    def write_xlsx(self, path) -> None:
        from openpyxl import Workbook
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        from openpyxl.styles import Alignment, Font, PatternFill

        red = InlineFont(strike=True, color="FFC00000")
        green = InlineFont(u="single", color="FF00873C")
        wb = Workbook()
        ws = wb.active
        ws.title = "요약"
        c = self.counts()
        for row in (["전 판", self.old.name], ["후 판", self.new.name], [],
                    ["문단 수정", c[EDITED]], ["문단 추가", c[ADDED]], ["문단 삭제", c[REMOVED]],
                    ["표 칸 변경", len(self.table.changes)], ["행 추가·삭제", len(self.table.gaps)], [],
                    ["읽는 법", "문단 변경 시트의 '비교' 칸: 빨간 취소선 = 지운 글, 초록 밑줄 = 새 글"]):
            ws.append(row)
        for n in self.notes + list(self.table.structure):
            ws.append(["안내", n])
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 90

        wt = wb.create_sheet("문단 변경")
        wt.append(["시트", "전 행", "후 행", "구분", "비교", "전 문구", "후 문구"])
        for t in self.texts:
            if t.kind == EDITED:
                rich = CellRichText([TextBlock(red, s) if op == "delete" else TextBlock(green, s)
                                     if op == "insert" else s for op, s in t.ops if s])
            else:
                rich = t.new if t.kind == ADDED else t.old
            wt.append([t.sheet, t.old_row or "", t.new_row or "", t.kind, rich, t.old, t.new])
            fill = {ADDED: "C6EFCE", REMOVED: "FFC7CE", EDITED: "FFEB9C"}[t.kind]
            wt.cell(wt.max_row, 4).fill = PatternFill("solid", fgColor=fill)
            wt.cell(wt.max_row, 5).alignment = Alignment(wrap_text=True, vertical="top")
        for col, width in zip("ABCDEFG", (16, 7, 7, 7, 80, 40, 40)):
            wt.column_dimensions[col].width = width

        wn = wb.create_sheet("표 변경")
        wn.append(["시트", "표", "전 행", "후 행", "계정과목", "열", "전", "후"])
        for ch in self.table.changes:
            wn.append([ch.sheet, ch.table, ch.old_row or "", ch.new_row or "", ch.label, ch.role, ch.old, ch.new])
        wg = wb.create_sheet("행 추가·삭제")
        wg.append(["시트", "표", "구분", "행", "계정과목"])
        for g in self.table.gaps:
            d = g.as_dict()
            wg.append([d.get("sheet"), d.get("table"), d.get("kind"), d.get("row") or "", d.get("label")])
        for sheet in (wt, wn, wg):
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            sheet.freeze_panes = "A2"
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            wb.save(p)
        except OSError as exc:
            raise OutputError(f"오류: 비교 보고서를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


def _blocks(sheet) -> list:
    return [b for b in sheet.items if isinstance(b, MapBlock) and _norm(b.text)]


def _text_changes(old_sheet, new_sheet) -> list:
    a, b = _blocks(old_sheet), _blocks(new_sheet)
    an, bn = [_norm(x.text) for x in a], [_norm(x.text) for x in b]
    out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, an, bn, autojunk=False).get_opcodes():
        if op == "equal":
            continue
        olds, news = a[i1:i2], b[j1:j2]
        n = min(len(olds), len(news)) if op == "replace" else 0
        for o, w in zip(olds[:n], news[:n]):
            out.append(TextChange(new_sheet.name, o.excel_row, w.excel_row, EDITED, _norm(o.text), _norm(w.text)))
        for o in olds[n:]:
            out.append(TextChange(new_sheet.name, o.excel_row, 0, REMOVED, _norm(o.text), ""))
        for w in news[n:]:
            out.append(TextChange(new_sheet.name, 0, w.excel_row, ADDED, "", _norm(w.text)))
    return out


def compare(old_path, new_path) -> CompareReport:
    """두 판을 견준다 (`.dsd`·`.xlsx`)."""
    d = diff(old_path, new_path)
    report = CompareReport(Path(old_path), Path(new_path), table=d)
    old_by_name = {s.name: s for s in d.old.sheets}
    used = set()
    for i, ns in enumerate(d.new.sheets):
        os_ = old_by_name.get(ns.name)
        if os_ is None and i < len(d.old.sheets) and d.old.sheets[i].name not in {s.name for s in d.new.sheets}:
            os_ = d.old.sheets[i]
        if os_ is None:
            report.notes.append(f"전 판에 없는 시트입니다: {ns.name} (문단은 모두 추가로 셉니다)")
            report.texts.extend(TextChange(ns.name, 0, b.excel_row, ADDED, "", _norm(b.text)) for b in _blocks(ns))
            continue
        used.add(os_.name)
        report.texts.extend(_text_changes(os_, ns))
    for os_ in d.old.sheets:
        if os_.name not in used and os_.name not in {s.name for s in d.new.sheets}:
            report.notes.append(f"후 판에 없는 시트입니다: {os_.name} (문단은 모두 삭제로 셉니다)")
            report.texts.extend(TextChange(os_.name, b.excel_row, 0, REMOVED, _norm(b.text), "") for b in _blocks(os_))
    return report


__all__ = ["EDITED", "ADDED", "REMOVED", "TextChange", "CompareReport", "compare", "word_ops", "inline_ko"]
