"""`dsd-bridge describe` — AI 편집기를 위한 **구조 지도** (1.3 flow 도구).

통합문서(또는 `.dsd`)를 열어 "어느 시트의 어느 표에, 어떤 열 역할로, 어떤 행이 있는가" 와
"기간 표기가 적힌 셀이 어디인가" 를 마크다운(기본) 또는 JSON(`--json`)으로 낸다. 롤포워드를
사람 대신 하는 AI 가 **행 번호와 열 문자**를 알고 편집하도록 하는 것이 목적이므로, 값은
읽기 쉬운 형태로 줄이고 식별자(`_ID`·`_TABLE`)는 그대로 낸다.

읽기 전용이다 — 이 모듈은 파일을 쓰지 않는다(`--json` 저장 제외).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from .docmap import (
    COL_ROLES,
    DocMap,
    MapBlock,
    MapTable,
    excel_col_letter,
    load,
    period_hits,
)
from .errors import InputError, OutputError
from .text import DEFAULT_PATTERN, format_number

#: 기간 표기 셀 목록에서 한 줄에 보여 줄 최대 글자 수.
TEXT_CLIP = 60


def _clip(s: str, n: int = TEXT_CLIP) -> str:
    t = " ".join(str(s).split())
    return t if len(t) <= n else t[: n - 1] + "…"


def _compact(cell) -> str:
    """셀 하나의 표시 문자열 — 숫자는 콤마 표기로, 그 밖은 첫 줄만."""
    if cell is None:
        return ""
    a = cell.amount
    if a is not None:
        return format_number(a, DEFAULT_PATTERN)
    return _clip(cell.text, 40)


@dataclass(frozen=True)
class PeriodCell:
    """기간 표기가 든 셀 하나."""

    sheet: str
    where: str            # "표" | "문단"
    cell: str             # `시트!F31` 또는 `시트!F31`(문단은 F 열)
    row_id: str
    table: str
    patterns: tuple
    text: str

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "where": self.where, "cell": self.cell,
                "row_id": self.row_id, "table": self.table,
                "patterns": [{"name": n, "text": t} for n, t in self.patterns],
                "text": self.text}


@dataclass
class DescribeReport:
    doc: DocMap
    sheet_filter: str = ""
    period_cells: list = field(default_factory=list)

    @property
    def sheets(self) -> list:
        if not self.sheet_filter:
            return list(self.doc.sheets)
        return [s for s in self.doc.sheets if s.name == self.sheet_filter]

    # ── 마크다운 ──────────────────────────────────────────────────────────
    def lines_ko(self) -> list:
        doc = self.doc
        sheets = self.sheets
        tables = [t for s in sheets for t in s.tables]
        out = [
            f"# 구조 지도 — {doc.path.name}",
            "",
            f"- 입력: `{doc.source}` / 시트 {len(sheets)}개 / 표 {len(tables)}개"
            + (f" / contents_sha256 `{doc.contents_sha256[:12]}…`" if doc.contents_sha256 else ""),
            "",
        ]
        for s in sheets:
            out.append(f"## 시트 {s.order + 1}. {s.name}")
            out.append("")
            if not s.tables:
                out.append("표가 없습니다.")
                out.append("")
            for t in s.tables:
                out.extend(self._table_lines(t))
        out.extend(self._period_lines())
        return out

    def _table_lines(self, t: MapTable) -> list:
        cols = sorted(t.logical_cols)
        letters = t.col_letters()
        span = (f" (행 {t.start_row}~{t.end_row})"
                if t.start_row and t.end_row else "")
        kind = "기간표" if t.is_period else "일반표"
        out = [
            f"### 표 {t.index + 1} — `{t.eid}` · {t.layout} · {kind}{span}",
            "",
        ]
        if t.is_period:
            roles = " | ".join(
                f"{letters.get(c, excel_col_letter(c))}={COL_ROLES[c]}" for c in cols
                if c in COL_ROLES
            )
            out.append(f"- 열 역할: {roles}")
        else:
            spec = " | ".join(f"{letters.get(c, excel_col_letter(c))}=논리 {c}" for c in cols)
            out.append(f"- 열: {spec or '(없음)'}")
        if t.headers:
            out.append(f"- 머리글: {_clip(t.header_text, 120)}")
        if t.aclass:
            out.append(f"- ACLASS: {t.aclass}")
        out.append("")
        if t.is_period:
            out.append("| 행번호 | _ID | 라벨 | 당기 | 전기 |")
            out.append("|---|---|---|---|---|")
            for r in t.rows:
                cur = r.amount(0)
                prior = r.amount(1)
                out.append(
                    f"| {r.excel_row or '-'} | {r.eid} | {_clip(r.label, 40)} | "
                    f"{format_number(cur, DEFAULT_PATTERN) if cur is not None else ''} | "
                    f"{format_number(prior, DEFAULT_PATTERN) if prior is not None else ''} |"
                )
        else:
            out.append("| 행번호 | _ID | 라벨 | 값 |")
            out.append("|---|---|---|---|")
            for r in t.rows:
                vals = " / ".join(_compact(c) for c in r.cells[1:])
                out.append(f"| {r.excel_row or '-'} | {r.eid} | {_clip(r.label, 40)} | "
                           f"{_clip(vals, 60)} |")
        out.append("")
        return out

    def _period_lines(self) -> list:
        cells = self.period_cells
        out = ["## 기간 표기 셀", ""]
        if not cells:
            out.append("기간 표기를 찾지 못했습니다.")
            out.append("")
            return out
        out.append(f"롤포워드 때 갱신을 검토할 셀 {len(cells)}개입니다.")
        out.append("")
        out.append("| 위치 | 구분 | 표 | 맞은 표기 | 내용 |")
        out.append("|---|---|---|---|---|")
        for c in cells:
            pats = ", ".join(f"{n}:{t}" for n, t in c.patterns)
            out.append(f"| {c.cell} | {c.where} | {c.table or '-'} | {_clip(pats, 50)} | "
                       f"{_clip(c.text)} |")
        out.append("")
        return out

    # ── JSON ─────────────────────────────────────────────────────────────
    def json_obj(self) -> dict:
        return {
            "file": str(self.doc.path),
            "source": self.doc.source,
            "contents_sha256": self.doc.contents_sha256,
            "sheets": [self._sheet_obj(s) for s in self.sheets],
            "period_cells": [c.as_dict() for c in self.period_cells],
        }

    def _sheet_obj(self, s) -> dict:
        return {
            "name": s.name,
            "order": s.order,
            "tables": [self._table_obj(t) for t in s.tables],
            "blocks": [{"id": b.eid, "kind": b.kind, "row": b.excel_row,
                        "text": _clip(b.text, 200)} for b in s.blocks],
        }

    def _table_obj(self, t: MapTable) -> dict:
        letters = t.col_letters()
        return {
            "id": t.eid,
            "index": t.index,
            "layout": t.layout,
            "period": t.is_period,
            "aclass": t.aclass,
            "start_row": t.start_row,
            "end_row": t.end_row,
            "columns": [
                {"col": c, "letter": letters.get(c, excel_col_letter(c)),
                 "role": COL_ROLES.get(c, "") if t.is_period else ""}
                for c in sorted(t.logical_cols)
            ],
            "headers": [
                {"id": h.eid, "row": h.excel_row,
                 "cells": [{"col": c.col, "letter": c.letter, "text": _clip(c.text, 60)}
                           for c in h.cells]}
                for h in t.headers
            ],
            "rows": [self._row_obj(t, r) for r in t.rows],
        }

    def _row_obj(self, t: MapTable, r) -> dict:
        cur, prior = (r.amount(0), r.amount(1)) if t.is_period else (None, None)
        return {
            "id": r.eid,
            "row": r.excel_row,
            "label": r.label,
            "cells": [{"col": c.col, "letter": c.letter, "text": _clip(c.text, 60)}
                      for c in r.cells],
            "current": None if cur is None else str(cur),
            "prior": None if prior is None else str(prior),
        }

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n",
                         encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 구조 지도를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


def _collect_period_cells(doc: DocMap, sheets) -> list:
    out = []
    for s in sheets:
        for item in s.items:
            if isinstance(item, MapBlock):
                hits = period_hits(item.text)
                if hits:
                    where = f"{s.name}!{excel_col_letter(0)}{item.excel_row}" \
                        if item.excel_row else f"{s.name}!{item.eid}"
                    out.append(PeriodCell(s.name, "문단", where, item.eid, "",
                                          tuple(hits), _clip(item.text, 200)))
                continue
            for r in list(item.headers) + list(item.rows):
                for c in r.cells:
                    hits = period_hits(c.text)
                    if not hits:
                        continue
                    where = (f"{s.name}!{c.letter}{r.excel_row}" if r.excel_row
                             else f"{s.name}!{r.eid}#{c.col}")
                    out.append(PeriodCell(s.name, "표", where, r.eid, item.eid,
                                          tuple(hits), _clip(c.text, 200)))
    return out


def describe(path, *, sheet: str = "") -> DescribeReport:
    """`.dsd` 또는 `.xlsx` 하나의 구조 지도를 만든다."""
    doc = load(path)
    if sheet and doc.sheet(sheet) is None:
        names = ", ".join(s.name for s in doc.sheets)
        raise InputError(f"오류: 시트를 찾을 수 없습니다: {sheet}\n  있는 시트: {names}")
    report = DescribeReport(doc, sheet)
    report.period_cells = _collect_period_cells(doc, report.sheets)
    return report


__all__ = ["DescribeReport", "PeriodCell", "describe"]
