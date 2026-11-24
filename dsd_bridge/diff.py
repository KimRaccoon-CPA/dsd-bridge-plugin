"""`dsd-bridge diff` — 통합문서 둘을 견준다 (1.3 flow 도구).

두 가지 쓰임을 같은 명령으로 다룬다.

1. **같은 문서를 편집한 두 판** — 두 통합문서의 `_dsdx.contents_sha256` 이 같으면 행
   요소 ID(`_ID`)가 같은 뜻이므로 그것으로 짝짓는다.
2. **다른 문서 둘**(전기 보고서 ↔ 당기 보고서) — 행 ID 가 서로 다르므로
   `(시트 순서, 표 순서, 정규화 라벨, 같은 라벨 안의 순번)` 으로 짝짓는다. 짝이 없는 행은
   삽입(새 문서에만) · 삭제(옛 문서에만) 로 보고한다.

`--rollforward` 는 그 위에 **이월 검증**을 얹는다 — 기간표(`docmap.MapTable.is_period`)의
짝지어진 행마다

- `이월 OK` 새 문서의 **전기**(논리 열 3·4) = 옛 문서의 **당기**(1·2)
- `이월 불일치` 위가 어긋남
- `당기 미기입` 새 문서의 당기가 비었거나 옛 문서의 당기와 같다 (아직 안 채운 것으로 보임)
- `기간 표기 미갱신` 새 문서의 `제 N(당) 기` 표기가 옛 문서보다 크지 않다

빈칸 · `-` · 전각 공백은 모두 같게(0/값 없음) 본다 — `dsd_bridge.check.parse_amount` 규칙이다.

이 명령은 **검사 결과를 알리는 참고용**이므로 파일을 읽을 수 없을 때(코드 2)를 빼면 언제나
종료 코드 0 이다. 자본변동표처럼 구조가 바뀐 표는 오류가 아니라 `구조 차이` 로 적는다.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from .docmap import COL_ROLES, DocMap, MapTable, load, period_numbers
from .errors import OutputError
from .text import DEFAULT_PATTERN, format_number

#: 이월 판정.
CARRY_OK = "이월 OK"
CARRY_BAD = "이월 불일치"
CARRY_EMPTY = "당기 미기입"
PERIOD_STALE = "기간 표기 미갱신"

#: 행 짝짓기 결과.
MATCH_ID = "_ID"
MATCH_LABEL = "라벨"


def _fmt(v) -> str:
    return "" if v is None else format_number(v, DEFAULT_PATTERN)


def _same(a, b) -> bool:
    """금액 둘이 같은가 — 값 없음과 0 은 같게 본다 (`-`·빈칸·전각 공백)."""
    if a is None and b is None:
        return True
    if a is None:
        return b == 0
    if b is None:
        return a == 0
    return a == b


def _role(table: MapTable, col: int) -> str:
    if table.is_period and col in COL_ROLES:
        return COL_ROLES[col]
    letters = table.col_letters()
    return f"열 {col}({letters.get(col, '?')})"


# ─────────────────────────────────────────────────────────────────────────────
# 결과
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class CellChange:
    sheet: str
    table: str
    old_row: int
    new_row: int
    label: str
    col: int
    role: str
    old: str
    new: str

    def line_ko(self) -> str:
        where = f"{self.new_row}행" if self.new_row else self.table
        return (f"[{self.sheet}] {where} {self.label} / {self.role}: "
                f"{self.old or '(빈칸)'} → {self.new or '(빈칸)'}")

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "table": self.table, "old_row": self.old_row,
                "new_row": self.new_row, "label": self.label, "col": self.col,
                "role": self.role, "old": self.old, "new": self.new}


@dataclass(frozen=True)
class RowGap:
    """짝이 없는 행."""

    sheet: str
    table: str
    kind: str             # "삽입" | "삭제"
    row: int
    eid: str
    label: str

    def line_ko(self) -> str:
        where = f"{self.row}행" if self.row else self.eid
        return f"[{self.sheet}] {self.kind} — {where} {self.label}"

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "table": self.table, "kind": self.kind,
                "row": self.row, "id": self.eid, "label": self.label}


@dataclass(frozen=True)
class CarryItem:
    """이월 검증 한 건."""

    sheet: str
    table: str
    row: int
    label: str
    verdict: str
    old_current: object = None
    new_prior: object = None
    new_current: object = None
    note: str = ""

    def line_ko(self) -> str:
        where = f"{self.row}행" if self.row else self.table
        body = (f"[{self.sheet}] {where} {self.label}: "
                f"전기 {_fmt(self.new_prior) or '(빈칸)'} / "
                f"전기 보고서 당기 {_fmt(self.old_current) or '(빈칸)'} → {self.verdict}")
        return f"{body} ({self.note})" if self.note else body

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "table": self.table, "row": self.row,
                "label": self.label, "verdict": self.verdict,
                "old_current": None if self.old_current is None else str(self.old_current),
                "new_prior": None if self.new_prior is None else str(self.new_prior),
                "new_current": None if self.new_current is None else str(self.new_current),
                "note": self.note}


@dataclass(frozen=True)
class PeriodItem:
    """기간 표기 갱신 검사 한 건."""

    sheet: str
    table: str
    row: int
    where: str
    old_text: str
    new_text: str
    old_no: int
    new_no: int

    def line_ko(self) -> str:
        return (f"[{self.sheet}] {self.where}: '{self.old_text}' → '{self.new_text}' "
                f"(기수 {self.old_no} → {self.new_no})")

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "table": self.table, "row": self.row,
                "where": self.where, "old": self.old_text, "new": self.new_text,
                "old_no": self.old_no, "new_no": self.new_no}


@dataclass
class DiffReport:
    old: DocMap
    new: DocMap
    match_by: str = MATCH_LABEL
    rollforward: bool = False
    changes: list = field(default_factory=list)
    gaps: list = field(default_factory=list)
    carries: list = field(default_factory=list)
    stale: list = field(default_factory=list)
    #: 짝을 짓지 못한 시트·표 `[(사유, 설명)]`.
    structure: list = field(default_factory=list)
    matched_rows: int = 0

    def carry_counts(self) -> dict:
        out = {CARRY_OK: 0, CARRY_BAD: 0, CARRY_EMPTY: 0}
        for c in self.carries:
            if c.verdict in out:
                out[c.verdict] += 1
        return out

    @property
    def exit_code(self) -> int:
        return 0

    def summary_ko(self) -> list:
        out = [
            f"비교: {self.old.path.name} → {self.new.path.name}",
            f"짝짓기: {self.match_by} 기준 / 짝지은 행 {self.matched_rows}개 / "
            f"바뀐 셀 {len(self.changes)}개 / 삽입·삭제 {len(self.gaps)}개",
        ]
        if self.rollforward:
            c = self.carry_counts()
            out.append(
                f"이월: {CARRY_OK} {c[CARRY_OK]} / {CARRY_BAD} {c[CARRY_BAD]} / "
                f"{CARRY_EMPTY} {c[CARRY_EMPTY]} / {PERIOD_STALE} {len(self.stale)}"
            )
        return out

    def lines_ko(self, verbose: bool = False) -> list:
        lines = []
        if self.structure:
            lines.append("── 구조 차이 ──")
            lines.extend(f"안내: {why}" for why in self.structure)
            lines.append("")
        if self.changes:
            lines.append(f"── 바뀐 셀 {len(self.changes)}개 ──")
            shown = self.changes if verbose else self.changes[:200]
            lines.extend(c.line_ko() for c in shown)
            if len(shown) < len(self.changes):
                lines.append(f"… 그 밖 {len(self.changes) - len(shown)}개 (-v 로 전부 봅니다)")
            lines.append("")
        if self.gaps:
            lines.append(f"── 삽입·삭제 {len(self.gaps)}개 ──")
            shown = self.gaps if verbose else self.gaps[:100]
            lines.extend(g.line_ko() for g in shown)
            if len(shown) < len(self.gaps):
                lines.append(f"… 그 밖 {len(self.gaps) - len(shown)}개 (-v 로 전부 봅니다)")
            lines.append("")
        if self.rollforward:
            bad = [c for c in self.carries if c.verdict == CARRY_BAD]
            empty = [c for c in self.carries if c.verdict == CARRY_EMPTY]
            if bad:
                lines.append(f"── {CARRY_BAD} {len(bad)}건 ──")
                lines.extend(c.line_ko() for c in bad)
                lines.append("")
            if empty:
                lines.append(f"── {CARRY_EMPTY} {len(empty)}건 ──")
                lines.extend(c.line_ko() for c in empty)
                lines.append("")
            if self.stale:
                lines.append(f"── {PERIOD_STALE} {len(self.stale)}건 ──")
                lines.extend(p.line_ko() for p in self.stale)
                lines.append("")
            if verbose:
                ok = [c for c in self.carries if c.verdict == CARRY_OK]
                if ok:
                    lines.append(f"── {CARRY_OK} {len(ok)}건 ──")
                    lines.extend(c.line_ko() for c in ok)
                    lines.append("")
        lines.extend(self.summary_ko())
        return lines

    def json_obj(self) -> dict:
        obj = {
            "old": str(self.old.path),
            "new": str(self.new.path),
            "match_by": self.match_by,
            "matched_rows": self.matched_rows,
            "changes": [c.as_dict() for c in self.changes],
            "gaps": [g.as_dict() for g in self.gaps],
            "structure": list(self.structure),
        }
        if self.rollforward:
            obj["rollforward"] = {
                "counts": self.carry_counts(),
                "items": [c.as_dict() for c in self.carries],
                "period_stale": [p.as_dict() for p in self.stale],
            }
        return obj

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n",
                         encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 비교 결과를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


# ─────────────────────────────────────────────────────────────────────────────
# 짝짓기
# ─────────────────────────────────────────────────────────────────────────────


def _pair_tables(old_sheet, new_sheet, report: DiffReport) -> list:
    """시트 안의 표를 순서로 짝짓는다. 개수가 다르면 앞에서부터 짝짓고 나머지는 구조 차이."""
    ot, nt = old_sheet.tables, new_sheet.tables
    n = min(len(ot), len(nt))
    if len(ot) != len(nt):
        report.structure.append(
            f"[{new_sheet.name}] 표 개수가 다릅니다 — 전: {len(ot)}개 / 후: {len(nt)}개. "
            f"앞에서부터 {n}개만 짝지었습니다."
        )
    return list(zip(ot[:n], nt[:n]))


def _pair_rows_by_id(old_table: MapTable, new_table: MapTable):
    """행 `_ID` 로 짝짓는다 (같은 문서를 편집한 두 판)."""
    old_by_id = {r.eid: r for r in old_table.rows if r.eid}
    new_by_id = {r.eid: r for r in new_table.rows if r.eid}
    pairs = [(old_by_id[k], new_by_id[k]) for k in old_by_id if k in new_by_id]
    only_old = [r for r in old_table.rows if r.eid not in new_by_id]
    only_new = [r for r in new_table.rows if r.eid not in old_by_id]
    return pairs, only_old, only_new


def _pair_rows_by_label(old_table: MapTable, new_table: MapTable):
    """정규화 라벨 + 같은 라벨 안의 순번으로 짝짓는다 (다른 문서 둘)."""
    def keyed(rows):
        seen: dict = {}
        out = []
        for r in rows:
            k = r.norm
            seen[k] = seen.get(k, 0) + 1
            out.append(((k, seen[k]), r))
        return out

    okeys = keyed(old_table.rows)
    nkeys = dict(keyed(new_table.rows))
    pairs, only_old = [], []
    used = set()
    for k, r in okeys:
        hit = nkeys.get(k)
        if hit is not None:
            pairs.append((r, hit))
            used.add(k)
        else:
            only_old.append(r)
    only_new = [r for k, r in keyed(new_table.rows) if k not in used]
    return pairs, only_old, only_new


def _compare_cells(old_table, new_table, old_row, new_row, report: DiffReport) -> None:
    cols = sorted({c.col for c in old_row.cells} | {c.col for c in new_row.cells})
    for col in cols:
        oc, nc = old_row.cell_of(col), new_row.cell_of(col)
        oa = oc.amount if oc is not None else None
        na = nc.amount if nc is not None else None
        if oa is not None or na is not None:
            if _same(oa, na):
                continue
            old_s, new_s = _fmt(oa), _fmt(na)
        else:
            old_s = " ".join(oc.text.split()) if oc is not None else ""
            new_s = " ".join(nc.text.split()) if nc is not None else ""
            if old_s == new_s:
                continue
        report.changes.append(CellChange(
            sheet=new_table.sheet, table=new_table.eid,
            old_row=old_row.excel_row, new_row=new_row.excel_row,
            label=new_row.label or old_row.label, col=col,
            role=_role(new_table, col), old=old_s, new=new_s,
        ))


def _check_period_cells(old_table, new_table, old_row, new_row, report: DiffReport) -> None:
    """짝지은 행에서 `제 N(당|전) 기` 표기가 갱신되었는지 본다."""
    for col in sorted({c.col for c in new_row.cells}):
        nc, oc = new_row.cell_of(col), old_row.cell_of(col)
        if nc is None or oc is None:
            continue
        new_nos = period_numbers(nc.text)
        old_nos = period_numbers(oc.text)
        if not new_nos or not old_nos:
            continue
        for (n_no, n_kind), (o_no, o_kind) in zip(new_nos, old_nos):
            if n_kind == o_kind and n_no <= o_no:
                report.stale.append(PeriodItem(
                    sheet=new_table.sheet, table=new_table.eid, row=new_row.excel_row,
                    where=(f"{new_table.sheet}!{nc.letter}{new_row.excel_row}"
                           if new_row.excel_row else f"{new_table.eid}#{col}"),
                    old_text=" ".join(oc.text.split())[:40],
                    new_text=" ".join(nc.text.split())[:40],
                    old_no=o_no, new_no=n_no,
                ))


def _carry_check(old_table, new_table, old_row, new_row, report: DiffReport) -> None:
    old_cur = old_row.amount(0)
    new_prior = new_row.amount(1)
    new_cur = new_row.amount(0)
    if old_cur is None and new_prior is None:
        return
    label = new_row.label or old_row.label

    def item(verdict: str, note: str = "") -> CarryItem:
        return CarryItem(
            sheet=new_table.sheet, table=new_table.eid, row=new_row.excel_row,
            label=label, verdict=verdict, old_current=old_cur, new_prior=new_prior,
            new_current=new_cur, note=note,
        )

    report.carries.append(item(CARRY_OK if _same(new_prior, old_cur) else CARRY_BAD))
    # 당기 미기입은 이월 판정과 별개다 — 전기를 옳게 옮겼어도 당기를 아직 안 채웠을 수 있다.
    if new_cur is None:
        report.carries.append(item(CARRY_EMPTY, "당기 칸이 비어 있습니다"))
    elif old_cur is not None and _same(new_cur, old_cur):
        report.carries.append(item(CARRY_EMPTY, "당기가 전기 보고서의 당기와 같습니다"))


def diff(old_path, new_path, *, rollforward: bool = False) -> DiffReport:
    """통합문서(또는 `.dsd`) 둘을 견준다."""
    old = load(old_path)
    new = load(new_path)
    same_doc = bool(old.contents_sha256) and old.contents_sha256 == new.contents_sha256
    report = DiffReport(old, new, MATCH_ID if same_doc else MATCH_LABEL, rollforward)

    n = min(len(old.sheets), len(new.sheets))
    if len(old.sheets) != len(new.sheets):
        report.structure.append(
            f"시트 개수가 다릅니다 — 전: {len(old.sheets)}개 / 후: {len(new.sheets)}개. "
            f"앞에서부터 {n}개만 짝지었습니다."
        )
    for os_, ns in zip(old.sheets[:n], new.sheets[:n]):
        if os_.name != ns.name:
            report.structure.append(f"시트 이름이 다릅니다 — 전: {os_.name} / 후: {ns.name}")
        for ot, nt in _pair_tables(os_, ns, report):
            pairs, only_old, only_new = (
                _pair_rows_by_id(ot, nt) if same_doc else _pair_rows_by_label(ot, nt)
            )
            report.matched_rows += len(pairs)
            for r in only_old:
                report.gaps.append(RowGap(nt.sheet, nt.eid, "삭제", r.excel_row, r.eid, r.label))
            for r in only_new:
                report.gaps.append(RowGap(nt.sheet, nt.eid, "삽입", r.excel_row, r.eid, r.label))
            # 본문이 여러 줄인 표에서 행 구성이 바뀐 것은 눈에 띄게 알린다 — 자본변동표는
            # 롤포워드 때 당기 블록이 전기 블록이 되면서 행이 재구성되므로 여기서 걸린다.
            if (only_old or only_new) and (
                ot.is_period or nt.is_period or max(len(ot.rows), len(nt.rows)) >= 4
            ):
                report.structure.append(
                    f"[{nt.sheet}] 표 {nt.eid} — 행 구성이 다릅니다 "
                    f"(삭제 {len(only_old)} / 삽입 {len(only_new)}, 짝지은 행 {len(pairs)}). "
                    "오류가 아니라 구조 차이입니다."
                )
            for old_row, new_row in pairs:
                _compare_cells(ot, nt, old_row, new_row, report)
                if rollforward:
                    _check_period_cells(ot, nt, old_row, new_row, report)
                    if nt.is_period and ot.is_period:
                        _carry_check(ot, nt, old_row, new_row, report)
            if rollforward:
                for oh, nh in zip(ot.headers, nt.headers):
                    _check_period_cells(ot, nt, oh, nh, report)
    return report


__all__ = [
    "CARRY_OK",
    "CARRY_BAD",
    "CARRY_EMPTY",
    "PERIOD_STALE",
    "MATCH_ID",
    "MATCH_LABEL",
    "CellChange",
    "RowGap",
    "CarryItem",
    "PeriodItem",
    "DiffReport",
    "diff",
]
