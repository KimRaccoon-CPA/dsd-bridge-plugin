"""주석 ↔ 재무제표 상호대사 `dsd-bridge check --xref` (계획 1.2 WP7).

재무제표(재무상태표·손익계산서·현금흐름표·자본변동표)의 금액 하나하나를 **주석의 표**에서
찾아 준다. 단순히 첫 일치만 찾는 방식과 달리:

- 후보를 **모두** 모으고 행 라벨·주석 제목이 가장 비슷한 것을 고른다
- 절댓값으로 찾되 **`같은 부호` / `부호 반대`** 를 함께 적는다
- 정확히 같으면 `대사`, 허용 오차 안이면 **`단수차`** 로 나눠 적는다
- 주석 표의 `(단위: 천원)` 을 읽어 **원 단위로 환산해** 견준다

이 검사는 **참고용**이다. 합계검증과 달리 판정이 아니므로 **종료 코드에 영향을 주지 않으며**,
`--xref` 를 주지 않으면 아무것도 하지 않는다(기존 `check` 결과는 그대로다).

어느 표를 보는가
----------------
- **재무제표 쪽**: `dsd-bridge check` 가 이미 검사하는 표 그대로다 (`ACODE`/`ALEVEL` 이 있는 표준
  재무제표 표 + 자본변동표). 표준 표는 `당기`·`전기` 를 따로, 자본변동표는 **열 이름**
  (자본금·이익잉여금·총계 …)을 기간 자리에 적는다.
- **주석 쪽**: 재무제표 표가 아니면서 **주석 번호 제목**(WP5 의 `note_heading` 규칙 — `3. 사용제한예금`)
  아래에 있는 모든 표. `--split-notes` 없이도 같은 규칙으로 번호를 알 수 있다.
  `외부감사 실시내용` 섹션은 제외한다 — 그 섹션도 `1. 감사대상업무` 처럼 번호 제목을 쓰지만
  금액표가 아니라 인원·시간표다.

단위
----
주석 표는 대개 `(단위: 천원)` 을 표 바로 위의 한 줄짜리 표나 문단에 적는다. 섹션을 훑으면서
`단위 … 원|천원|백만원` 을 만나면 그때부터의 표에 그 단위를 물리고, **주석 번호가 바뀌면 원**
으로 되돌린다. 재무제표 표에도 같은 규칙을 쓴다(재무제표 표는 보통 `(단위 : 원)` · `AUNIT="WON"`).
견줄 때는 양쪽을 **원 단위로 환산**한다. 구현하는 단위는 `원`·`천원`·`백만원` 셋뿐이다.

고르는 법
---------
1. 절댓값이 **정확히 같은** 주석 셀을 모은다 → 있으면 `대사`.
2. 없으면 `|차이| ≤ tolerance` 인 셀을 모은다 → 있으면 `단수차`.
3. 그것도 없으면 `미대사`.

후보가 여럿이면 **재무제표 행 라벨**과 **주석 행 라벨 + 주석 제목**이 공유하는 글자 수로
순위를 매기고, 같으면 문서에서 **먼저 나온 주석**을 고른다. 고른 것 말고 몇 개가 더 있었는지를
`(후보 n)` 으로 남기므로, 금액이 여러 곳에 있어 애매하다는 사실이 숨지 않는다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from decimal import ROUND_HALF_UP, Decimal

from .check import (
    PERIOD_NAMES,
    ROUNDING,
    _ce_name,
    _ce_plan,
    _label,
    _nospace,
    parse_amount,
)
from .model import note_heading
from .text import DEFAULT_PATTERN, format_number

#: 판정.
MATCHED = "대사"
MISSING = "미대사"
#: 요약에 세는 순서 (`ROUNDING` 은 합계검증과 같은 `단수차` 다).
XREF_VERDICTS = (MATCHED, ROUNDING, MISSING)

#: 부호 관계.
SIGN_SAME = "같은 부호"
SIGN_OPPOSITE = "부호 반대"

#: 기본 단위와 배수 (원 단위 환산). 구현하는 것은 이 셋뿐이다.
DEFAULT_UNIT = "원"
UNIT_SCALES = {"원": 1, "천원": 1000, "백만원": 1000000}
#: `(단위 : 원)` · `(단위: 천원)` — `단위` 뒤에 한글 아닌 글자 몇 개를 건너뛰고 단위 이름이 온다.
_RE_UNIT = re.compile(r"단위[^가-힣]{0,4}(백만원|천원|원)")

#: 라벨 앞의 번호 — `1.` · `1)` · `(1)` · `Ⅰ.`. 숫자와 로마숫자만 지운다 (한글 `가.` 는 두지 않는다 —
#: `가.` 를 지우면 `가지급금` 같은 과목명을 다칠 위험이 있다).
_RE_NUMBERING = re.compile(
    r"^\s*(?:[(（]\s*\d{1,3}\s*[)）]|\d{1,3}\s*[.)]|[IVXⅠ-Ⅹ]{1,5}\s*\.)\s*"
)

#: `-v` 에서 한 검사에 보여 줄 후보 수의 상한 (금액 `1` 처럼 흔한 값은 후보가 수십 개다).
VERBOSE_CANDIDATES = 10

#: 상호대사에서 제외하는 섹션 — 이름(공백 제거)에 이것이 있으면 주석으로 보지 않는다.
EXCLUDED_SECTION = "외부감사"

#: 재무제표 금액이 하나도 없을 때 알리는 줄.
NOTE_NO_FS = (
    "안내: 상호대사할 재무제표 금액이 없습니다 — 합계검증 대상 표(ACODE/ALEVEL)를 찾지 못했습니다."
)
#: 주석 표가 하나도 없을 때 알리는 줄.
NOTE_NO_NOTES = (
    "안내: 상호대사할 주석 표가 없습니다 — 주석 번호 제목(`3. 사용제한예금` 꼴) 아래의 표를 "
    "찾지 못했습니다."
)


def strip_numbering(text: str) -> str:
    """라벨 앞의 번호를 지운다 (`1. 현금및현금성자산` → `현금및현금성자산`).

    다 지우고 나면 빈 문자열이 되는 라벨(`(1)` 뿐인 칸)은 원래 값을 그대로 둔다.
    """
    out = text
    while True:
        nxt = _RE_NUMBERING.sub("", out)
        if nxt == out:
            break
        out = nxt
    return out.strip() or text


def unit_of(text: str) -> str:
    """문자열에서 `단위 … 천원` 을 읽는다. 없으면 빈 문자열."""
    m = _RE_UNIT.search(_nospace(text))
    return m.group(1) if m else ""


def scale(value: Decimal, unit: str) -> Decimal:
    """표에 적힌 값 → **원 단위** 값."""
    return value * UNIT_SCALES.get(unit, 1)


# ─────────────────────────────────────────────────────────────────────────────
# 주석 표 모으기
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class NoteCell:
    """주석 표의 금액 셀 하나."""

    note: int               # 주석 번호
    title: str              # 주석 제목
    index: int              # 그 주석 안에서 몇 번째 표인가 (1-기반)
    sheet: str              # 섹션(시트) 이름
    table: str              # 표 요소 ID
    row: str                # 행 요소 ID
    label: str              # 행 라벨 (그 행의 첫 글자 칸)
    header: str             # 열 머리글 (2단이면 `-` 로 이음)
    unit: str               # 표에 적힌 단위
    value: Decimal          # 표에 적힌 값
    amount: Decimal         # 원 단위로 환산한 값
    seq: int = 0            # 문서 순서 (후보 순위가 같을 때의 갈림)

    @property
    def where_ko(self) -> str:
        title = f" {self.title}" if self.title else ""
        parts = [f"주석 {self.note}{title}", f"표{self.index}"]
        if self.label:
            parts.append(self.label)
        if self.header:
            parts.append(self.header)
        return " / ".join(parts)

    def as_dict(self) -> dict:
        return {
            "note": self.note,
            "title": self.title,
            "table_index": self.index,
            "sheet": self.sheet,
            "table": self.table,
            "row": self.row,
            "label": self.label,
            "header": self.header,
            "unit": self.unit,
            "value": str(self.value),
            "amount": str(self.amount),
        }


def _grid(rows) -> dict:
    """`(행 인덱스, 셀 인덱스)` → `(시작 열, 폭)`.

    `COLSPAN`/`ROWSPAN` 을 반영한 격자다. 주석 표는 `ADELIM` 이 없어(`seq` 배치)
    셀 순서만으로는 열을 알 수 없다 — `합  계`(COLSPAN=2) 다음 칸이 몇 번째 열인지는 앞 칸들이
    차지한 폭을 세어야 나온다.
    """
    taken: set = set()
    out: dict = {}
    for ri, row in enumerate(rows):
        col = 0
        for ci, cell in enumerate(row.cells):
            while (ri, col) in taken:
                col += 1
            cspan, rspan = max(1, cell.colspan), max(1, cell.rowspan)
            for dr in range(rspan):
                for dc in range(cspan):
                    taken.add((ri + dr, col + dc))
            out[(ri, ci)] = (col, cspan)
            col += cspan
    return out


def _headers(rows, grid: dict) -> dict:
    """열 → 머리글 (`THEAD` 행들의 글자를 위에서부터 `-` 로 이어 붙인다).

    단위 칸(`(단위: 천원)`)은 머리글이 아니다 — 많은 주석 표가 머리글 첫 줄 오른쪽 끝에
    단위를 적어 두는데, 그것까지 이으면 `(단위: 원)-당기말` 처럼 읽기 어려워진다.
    """
    out: dict = {}
    for ri, row in enumerate(rows):
        if row.kind != "THEAD-TR":
            continue
        for ci, cell in enumerate(row.cells):
            start, span = grid[(ri, ci)]
            name = _label(cell.text)
            if not name or unit_of(cell.text):
                continue
            for col in range(start, start + span):
                out[col] = f"{out[col]}-{name}" if out.get(col) else name
    return out


def _row_label(row) -> str:
    """행 라벨 = 그 행에서 **처음 나오는 글자 칸**. 없으면 첫 칸."""
    for cell in row.cells:
        text = _label(cell.text)
        if text and parse_amount(cell.text) is None:
            return strip_numbering(text)
    return strip_numbering(_label(row.cells[0].text)) if row.cells else ""


def _table_cells(table, sheet: str, note: int, title: str, index: int, unit: str) -> list:
    """표 하나의 금액 셀 (`0` 과 값 없음은 뺀다)."""
    grid = _grid(table.rows)
    heads = _headers(table.rows, grid)
    out = []
    for ri, row in enumerate(table.rows):
        if row.kind == "THEAD-TR":
            continue
        label = _row_label(row)
        for ci, cell in enumerate(row.cells):
            value = parse_amount(cell.text)
            if value is None or value == 0:
                continue
            start, _span = grid[(ri, ci)]
            out.append(NoteCell(
                note, title, index, sheet, table.eid, row.eid, label,
                heads.get(start, ""), unit, value, scale(value, unit),
            ))
    return out


def _table_text(table) -> str:
    return " ".join(c.text for r in table.rows for c in r.cells)


def collect(model, skip_tables=()) -> tuple:
    """문서 모델에서 `(주석 금액 셀 목록, 표 요소 ID → 단위)` 를 만든다.

    `skip_tables` 는 재무제표 쪽(합계검증 대상)의 표 요소 ID 다 — 주석 쪽에서 두 번 세지 않는다.
    단위는 **모든 표**에 대해 돌려준다 (재무제표 표의 `(단위 : 원)` 도 같은 규칙으로 읽는다).
    """
    skip = set(skip_tables)
    cells: list = []
    units: dict = {}
    for section in model.sections:
        note, title, unit, index = 0, "", DEFAULT_UNIT, 0
        excluded = EXCLUDED_SECTION in _nospace(section.sheet_name)
        for block in section.blocks:
            if block.table is None:
                heading = note_heading(block.text)
                if heading is not None and heading[0] > note:
                    note, title, unit, index = heading[0], heading[1], DEFAULT_UNIT, 0
                found = unit_of(block.text)
                if found:
                    unit = found
                continue
            found = unit_of(_table_text(block.table))
            if found:
                unit = found
            units[block.table.eid] = unit
            if excluded or not note or block.table.eid in skip:
                continue
            got = _table_cells(block.table, section.sheet_name, note, title, index + 1, unit)
            if got:
                index += 1
                cells.extend(got)
    # 문서 순서를 셀에 새긴다 — 후보 순위가 같으면 먼저 나온 주석을 고르기 위한 갈림이다.
    return [replace(c, seq=i) for i, c in enumerate(cells)], units


# ─────────────────────────────────────────────────────────────────────────────
# 재무제표 쪽 금액
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class FsAmount:
    """재무제표 표의 금액 하나."""

    sheet: str
    table: str
    row: str
    label: str
    acode: str
    period: str             # `당기`/`전기`, 자본변동표는 열 이름
    value: Decimal
    unit: str
    amount: Decimal         # 원 단위로 환산한 값


def fs_amounts(tables, units) -> list:
    """합계검증 대상 표에서 금액 셀을 모은다 (`0` 과 값 없음은 뺀다)."""
    out = []
    for table in tables:
        if table.excluded:
            continue
        unit = units.get(table.eid, DEFAULT_UNIT)
        plan = _ce_plan(table)
        for row in table.rows:
            label = strip_numbering(row.label)
            if plan is not None:
                pairs = [(_ce_name(plan, c), row.vals.get(c)) for c in plan.cols]
            else:
                pairs = [(name, row.amount(p)) for p, name in enumerate(PERIOD_NAMES)]
            for period, value in pairs:
                if value is None or value == 0:
                    continue
                out.append(FsAmount(table.sheet, table.eid, row.eid, label, row.acode,
                                    period, value, unit, scale(value, unit)))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 대사
# ─────────────────────────────────────────────────────────────────────────────


def _shared(a: str, b: str) -> int:
    """두 라벨이 공유하는 글자 수 (공백·구두점 제외)."""
    keep = {c for c in _nospace(a) if c.isalnum()}
    return len(keep & {c for c in _nospace(b) if c.isalnum()})


def _rank(fs: FsAmount, cands: list) -> list:
    """후보를 라벨이 비슷한 순으로. 점수가 같으면 문서 순서(먼저 나온 주석)를 지킨다."""
    return sorted(cands, key=lambda c: (-_shared(fs.label, c.label + c.title), c.seq))


@dataclass(frozen=True)
class XrefItem:
    """상호대사 한 건."""

    fs: FsAmount
    verdict: str
    match: NoteCell | None = None
    others: int = 0
    sign: str = ""
    diff: Decimal = Decimal(0)
    candidates: tuple = ()

    @property
    def where_ko(self) -> str:
        return (f"{_label(self.fs.sheet)} {self.fs.label} {self.fs.period} "
                f"{format_number(self.fs.amount, DEFAULT_PATTERN)}")

    def line_ko(self) -> str:
        if self.match is None:
            return f"[상호대사] {self.where_ko} → {MISSING}"
        extra = []
        if self.match.unit != DEFAULT_UNIT:
            extra.append(f"단위 {self.match.unit}")
        if self.others:
            extra.append(f"후보 {self.others}")
        if self.sign == SIGN_OPPOSITE:
            extra.append(SIGN_OPPOSITE)
        if self.verdict == ROUNDING:
            extra.append(f"{ROUNDING} {format_number(self.diff, DEFAULT_PATTERN)}")
        tail = f" ({', '.join(extra)})" if extra else ""
        return f"[상호대사] {self.where_ko} → {self.match.where_ko}{tail}"

    def candidate_lines_ko(self) -> list:
        """`-v` 용 — 후보 전부 (상한 `VERBOSE_CANDIDATES`)."""
        lines = []
        for c in self.candidates[:VERBOSE_CANDIDATES]:
            lines.append(f"    후보: {c.where_ko} "
                         f"{format_number(c.value, DEFAULT_PATTERN)}"
                         f"{'' if c.unit == DEFAULT_UNIT else ' ' + c.unit}")
        rest = len(self.candidates) - VERBOSE_CANDIDATES
        if rest > 0:
            lines.append(f"    후보 {rest}개 더 있습니다")
        return lines

    def as_dict(self) -> dict:
        return {
            "sheet": self.fs.sheet,
            "table": self.fs.table,
            "row": self.fs.row,
            "label": self.fs.label,
            "acode": self.fs.acode,
            "period": self.fs.period,
            "amount": str(self.fs.amount),
            "unit": self.fs.unit,
            "verdict": self.verdict,
            "diff": str(self.diff),
            "sign": self.sign,
            "others": self.others,
            "match": self.match.as_dict() if self.match is not None else None,
            "candidates": [c.as_dict() for c in self.candidates],
        }


def _sign(a: Decimal, b: Decimal) -> str:
    return SIGN_SAME if (a > 0) == (b > 0) else SIGN_OPPOSITE


def rounded(amount: Decimal, mult: int) -> Decimal:
    """원 단위 금액을 배수 `mult` 의 단위로 **반올림**한다 (Excel `ROUND(x/10^3,)` 과 같다).

    주석이 천원 단위면 원 단위 금액은 이미 **끊어 적힌** 값이다 — 샘플 `현금및현금성자산`
    이 재무상태표에 `1,234,567,890` 원, 주석에 `1,234,567` 천원이면 주석 값을 원으로
    되돌려(× 1,000) 견줄 때 890 원이 벌어져 영영 만나지 못하므로, **재무제표 금액을 주석의
    단위로 내려** 견준다.
    """
    if mult == 1:
        return amount
    return (amount / mult).quantize(Decimal(1), rounding=ROUND_HALF_UP)


def _match_one(fs: FsAmount, by_unit: dict, tolerance: int) -> XrefItem:
    """`fs` 하나를 주석 셀에서 찾는다. `by_unit` 은 `{배수: {절댓값: [셀]}}` 이다."""
    base = abs(fs.amount)
    targets = {mult: rounded(base, mult) for mult in by_unit}
    cands, verdict = [], MATCHED
    for mult, table in by_unit.items():
        cands.extend(table.get(targets[mult], ()))
    if not cands:
        verdict = ROUNDING
        for mult, table in by_unit.items():
            for d in range(-tolerance, tolerance + 1):
                if d:
                    cands.extend(table.get(targets[mult] + d, ()))
    if not cands:
        return XrefItem(fs, MISSING)
    ranked = _rank(fs, cands)
    best = ranked[0]
    diff = targets[UNIT_SCALES.get(best.unit, 1)] - abs(best.value)
    return XrefItem(fs, verdict, best, len(ranked) - 1,
                    _sign(fs.amount, best.value), diff, tuple(ranked))


@dataclass
class XrefReport:
    """상호대사 결과 — 종료 코드에는 영향을 주지 않는다."""

    tolerance: int
    min_amount: int = 0
    items: list = field(default_factory=list)
    fs_total: int = 0            # 재무제표 금액 수 (`--xref-min` 으로 거른 것 포함)
    note_cells: int = 0
    note_tables: int = 0

    def counts(self) -> dict:
        out = {v: 0 for v in XREF_VERDICTS}
        for it in self.items:
            out[it.verdict] += 1
        return out

    def summary_ko(self) -> str:
        c = self.counts()
        return (f"상호대사: {MATCHED} {c[MATCHED]} / {ROUNDING} {c[ROUNDING]} / "
                f"{MISSING} {c[MISSING]}")

    def lines_ko(self, verbose: bool = False) -> list:
        lines = []
        if verbose:
            skipped = self.fs_total - len(self.items)
            lines.append(
                f"상호대사 대상: 재무제표 금액 {len(self.items)}개"
                + (f" (금액 {format_number(Decimal(self.min_amount), DEFAULT_PATTERN)} 미만 "
                   f"{skipped}개 제외)" if skipped else "")
                + f" / 주석 표 {self.note_tables}개 · 금액 {self.note_cells}개"
            )
        misses = []
        for it in self.items:
            if it.match is None:
                misses.append(it)
                continue
            lines.append(it.line_ko())
            if verbose:
                lines.extend(it.candidate_lines_ko())
        if misses:
            lines.append(f"{MISSING} {len(misses)}건")
            lines.extend(it.line_ko() for it in misses)
        if not self.fs_total:
            lines.append(NOTE_NO_FS)
        elif not self.note_cells:
            lines.append(NOTE_NO_NOTES)
        lines.append(self.summary_ko())
        return lines

    def json_obj(self) -> list:
        return [it.as_dict() for it in self.items]


def build(tables, model, tolerance: int, min_amount: int = 0) -> XrefReport:
    """합계검증 대상 표 + 문서 모델 → `XrefReport`."""
    skip = {t.eid for t in tables if not t.excluded}
    cells, units = collect(model, skip)
    amounts = fs_amounts(tables, units)
    by_unit: dict = {}
    for c in cells:
        by_unit.setdefault(UNIT_SCALES.get(c.unit, 1), {}).setdefault(
            abs(c.value), []).append(c)
    report = XrefReport(tolerance, min_amount, [], len(amounts), len(cells),
                        len({c.table for c in cells}))
    for fs in amounts:
        if abs(fs.amount) < min_amount:
            continue
        report.items.append(_match_one(fs, by_unit, tolerance))
    return report


__all__ = [
    "MATCHED",
    "MISSING",
    "XREF_VERDICTS",
    "SIGN_SAME",
    "SIGN_OPPOSITE",
    "DEFAULT_UNIT",
    "UNIT_SCALES",
    "EXCLUDED_SECTION",
    "NOTE_NO_FS",
    "NOTE_NO_NOTES",
    "VERBOSE_CANDIDATES",
    "NoteCell",
    "FsAmount",
    "XrefItem",
    "XrefReport",
    "strip_numbering",
    "unit_of",
    "scale",
    "collect",
    "fs_amounts",
    "build",
]
