"""문서 구조 지도 — `.dsd` 와 내보낸 `.xlsx` 를 **같은 모양**으로 읽는다 (1.3 flow 도구).

`describe` · `diff` · `check --settlement` 이 공유하는 읽기 전용 계층이다. 모델(`dsd_bridge.model`)·
export·import 는 건드리지 않는다(규범 2).

구조::

    DocMap ─ MapSheet ─ MapTable ─ MapRow ─ MapCell
                      └ MapBlock (표가 아닌 `P`/`TITLE` 행)

논리 열
-------
셀의 논리 열은 `dsd_bridge.check` 와 같은 규칙으로 정한다 — 표 안의 **모든** 셀에 `ADELIM` 이 있으면
그 값이고, 하나라도 없으면 `TR` 안의 셀 순서다(첫 셀 = 과목 = 0). 표준 재무제표 표는
`0`=과목, `1·2`=당기(세부·합계), `3·4`=전기(세부·합계) 다 (`COL_ROLES`).

기간표 판정 (`MapTable.is_period`)
----------------------------------
이월(roll-forward) 검증과 열 역할 표시는 **기간표**에만 한다. 판정은

1. 본문 논리 열이 `0~4` 안에 있고 `4` 를 쓴다, 그리고
2. `ADELIM` 배치이거나(= DART 가 표준 재무제표로 표시한 표), 머리글(`THEAD`)에 당기 표기와
   전기 표기가 모두 있다.

`ADELIM` 이 없는 사모투자회사 양식은 머리글 `제 N(당) 기` /
`제 N-1(전) 기` 로 걸린다. 자본변동표(머리글이 자본 구성요소)와 주석 표(`구분|수량|…`)는
걸리지 않는다 — 그 표들을 이월 검증에 넣으면 엉뚱한 열을 견주게 된다.

기간 표기
---------
`PERIOD_PATTERNS` 는 기간(기수·날짜)을 나타내는 문구다. 마지막 `(당|전)(반)?기(초|말)?` 는
**뒤에 한글이 더 붙지 않을 때만** 잡는다 — `당기순이익`·`전기순이익` 처럼 기간 표기가 아닌
계정과목까지 잡으면 갱신 대상 목록이 쓸모없어지기 때문이다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .check import CheckRow, parse_amount
from .errors import InputError
from .model import LAYOUT_ADELIM, LAYOUT_SEQ
from .sheetstate import FIRST_VALUE_COL
from .text import cr_to_lf

#: 표준 재무제표 표의 논리 열.
STD_COLS = frozenset({0, 1, 2, 3, 4})
#: 논리 열의 역할 이름.
COL_ROLES = {0: "과목", 1: "당기 세부", 2: "당기 합계", 3: "전기 세부", 4: "전기 합계"}
#: 기간 이름 (0=당기, 1=전기) — `dsd_bridge.check.PERIOD_NAMES` 와 같다.
PERIOD_NAMES = ("당기", "전기")

_RE_WS = re.compile(r"[\s　]+")

#: 기간 표기 정규식 `(이름, 패턴)` — `describe` 의 "기간 표기 셀" 과 `diff` 의 기간 표기 검사에 쓴다.
PERIOD_PATTERNS = (
    ("기수", re.compile(r"제\s*\d+\s*\(\s*(?:당|전)\s*\)\s*기")),
    ("연월일", re.compile(r"\d{4}\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일")),
    ("기초기말", re.compile(r"\d{4}\.\d{2}\.\d{2}\s*\(\s*(?:당|전)기(?:초|말)\s*\)")),
    ("기간어", re.compile(r"(?:당|전)(?:반)?기(?:초|말)?(?![가-힣])")),
)

#: `제 11(당) 기` 에서 기수와 당·전을 뽑는다 (기간 표기 갱신 검사).
RE_PERIOD_NO = re.compile(r"제\s*(\d+)\s*\(\s*(당|전)\s*\)\s*기")

#: 머리글이 당기/전기를 나타내는가 (기간표 판정).
RE_CUR_HEAD = re.compile(r"\(\s*당\s*\)|당\s*반?\s*기")
RE_PREV_HEAD = re.compile(r"\(\s*전\s*\)|전\s*반?\s*기")


def excel_col_letter(offset: int) -> str:
    """물리 열 오프셋(F=0)을 Excel 열 문자로."""
    from openpyxl.utils import get_column_letter

    return get_column_letter(FIRST_VALUE_COL + max(0, int(offset)))


def period_hits(text) -> list:
    """텍스트에서 기간 표기를 찾는다 — `[(패턴 이름, 맞은 문구), …]` (중복 제거, 순서 유지)."""
    s = cr_to_lf("" if text is None else str(text))
    out, seen = [], set()
    for name, rx in PERIOD_PATTERNS:
        for m in rx.finditer(s):
            key = (name, m.group(0))
            if key not in seen:
                seen.add(key)
                out.append(key)
    return out


def period_numbers(text) -> list:
    """`제 N(당) 기` 표기 `[(N, "당"|"전"), …]`."""
    s = cr_to_lf("" if text is None else str(text))
    return [(int(m.group(1)), m.group(2)) for m in RE_PERIOD_NO.finditer(s)]


# ─────────────────────────────────────────────────────────────────────────────
# 라벨 정규화 (대사 스크립트의 `norm` 을 일반화)
# ─────────────────────────────────────────────────────────────────────────────

_RE_NOTE_REF = re.compile(r"[<\(（]\s*주\s*석\s*[\d, ]*[>\)）]")
_RE_DATE_KO = re.compile(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일")
_RE_ROMAN = re.compile(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXivx]+[.)]")
_RE_NUMBERING = re.compile(r"^\(?\d{1,2}[.)]")
_DROP_WORDS = ("(또는당기순손실)", "(또는당기순이익)", "(손실)", "(이익)")


def first_line(text) -> str:
    """첫 **비어 있지 않은** 줄. 값이 없으면 빈 문자열.

    예시 스크립트는 `split("\\n")[0]` 이었는데, `&cr;` 로 시작하는 셀
    (`\\n   매출채권 <주석4>`)이 빈 문자열이 되어 대사에서 빠졌다. 첫 비어 있지 않은
    줄을 쓰면 `20X1.12.31(전기말)&cr;(발행주식수 : …)` 같은 두 줄 셀도 그대로 잘 읽힌다.
    """
    if text is None:
        return ""
    for line in cr_to_lf(str(text)).splitlines():
        if line.strip():
            return line
    return ""


def norm_label(text) -> str:
    """계정과목 정규화 — 보고서와 정산표, 전기 보고서와 당기 보고서를 같은 잣대로 견준다.

    공백(전각 포함) 제거 · `<주석3>`/`(주석 6)` 제거 · `2026년 06월 30일` → `2026.06.30` ·
    `Ⅰ.`/`(1)`/`1.` 같은 번호 접두 제거 · `및` → `와` · `(손실)` 류 제거 · 끝의 `:` 제거.
    """
    s = first_line(text)
    if not s:
        return ""
    s = _RE_DATE_KO.sub(lambda m: f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}", s)
    s = _RE_NOTE_REF.sub("", s)
    s = _RE_WS.sub("", s)
    s = _RE_ROMAN.sub("", s)
    s = _RE_NUMBERING.sub("", s)
    s = s.replace("및", "와")
    for w in _DROP_WORDS:
        s = s.replace(w, "")
    return s.rstrip(":：")


def display_label(text) -> str:
    """사람이 읽을 라벨 — 첫 줄의 공백을 하나로 줄인다."""
    return " ".join(first_line(text).split())


def similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher

    return SequenceMatcher(None, a, b).ratio()


# ─────────────────────────────────────────────────────────────────────────────
# 구조
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class MapCell:
    """셀 하나."""

    col: int              # 논리 열
    excel_col: int        # 물리 열 (F=0)
    value: object = None  # `.xlsx` 는 openpyxl 값, `.dsd` 는 원문 문자열

    @property
    def text(self) -> str:
        return "" if self.value is None else str(self.value)

    @property
    def amount(self):
        return parse_amount(self.value)

    @property
    def letter(self) -> str:
        return excel_col_letter(self.excel_col)


@dataclass(frozen=True)
class MapRow:
    """표 행 하나."""

    eid: str
    kind: str             # "TR" | "THEAD-TR"
    excel_row: int        # `.xlsx` 입력이면 Excel 행 번호, `.dsd` 면 0
    cells: tuple = ()

    @property
    def label(self) -> str:
        for c in self.cells:
            if c.col == 0:
                return display_label(c.value)
        return display_label(self.cells[0].value) if self.cells else ""

    @property
    def norm(self) -> str:
        for c in self.cells:
            if c.col == 0:
                return norm_label(c.value)
        return norm_label(self.cells[0].value) if self.cells else ""

    @property
    def vals(self) -> dict:
        """논리 열 → 금액 (값이 없는 열은 키가 없다)."""
        out = {}
        for c in self.cells:
            a = c.amount
            if a is not None and c.col not in out:
                out[c.col] = a
        return out

    def cell_of(self, col: int):
        for c in self.cells:
            if c.col == col:
                return c
        return None

    def amount(self, period: int):
        """`period` (0=당기, 1=전기) 의 금액 — `dsd_bridge.check.CheckRow` 와 같은 규칙."""
        return CheckRow(self.eid, self.label, None, "", self.vals).amount(period)

    @property
    def has_amount(self) -> bool:
        return bool(self.vals)


@dataclass(frozen=True)
class MapTable:
    """표 하나."""

    sheet: str
    eid: str
    index: int            # 시트 안에서의 표 순번 (0-기반)
    layout: str = LAYOUT_SEQ
    aclass: str = ""
    headers: tuple = ()   # THEAD 행
    rows: tuple = ()      # 본문 행
    start_row: int = 0    # `.xlsx` 의 `TABLE-START` 행 번호
    end_row: int = 0

    @property
    def logical_cols(self) -> frozenset:
        return frozenset(c.col for r in self.rows for c in r.cells)

    @property
    def header_text(self) -> str:
        return " | ".join(
            " ".join(c.text.split()) for h in self.headers for c in h.cells if c.text.strip()
        )

    @property
    def is_period(self) -> bool:
        """표준 재무제표(당기·전기) 표인가 — 모듈 docstring 의 판정."""
        cols = self.logical_cols
        if not cols <= STD_COLS or 4 not in cols:
            return False
        if self.layout == LAYOUT_ADELIM:
            return True
        heads = " ".join(c.text for h in self.headers for c in h.cells)
        return bool(RE_CUR_HEAD.search(heads) and RE_PREV_HEAD.search(heads))

    def col_letters(self) -> dict:
        """논리 열 → Excel 열 문자 (본문 행에서 실제로 쓰인 물리 열)."""
        out = {}
        for r in self.rows:
            for c in r.cells:
                out.setdefault(c.col, excel_col_letter(c.excel_col))
        return out


@dataclass(frozen=True)
class MapBlock:
    """표가 아닌 블록 행 (`P` / `TITLE`)."""

    eid: str
    kind: str
    excel_row: int
    text: str


@dataclass(frozen=True)
class MapSheet:
    name: str
    order: int
    items: tuple = ()     # MapTable | MapBlock, 문서 순서

    @property
    def tables(self) -> list:
        return [i for i in self.items if isinstance(i, MapTable)]

    @property
    def blocks(self) -> list:
        return [i for i in self.items if isinstance(i, MapBlock)]


@dataclass(frozen=True)
class DocMap:
    path: Path
    source: str           # "dsd" | "xlsx"
    contents_sha256: str = ""
    sheets: tuple = ()

    def sheet(self, name: str):
        for s in self.sheets:
            if s.name == name:
                return s
        return None

    @property
    def tables(self) -> list:
        return [t for s in self.sheets for t in s.tables]

    def row_by_id(self, eid: str):
        for s in self.sheets:
            for t in s.tables:
                for r in t.rows:
                    if r.eid == eid:
                        return s, t, r
        return None


# ─────────────────────────────────────────────────────────────────────────────
# 구축
# ─────────────────────────────────────────────────────────────────────────────


def _use_adelim(cell_delims) -> bool:
    """표 안의 **모든** 셀에 `ADELIM` 이 있으면 그것을 논리 열로 쓴다 (`dsd_bridge.check` 와 같은 규칙)."""
    delims = list(cell_delims)
    return bool(delims) and all(str(d).strip().isdigit() for d in delims)


def from_model(model, path, *, source: str = "dsd", sha: str = "") -> DocMap:
    """문서 모델(`dsd_bridge.model.DocModel`) → `DocMap`."""
    from .model import table_layout

    sheets = []
    for order, section in enumerate(model.sections):
        items = []
        tindex = 0
        for block in section.blocks:
            if block.table is None:
                items.append(MapBlock(block.eid, block.kind, 0, block.display))
                continue
            t = block.table
            layout = table_layout(t)
            body = [r for r in t.rows if r.kind != "THEAD-TR"]
            use = _use_adelim(c.adelim for r in body for c in r.cells)
            heads, rows = [], []
            for r in t.rows:
                cols = layout.cols_for(r.eid)
                cells = tuple(
                    MapCell(
                        col=(int(c.adelim.strip()) if use and r.kind != "THEAD-TR" else i),
                        excel_col=(cols[i] if i < len(cols) else i),
                        value=c.display,
                    )
                    for i, c in enumerate(r.cells)
                )
                row = MapRow(r.eid, r.kind, 0, cells)
                (heads if r.kind == "THEAD-TR" else rows).append(row)
            items.append(MapTable(
                sheet=section.sheet_name, eid=t.eid, index=tindex, layout=layout.kind,
                aclass=t.aclass, headers=tuple(heads), rows=tuple(rows),
            ))
            tindex += 1
        sheets.append(MapSheet(section.sheet_name, order, tuple(items)))
    return DocMap(Path(path), source, sha, tuple(sheets))


def from_dsd(path) -> DocMap:
    """`.dsd` → `DocMap`."""
    from .check import model_from_dsd
    from .container import DsdContainer

    model = model_from_dsd(path)
    try:
        sha = DsdContainer.read(Path(path)).contents_sha256()
    except Exception:                       # pragma: no cover - model_from_dsd 가 먼저 잡는다
        sha = ""
    return from_model(model, path, source="dsd", sha=sha)


def from_xlsx(path) -> DocMap:
    """내보낸 `.xlsx` → `DocMap`. `_dsdx_attrs` 가 없으면 `ADELIM` 없이(셀 순서로) 읽는다."""
    from .attrs import read_attrs_sheet
    from .sheetstate import load_workbook_state

    state = load_workbook_state(path)
    by_id = {d["_ID"]: d for d in read_attrs_sheet(path) if d.get("_ID")}
    layouts = {}
    for entry in state.section_map:
        for t in entry.get("tables") or []:
            if isinstance(t, dict) and t.get("id"):
                layouts[t["id"]] = str(t.get("layout") or LAYOUT_SEQ)

    sheets = []
    for sheet in state.sheets:
        items = []
        tindex = 0
        pending_heads: list = []
        pending_rows: list = []
        table_eid, start_row = "", 0
        for r in sheet.rows:
            if r.is_header:
                continue
            if r.base_kind == "TABLE-START":
                pending_heads, pending_rows = [], []
                table_eid, start_row = r.table, r.row
                continue
            if r.base_kind == "TABLE-END":
                if table_eid:
                    items.append(_xlsx_table(sheet.name, table_eid, tindex, layouts,
                                             by_id, pending_heads, pending_rows,
                                             start_row, r.row))
                    tindex += 1
                pending_heads, pending_rows = [], []
                table_eid, start_row = "", 0
                continue
            if r.base_kind in ("P", "TITLE"):
                text = "" if not r.values else ("" if r.values[0] is None else str(r.values[0]))
                items.append(MapBlock(r.id, r.base_kind, r.row, text))
                continue
            if r.base_kind in ("TR", "THEAD-TR") and table_eid:
                (pending_heads if r.base_kind == "THEAD-TR" else pending_rows).append(r)
        if table_eid:                       # `TABLE-END` 가 지워진 통합문서
            items.append(_xlsx_table(sheet.name, table_eid, tindex, layouts, by_id,
                                     pending_heads, pending_rows, start_row, 0))
        sheets.append(MapSheet(sheet.name, sheet.order, tuple(items)))
    sha = str((state.meta or {}).get("contents_sha256") or "")
    return DocMap(Path(path), "xlsx", sha, tuple(sheets))


def _delims_of(attrs: dict, n: int) -> list:
    raw = (attrs or {}).get("adelims") or ""
    parts = raw.split("|") if raw else []
    return [parts[i] if i < len(parts) else "" for i in range(n)]


def _xlsx_table(sheet: str, eid: str, index: int, layouts: dict, by_id: dict,
                heads: list, rows: list, start_row: int, end_row: int) -> MapTable:
    use = _use_adelim(
        d for r in rows for d in _delims_of(by_id.get(r.id), len(r.cell_ids))
    )
    aclass = ""
    for r in rows:
        a = by_id.get(r.id)
        if a and a.get("aclass"):
            aclass = a["aclass"]
            break

    def build(r, thead: bool) -> MapRow:
        ids = r.cell_ids
        delims = _delims_of(by_id.get(r.id), len(ids))
        cells = tuple(
            MapCell(
                col=(int(delims[i].strip()) if use and not thead and delims[i].strip().isdigit()
                     else i),
                excel_col=r.col_of(i),
                value=r.value_for_cell(i),
            )
            for i in range(len(ids))
        )
        return MapRow(r.id, r.base_kind, r.row, cells)

    return MapTable(
        sheet=sheet, eid=eid, index=index,
        layout=layouts.get(eid, LAYOUT_SEQ), aclass=aclass,
        headers=tuple(build(r, True) for r in heads),
        rows=tuple(build(r, False) for r in rows),
        start_row=start_row, end_row=end_row,
    )


def load(path) -> DocMap:
    """`.dsd` 또는 `.xlsx` 를 읽어 `DocMap` 을 만든다."""
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix == ".dsd":
        return from_dsd(p)
    if suffix == ".xlsx":
        return from_xlsx(p)
    raise InputError(f"오류: .dsd 또는 .xlsx 만 읽을 수 있습니다: {p}")


__all__ = [
    "STD_COLS",
    "COL_ROLES",
    "PERIOD_NAMES",
    "PERIOD_PATTERNS",
    "RE_PERIOD_NO",
    "MapCell",
    "MapRow",
    "MapTable",
    "MapBlock",
    "MapSheet",
    "DocMap",
    "excel_col_letter",
    "period_hits",
    "period_numbers",
    "first_line",
    "norm_label",
    "display_label",
    "similarity",
    "from_model",
    "from_dsd",
    "from_xlsx",
    "load",
]
