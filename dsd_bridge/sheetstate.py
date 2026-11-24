"""편집된 `.xlsx` 를 행 단위 상태로 읽는다 (계획 §5 단계 5 `_load_sheet_state`).

`load_workbook(path, data_only=True)` 로 연다 — 수식 셀은 마지막 계산값(캐시)을 준다.
숨김 A~E 열을 `_ID` / `_KIND` / `_TABLE` / `_CELLS` / `_SEQ` 로, F 열부터를 `values` 로 읽는다.
해석(행 분류·게이트)은 하지 않는다 — 그것은 `dsd_bridge.gates` 와 `dsd_bridge.import_xlsx` 의 몫이다.

**열 배치 (1.2 WP3)**: `values` 는 **물리 열**(F=0) 순서 그대로이고, `_CELLS` 는 **원문 셀**
순서다. 두 순서는 `adelim` 배치의 표에서 달라지므로, 읽은 뒤 `_dsdx.section_map[].tables[]`
의 `layout`/`cols`/`rowcols` 를 보고 행마다 `SheetRow.cols`(셀 → 물리 열)를 채운다. 그 뒤로는

    r.value_for_cell(i)   # i 번째 셀(`_CELLS` 순서)의 Excel 값
    r.col_of(i)           # i 번째 셀이 놓인 물리 열 (F=0)
    r.mapped_cols         # 이 행에서 셀이 놓이는 물리 열 전부

만 쓰면 배치를 몰라도 된다. `_SEQ` 가 빈 삽입 행은 **앵커 행**(바로 위, 없으면 아래의 원본
표 행)의 배치를 물려받는다 — 삽입 행의 값도 앵커의 셀과 같은 논리 열로 짝지어진다.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from .errors import InputError
from .model import HEADER_KIND, HIDDEN_COLUMNS, RO_SUFFIX

META_SHEET = "_dsdx"
#: dsd-bridge 내부 시트의 이름 접두사 — `_dsdx`, `_dsdx_attrs`(1.2 WP2) 등. 이 접두사로 시작하는 시트는
#: 섹션(데이터) 시트가 아니므로 행 상태를 읽지 않고 게이트도 보지 않는다. 섹션 시트 이름은 원문
#: 제목에서 오므로(`sheetname.make_sheet_name`) 이 접두사와 부딪히지 않는다.
INTERNAL_SHEET_PREFIX = "_dsdx"
FIRST_VALUE_COL = len(HIDDEN_COLUMNS) + 1     # F
GUIDE_KINDS = frozenset({"TABLE-START", "TABLE-END"})
#: AC-28/29 의 검사 대상이 되는 `_KIND` (`:RO` 제거 후).
MAPPED_KINDS = frozenset({"TR", "THEAD-TR", "P", "TITLE"})


def is_internal_sheet(name) -> bool:
    """dsd-bridge 가 만든 내부 시트인가 (`_dsdx`, `_dsdx_attrs` …)."""
    return str(name).startswith(INTERNAL_SHEET_PREFIX)


def _is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v == "")


def _as_str(v) -> str:
    if v is None:
        return ""
    return v if isinstance(v, str) else str(v)


@dataclass
class SheetRow:
    """시트의 한 행 — 숨김 열 5개 + 표시 값."""

    row: int                          # 1-기반 Excel 행 번호
    id: str
    kind: str
    table: str
    cells: str
    seq: int | None                   # 빈 값이면 None
    seq_raw: object = None            # 원래 셀 값 (정수가 아닌 값 진단용)
    values: list = field(default_factory=list)   # F 열부터. 끝의 빈 값은 잘라낸다
    formulas: set = field(default_factory=set)   # 수식이 든 표시 셀의 0-기반 인덱스 (F=0)
    hidden_formula: bool = False      # 숨김 A~E열에 수식이 있다 (손상)
    cols: list = field(default_factory=list)     # 셀 → 물리 열 (F=0). 표 행에만, 비면 항등

    @property
    def base_kind(self) -> str:
        return self.kind.removesuffix(RO_SUFFIX)

    @property
    def is_header(self) -> bool:
        return self.base_kind == HEADER_KIND

    @property
    def is_guide(self) -> bool:
        return self.base_kind in GUIDE_KINDS

    @property
    def is_mapped_kind(self) -> bool:
        return self.base_kind in MAPPED_KINDS

    @property
    def seq_corrupt(self) -> bool:
        """E 열에 값이 있는데 정수로 읽히지 않는다 (숨김 열 손상 — 게이트가 코드 6 으로 잡는다)."""
        return self.seq is None and not _is_blank(self.seq_raw)

    @property
    def cell_ids(self) -> list:
        """`_CELLS` 의 셀 ID 목록 (`:RO` 제거)."""
        if not self.cells:
            return []
        return [c.removesuffix(RO_SUFFIX) for c in self.cells.split("|")]

    @property
    def ro_cell_ids(self) -> set:
        if not self.cells:
            return set()
        return {c.removesuffix(RO_SUFFIX) for c in self.cells.split("|") if c.endswith(RO_SUFFIX)}

    @property
    def rightmost_value_index(self) -> int:
        """값이 들어 있는 0-기반 최우측 **물리 열**. 값이 하나도 없으면 -1."""
        for i in range(len(self.values) - 1, -1, -1):
            if not _is_blank(self.values[i]):
                return i
        return -1

    def value_at(self, i: int):
        """**물리 열** i (F=0) 의 값."""
        return self.values[i] if i < len(self.values) else None

    def col_of(self, i: int) -> int:
        """`_CELLS` 의 i 번째 셀이 놓인 물리 열 (F=0). 배치가 없으면 항등."""
        return self.cols[i] if i < len(self.cols) else i

    def value_for_cell(self, i: int):
        """`_CELLS` 의 i 번째 셀에 해당하는 Excel 값."""
        return self.value_at(self.col_of(i))

    def mapped_cols(self, ncells: int) -> set:
        """이 행에서 셀이 놓이는 물리 열 전부 (`ncells` = 셀 수, 배치가 없을 때의 상한)."""
        return set(self.cols) if self.cols else set(range(ncells))

    def unmapped_value_col(self, ncells: int) -> int:
        """값이 들어 있는데 셀이 없는 **최우측** 물리 열. 없으면 -1 (AC-30).

        `seq` 배치에서는 `rightmost_value_index >= ncells` 와 같은 판정이고, `adelim` 배치에서는
        표 오른쪽 바깥뿐 아니라 **가운데의 빈 논리 열**에 넣은 값도 잡는다.
        """
        mapped = self.mapped_cols(ncells)
        for i in range(len(self.values) - 1, -1, -1):
            if not _is_blank(self.values[i]) and i not in mapped:
                return i
        return -1


@dataclass
class SheetState:
    name: str
    order: int
    rows: list = field(default_factory=list)

    def by_seq(self) -> dict:
        """`_SEQ` → 행 (중복은 뒤의 것이 덮는다 — 중복 검사는 게이트가 한다)."""
        return {r.seq: r for r in self.rows if r.seq is not None}

    @property
    def title_id(self) -> str:
        """이 시트가 나타내는 섹션의 제목 요소 ID — **첫 매핑 행**의 `_ID`.

        `project()` 는 시트의 첫 매핑 행(헤더 행 다음)을 언제나 섹션 제목으로 놓는다. 제목은
        `TITLE` 이거나(`ATOC="Y"`), `--split-notes` (1.2 WP5) 로 나눈 주석 시트라면 `P` 다.
        표 행(`TR`)이 먼저 오는 일은 없으므로 그런 시트는 제목 행을 지운 것이며, 빈 문자열을
        돌려주어 이름으로 찾게 한다 — 기존 판정(제목 삭제 금지 / 대량 소실)이 그대로 나온다.
        """
        for r in self.rows:
            if r.is_mapped_kind and r.seq is not None:
                return r.id if r.base_kind in ("TITLE", "P") else ""
        return ""


@dataclass(frozen=True)
class SectionSheet:
    """`section_map` 항목 하나를 시트에 대응시킨 결과 (AC-23)."""

    sheet: object | None              # SheetState | None (시트가 없으면 None)
    renamed: bool = False             # title_id 로 찾았는데 이름이 기록과 다르다
    conflict: str = ""                # 비어 있지 않으면 한국어 오류 메시지 (코드 4)


@dataclass
class WorkbookState:
    path: Path
    meta: dict | None                 # `_dsdx` 가 없으면 None
    sheets: list = field(default_factory=list)

    def sheet(self, name: str):
        for s in self.sheets:
            if s.name == name:
                return s
        return None

    def sheet_by_title_id(self, title_id: str):
        if not title_id:
            return None
        for s in self.sheets:
            if s.title_id == title_id:
                return s
        return None

    def section_sheet(self, entry: dict) -> SectionSheet:
        """`section_map` 항목의 시트를 찾는다 — **이름이 아니라 `title_id` 가 정체성**이다 (AC-23).

        1. 제목 요소 ID 가 같은 시트가 있으면 그 시트다 (이름이 바뀌어도 섹션은 유지된다).
        2. 없으면 기록된 이름의 시트로 되돌아간다 — 제목 행을 지운 경우가 여기 해당하며,
           삭제 판정(제목 삭제 금지 / 대량 소실)은 호출 측이 그대로 낸다.
        3. 이름은 같은데 그 시트의 제목이 **다른 섹션**의 제목이면 짝이 어긋난 것이므로
           `conflict` 에 한국어 오류 메시지를 담는다 (게이트는 판정만, import 가 코드 4).
        """
        title_id = str(entry.get("title_id") or "")
        name = entry["sheet"]
        found = self.sheet_by_title_id(title_id)
        if found is not None:
            return SectionSheet(found, renamed=found.name != name)
        by_name = self.sheet(name)
        if by_name is None:
            return SectionSheet(None)
        other = by_name.title_id
        if other and other != title_id and any(
            str(e.get("title_id") or "") == other for e in self.section_map
        ):
            return SectionSheet(None, conflict=(
                f"오류: 시트 '{name}' 은(는) 다른 섹션(제목 요소 {other})의 시트입니다. "
                f"섹션(제목 요소 {title_id})의 시트를 찾을 수 없습니다. 시트 이름을 서로 바꾸었거나 "
                "시트를 지웠다면 Ctrl+Z 로 되돌린 뒤 다시 저장하세요."
            ))
        return SectionSheet(by_name)

    @property
    def section_map(self) -> list:
        if not self.meta or not self.meta.get("section_map"):
            return []
        try:
            return json.loads(self.meta["section_map"])
        except (TypeError, ValueError) as exc:
            raise InputError(
                f"오류: _dsdx 시트의 section_map 을 읽을 수 없습니다: {exc}"
            ) from exc

    @property
    def split_notes(self) -> bool:
        """`_dsdx.split_notes` — 이 통합문서가 주석 번호마다 시트를 나눠 만들어졌는가 (1.2 WP5).

        import 는 이 값으로 원본에서 같은 섹션 구성을 다시 만든다. 값이 없으면(1.1 이하,
        또는 옵션 없이 만든 통합문서) `False` 다.
        """
        v = self.meta.get("split_notes") if self.meta else None
        return str(v).strip().lower() in ("1", "true", "y", "yes")

    @property
    def table_layouts(self) -> dict:
        """`adelim` 배치인 표만 `{표 요소 ID: section_map 의 tables[] 항목}` 로 (1.2 WP3).

        `seq` 배치(그리고 배치 기록이 없는 1.1 이하 통합문서)는 담지 않는다 — 그 표의 행은
        `cols` 가 비어 항등 배치가 된다.
        """
        out = {}
        for entry in self.section_map:
            for t in entry.get("tables") or []:
                if isinstance(t, dict) and t.get("layout") == "adelim" and t.get("id"):
                    out[t["id"]] = t
        return out


def anchor_table_row(rows: list, i: int):
    """`rows[i]` (삽입 행) 의 앵커 — 바로 위(없으면 아래)의 원본 표 행. 표 밖이면 `None`.

    삽입 행은 앵커의 `TR` 을 바이트 복제하므로 셀 수도 열 배치도 앵커를 따른다
    (`dsd_bridge.rowedit` 의 앵커 결정 규칙과 같은 범위: 같은 표 안에서만 찾는다).
    """
    for j in range(i - 1, -1, -1):
        r = rows[j]
        if r.base_kind in ("TR", "THEAD-TR") and r.seq is not None:
            return r
        if r.base_kind == "TABLE-START":
            break
    for j in range(i + 1, len(rows)):
        r = rows[j]
        if r.base_kind in ("TR", "THEAD-TR") and r.seq is not None:
            return r
        if r.base_kind == "TABLE-END":
            break
    return None


def _attach_layout(state: "WorkbookState") -> None:
    """`section_map` 의 열 배치를 행마다 `SheetRow.cols` 로 붙인다 (1.2 WP3).

    원본 표 행·복제 삽입 행은 자기 `_TABLE`/`_ID` 로, 빈 삽입 행은 앵커 행에서 배치를 얻는다.
    셀 수와 배치 길이가 다르면(숨김 열 손상 등) 붙이지 않는다 — 게이트가 판정하도록 둔다.
    """
    if state.meta is None:
        return
    layouts = state.table_layouts
    if not layouts:
        return
    for sheet in state.sheets:
        rows = sheet.rows
        for r in rows:
            if r.is_header or r.is_guide or not r.id or not r.cells:
                continue
            entry = layouts.get(r.table)
            if entry is None:
                continue
            cols = (entry.get("rowcols") or {}).get(r.id) or entry.get("cols")
            if cols and len(cols) == len(r.cell_ids):
                r.cols = [int(c) for c in cols]
        for i, r in enumerate(rows):
            if r.seq is not None or r.is_header or r.is_guide or r.cols:
                continue
            anchor = anchor_table_row(rows, i)
            if anchor is not None and anchor.cols:
                r.cols = list(anchor.cols)


def _parse_seq(v):
    """E 열 값 → int | None. 정수로 볼 수 없으면 (None, 원값) 로 남겨 게이트가 진단한다."""
    if _is_blank(v):
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def load_workbook_state(path) -> WorkbookState:
    """`.xlsx` 를 읽어 `WorkbookState` 를 만든다. 열 수 없으면 `InputError`(코드 2)."""
    from openpyxl import load_workbook

    p = Path(path)
    if not p.is_file():
        raise InputError(f"오류: 통합문서를 찾을 수 없습니다: {p}")
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
        # 수식 셀 위치는 data_only=False 로만 알 수 있다. data_only=True 는 수식 셀에 **마지막 계산값**
        # (캐시)을 주는데, 캐시가 없으면 None 이라 "빈 셀" 과 구분되지 않는다 → 원본 텍스트가 지워진다.
        wb_f = load_workbook(p, data_only=False, read_only=False)
    except Exception as exc:  # openpyxl 은 다양한 예외를 던진다 (BadZipFile, KeyError, InvalidFileException …)
        raise InputError(
            f"오류: Excel 통합문서(.xlsx)로 열 수 없습니다: {p}\n  사유: {type(exc).__name__}: {exc}\n"
            "  CSV 나 xls(97-2003) 로 저장했다면 되돌리기가 불가능합니다. 반드시 .xlsx 로 저장하세요."
        ) from exc

    meta = None
    if META_SHEET in wb.sheetnames:
        meta = {}
        for k, v in wb[META_SHEET].iter_rows(min_row=2, max_col=2, values_only=True):
            if k is None:
                continue
            meta[_as_str(k)] = v

    sheets = []
    order = 0
    for name in wb.sheetnames:
        if is_internal_sheet(name):
            continue
        ws = wb[name]
        formula_map: dict = {}
        hidden_formula_rows: set = set()
        for row_cells in wb_f[name].iter_rows():
            for c in row_cells:
                if c.data_type == "f":
                    if c.column < FIRST_VALUE_COL:
                        hidden_formula_rows.add(c.row)
                    else:
                        formula_map.setdefault(c.row, set()).add(c.column - FIRST_VALUE_COL)
        rows = []
        for r, tup in enumerate(ws.iter_rows(values_only=True), start=1):
            tup = tuple(tup)
            hidden = tup[: FIRST_VALUE_COL - 1] + (None,) * max(0, FIRST_VALUE_COL - 1 - len(tup))
            values = list(tup[FIRST_VALUE_COL - 1 :])
            while values and _is_blank(values[-1]):
                values.pop()
            seq_raw = hidden[4]
            rows.append(
                SheetRow(
                    row=r,
                    id=_as_str(hidden[0]).strip(),
                    kind=_as_str(hidden[1]).strip(),
                    table=_as_str(hidden[2]).strip(),
                    cells=_as_str(hidden[3]).strip(),
                    seq=_parse_seq(seq_raw),
                    seq_raw=seq_raw,
                    values=values,
                    formulas=formula_map.get(r, set()),
                    hidden_formula=r in hidden_formula_rows,
                )
            )
        # 끝의 완전히 빈 행(숨김 열도 값도 없음)은 잘라낸다 — Excel 이 남기는 빈 행
        while rows and not rows[-1].id and not rows[-1].kind and not rows[-1].table \
                and not rows[-1].cells and rows[-1].seq is None and rows[-1].seq_raw is None \
                and not rows[-1].values and not rows[-1].formulas and not rows[-1].hidden_formula:
            rows.pop()
        sheets.append(SheetState(name=name, order=order, rows=rows))
        order += 1
    state = WorkbookState(path=p, meta=meta, sheets=sheets)
    _attach_layout(state)
    return state


__all__ = [
    "META_SHEET",
    "INTERNAL_SHEET_PREFIX",
    "is_internal_sheet",
    "FIRST_VALUE_COL",
    "GUIDE_KINDS",
    "MAPPED_KINDS",
    "SheetRow",
    "SheetState",
    "SectionSheet",
    "WorkbookState",
    "anchor_table_row",
    "load_workbook_state",
]
