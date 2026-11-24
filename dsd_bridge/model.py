"""문서 모델 `DocModel` 과 export 투영 `project()` (계획 §4.3, §4.4).

`SourceIndex` 위에 **문서 순서**로 섹션 / 블록 / 표 / 행 / 셀을 세운다. 모델은 원본을
바꾸지 않으며, 모든 요소를 `ordinal` 과 `eid` 로 `SourceIndex` 에 되돌려 연결한다.

- **셀** `CELL = {TD, TE, TU, TH}`
- **블록** = 조상에 셀도 `TABLE` 도 없는 `P` | `TABLE` | `TITLE` | `COVER-TITLE`
  (샘플A 210 / 샘플B 160)
- **섹션** = `TITLE|COVER-TITLE` 중 `@ATOC="Y"` 를 만날 때마다 새로 시작 (두 샘플 9개).
  `ATOC="N"` 제목은 자기 섹션의 일반 블록 행이다.
  `split_notes=True` (1.2 WP5) 면 **주석 번호 제목**(`note_heading`)에서도 섹션이 새로 시작한다.
- **머리말** = 첫 `ATOC="Y"` 이전의 블록 (샘플A 4 / 샘플B 5) — 어느 시트에도 속하지 않는다.
- **표** 의 행 = 그 `TABLE` 아래의 `TR` (중첩 `TABLE` 안은 제외). `THEAD` 아래면
  `kind="THEAD-TR"`, 아니면 `kind="TR"`.

`project(model)` 은 통합문서에 쓸 내용을 **순수 데이터**(dict/list/str/bool/int)로 만든다.
export 는 이것을 그대로 시트에 쓰고, import 는 원본에서 이것을 **다시 계산해** 기준값으로
쓴다(`_dsdx_raw` 폐기의 전제 — 결정성은 골든 테스트가 지킨다).

**규범 1 (계획 §4.4)**: 블록 정의·`project()` 규칙·열 배치·`dsd_bridge.text` 의 분류/포맷이
바뀌면 `WORKBOOK_FORMAT_VERSION` 의 마이너를 올리고 `tests/golden_project.json` 을 갱신한다.

열 배치 (1.2 WP3) — `표 하나`가 **표 단위로** 두 방식 중 하나를 쓴다::

    "seq"     원문 `TR` 안의 셀 순서대로 F, G, H … (1.1 까지의 유일한 방식)
    "adelim"  Excel 열 = F + `ADELIM` — 본문 셀의 `ADELIM` 이 곧 **논리 열 번호**다

`adelim` 판정 조건(모두 만족해야 한다, `table_layout()`):

1. `THEAD` 아래가 아닌 **본문 행**이 하나 이상 있고, 그 행들에 셀이 하나 이상 있다.
2. 모든 본문 셀에 `ADELIM` 이 있고 0 이상의 정수이며 한 행 안에서 **서로 다르다**.
3. 모든 본문 행의 `ADELIM` 묶음이 **서로 같다** (표 하나 = 배치 하나).
4. 논리 열 번호가 `MAX_LOGICAL_COL` 이하다 (엉뚱한 `ADELIM` 이 시트를 무한정 넓히지 않도록).

머리글(`THEAD`) 행은 `ADELIM` 이 없고 `COLSPAN` 으로 폭을 나타내므로, **본문의 논리 열
목록을 COLSPAN 만큼씩 차례로 소비**해 배치한다. 예) 본문이 `0,1,2,3,4` 이고 머리글이
`과목 | 당기(COLSPAN=2) | 전기(COLSPAN=2)` 면 `과목→0`, `당기→1~2`(병합), `전기→3~4`(병합).
본문 논리 열에 구멍이 있어도(자본변동표의 `0,1,2,5,6`) 같은 규칙이 그대로 맞는다 —
`이익잉여금` 머리글은 다섯 번째 논리 열인 5 로 간다. 어느 셀도 놓이지 않는 논리 열은 빈 셀이다.

`_CELLS`(숨김 D열)는 **원문 `TR` 순서의 셀 ID 목록**으로 그대로 둔다(게이트 AC-28/29 계약).
열 ↔ 셀 대응은 `_dsdx.section_map[].tables[]` 에 적는다 — `table_layout_entry()` 참고.
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal

from .sheetname import make_sheet_name
from .text import (
    NUMERIC_CLASSES,
    NumPattern,
    canonical_out,
    classify,
    cr_to_lf,
    excel_number_format,
    merge_mixed_text,
)
from .xmlsplice import SourceIndex

#: 통합문서 레이아웃·`project()` 규칙의 버전 (AC-31 의 유일한 버전 게이트).
#: `dsd_bridge.__version__` 과 별개다 — 무관한 버그 수정 릴리스가 기존 통합문서를 무효화하지 않도록.
#: 1.0 → 1.1 (2026-09-04): classify 가 float 으로 정확히 표현되지 않는 숫자를 text 로 강등 (AC-21 정밀화),
#: 시트 이름의 제어문자 제거 (규범 1 에 따라 마이너 상승).
#: 1.1 → 1.2 (2026-09-10): 표 행 `attrs`(WP2) 와 `ADELIM` 기반 열 배치(WP3).
#: 1.2 → 1.3 (2026-11-24): 도구 이름 dsdx → dsd-bridge(DSD브릿지). 헤더 문구(HEADER_A1·HEADER_NOTE)가 바뀜.
#: 숨김 메타 시트 이름 `_dsdx`·`_dsdx_attrs` 와 메타 키 `dsdx_version` 은 형식 표식이라 그대로 둔다.
WORKBOOK_FORMAT_VERSION = "1.3"

CELL_TAGS = frozenset({"TD", "TE", "TU", "TH"})
TITLE_TAGS = frozenset({"TITLE", "COVER-TITLE"})
BLOCK_TAGS = frozenset({"P", "TABLE"}) | TITLE_TAGS

#: 숨김 열 이름 (A~E, 계획 §4.4).
HIDDEN_COLUMNS = ("_ID", "_KIND", "_TABLE", "_CELLS", "_SEQ")
#: 1행(헤더 행) A1 의 문구.
HEADER_A1 = "⚠ 편집 금지 (dsd-bridge 내부용)"
#: 헤더 행의 `_KIND`. import 는 이 값으로 헤더 행을 제외한다.
HEADER_KIND = "HEADER"
#: 1행 F1 에 보이는 안내.
HEADER_NOTE = "※ A~E열은 dsd-bridge 내부용 숨김 열입니다. 편집하지 마세요. 값은 F열부터 입력합니다."
#: 읽기 전용(혼합 콘텐츠) 표시 접미사 — `_KIND`(블록) 또는 `_CELLS` 의 셀 ID(셀)에 붙는다.
RO_SUFFIX = ":RO"

#: 행마다 투영하는 DART 속성의 열 순서 (`_dsdx_attrs` 시트의 열 순서와 같다, 1.2 WP2).
#: 값은 **전부 문자열**이며 속성이 없으면 빈 문자열이다.
ATTR_KEYS = (
    "alevel", "acode", "tags", "adelims", "autosum", "aunit", "aunitvalue", "aclass",
)

#: `adelim` 배치가 허용하는 최대 논리 열 번호 (F + 63 = BM 열). 이보다 큰 `ADELIM` 이 있으면
#: 그 표는 `seq` 배치로 둔다 — 잘못된 값 하나가 시트를 수천 열로 넓히지 않도록.
MAX_LOGICAL_COL = 63

#: 표 배치 이름.
LAYOUT_SEQ = "seq"
LAYOUT_ADELIM = "adelim"

#: 주석 번호 제목 (1.2 WP5, `--split-notes`) — 블록의 **첫 줄**(공백을 하나로 줄인)이
#: `21. 우발부채` 처럼 `번호.` + 공백 + 글자여야 한다. 두 샘플 실측으로 정한 규칙이다:
#:
#: - 마침표만 받는다. 주석의 하위 항목은 `(1) 무형자산의 평가` · `1) 당기말 …` 꼴이라
#:   `)` 나 `-` 까지 받으면 하위 항목에서도 시트가 갈린다 (샘플A 주석에서 2건 오탐).
#: - 마침표 뒤 **공백을 요구**한다. `1.5배 증가하였습니다` 같은 소수로 시작하는 문단을
#:   제목으로 오인하지 않는다.
#: - 번호는 섹션 안에서 **커져야** 한다(`1, 2, … 15`). 본문 중간에서 번호가 되돌아가면
#:   주석 번호가 아니라 문단 안의 열거이므로 시트를 나누지 않는다.
NOTE_HEADING_RE = re.compile(r"^(\d{1,3})\.\s+(\S.*)$")


def note_heading(text: str):
    """주석 번호 제목이면 `(번호, 제목 나머지)`, 아니면 `None` (1.2 WP5).

    판정은 **첫 비어 있지 않은 줄**로만 한다 — 두 샘플의 주석 제목은 `TITLE` 이 아니라
    제목 줄과 본문이 한 `P` 에 `&cr;` 로 이어 붙은 문단이기 때문이다
    (`1. 회사의 개요&cr;&cr;주식회사 …`).
    """
    for line in cr_to_lf(text).splitlines():
        s = " ".join(line.split())
        if not s:
            continue
        m = NOTE_HEADING_RE.match(s)
        return (int(m.group(1)), m.group(2)) if m else None
    return None


# ─────────────────────────────────────────────────────────────────────────────
# 데이터 클래스
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Cell:
    """표 안의 셀 하나 (`TD`/`TE`/`TU`/`TH`)."""

    ordinal: int
    eid: str
    tag: str
    col: int                      # TR 내 0-기반 순서
    row_index: int                # 표 안에서의 행 인덱스 (Table.rows 기준)
    text: str                     # 언이스케이프된 원문 (`&cr;` 4글자 유지). 혼합이면 순수 연접
    mixed: bool
    cls: str
    value: Decimal | None = None
    pattern: NumPattern | None = None
    colspan: int = 1
    rowspan: int = 1
    line: int = 0
    acode: str = ""               # 표준계정코드 (`ACODE`)
    alevel: str = ""              # 계정 계층 깊이 (`ALEVEL`)
    adelim: str = ""              # 논리 열 번호 (`ADELIM`)
    autosum: str = ""             # 자동합계 표시 (`AUTOSUM`)

    @property
    def numeric(self) -> bool:
        return self.cls in NUMERIC_CLASSES

    @property
    def display(self) -> str:
        """Excel 셀에 보이는 문자열 (`&cr;` → `\\n`)."""
        return cr_to_lf(self.text)

    @property
    def excel_out(self) -> str:
        """변경 판정용 표준형 문자열 (§4.5)."""
        return canonical_out(self.value) if self.numeric else canonical_out(self.display)


@dataclass
class Row:
    """표의 행 하나 (`TR`)."""

    ordinal: int
    eid: str
    kind: str                     # "TR" | "THEAD-TR"
    index: int                    # Table.rows 안의 인덱스
    cells: list = field(default_factory=list)
    parent_tag: str = ""          # THEAD / TBODY / TABLE …
    line: int = 0
    acode: str = ""               # `TR` 자신의 `ACODE` (두 샘플은 셀에만 있다)
    alevel: str = ""              # `TR` 자신의 `ALEVEL`
    autosum: str = ""             # `TR` 자신의 `AUTOSUM`

    @property
    def has_mixed(self) -> bool:
        return any(c.mixed for c in self.cells)


@dataclass
class Table:
    """블록 `TABLE` 하나."""

    ordinal: int
    eid: str
    rows: list = field(default_factory=list)
    aclass: str = ""
    line: int = 0
    #: 이 표 안의 첫 `TU` 가 가진 단위 (`AUNIT`/`AUNITVALUE`). 표 밖의 `TU` 는 물려받지 않는다 —
    #: 두 샘플에서 `AUNIT="WON"` 은 본표(`ACLASS="FINANCE"`) 그 표의 머리 행에 있다.
    aunit: str = ""
    aunitvalue: str = ""

    @property
    def extraction(self) -> bool:
        """`ACLASS="EXTRACTION"` — DART 추출 대상 표 (AC-27)."""
        return self.aclass == "EXTRACTION"

    @property
    def ncols(self) -> int:
        """행 중 최대 셀 수 (`_dsdx.section_map[].tables[].ncols`)."""
        return max((len(r.cells) for r in self.rows), default=0)

    @property
    def has_rowspan(self) -> bool:
        return any(c.rowspan > 1 for r in self.rows for c in r.cells)

    @property
    def thead_rows(self) -> int:
        return sum(1 for r in self.rows if r.kind == "THEAD-TR")


@dataclass
class Block:
    """블록 하나 — `TITLE`(=`TITLE`/`COVER-TITLE`) | `P` | `TABLE`."""

    ordinal: int
    eid: str
    kind: str                     # "TITLE" | "P" | "TABLE"
    tag: str                      # 실제 XML 태그 (TITLE / COVER-TITLE / P / TABLE)
    text: str = ""                # TITLE/P 의 언이스케이프 원문 (혼합이면 순수 연접)
    mixed: bool = False
    table: Table | None = None
    atoc: str | None = None       # TITLE 의 ATOC 속성값
    line: int = 0

    @property
    def is_sheet_title(self) -> bool:
        return self.kind == "TITLE" and self.atoc == "Y"

    @property
    def display(self) -> str:
        return cr_to_lf(self.text)

    @property
    def excel_out(self) -> str:
        return canonical_out(self.display)

    @property
    def row_count(self) -> int:
        """이 블록이 시트에서 차지하는 **매핑 행** 수 (안내 행 제외)."""
        return len(self.table.rows) if self.table is not None else 1


@dataclass
class Section:
    """시트 하나 = `ATOC="Y"` 제목부터 다음 섹션 제목 직전까지.

    `--split-notes` (1.2 WP5) 면 주석 번호 제목(`note_heading`)도 섹션을 시작하므로
    `title` 이 `ATOC="Y"` 제목이 아니라 그 번호 제목(대개 `P`)일 수 있다. 어느 쪽이든
    **섹션의 첫 블록 = 그 섹션의 제목**이며 삭제할 수 없다(코드 5).
    """

    index: int                    # 0-기반
    title: Block
    sheet_name: str
    blocks: list = field(default_factory=list)   # 제목 블록 포함, 문서 순서

    @property
    def row_count(self) -> int:
        """매핑 행 수 = TITLE 1 + P 1 + 표마다 TR 수 (AC-23 의 분모)."""
        return sum(b.row_count for b in self.blocks)

    @property
    def tables(self) -> list:
        return [b.table for b in self.blocks if b.table is not None]


@dataclass
class DocModel:
    """문서 전체."""

    index: SourceIndex
    preamble: list = field(default_factory=list)     # 머리말 블록
    sections: list = field(default_factory=list)
    blocks: list = field(default_factory=list)       # 모든 블록 (머리말 포함), 문서 순서
    #: 통합문서에 투영되지 않는 것의 개수 — `{"nested_tables": n, "preamble_blocks": n}`.
    #: 중첩 `TABLE`(다른 표나 셀 안의 표)과 첫 `ATOC="Y"` 이전 머리말 블록은 어느 시트에도
    #: 나타나지 않는다. 되돌릴 때 그 바이트는 손대지 않으므로 보존되지만, 편집은 불가능하다.
    skipped: dict = field(default_factory=dict)
    #: 주석 번호 제목마다 섹션(시트)을 나눴는가 (1.2 WP5). export 가 `_dsdx.split_notes` 로
    #: 적고 import 가 그 값으로 같은 모델을 다시 만든다.
    split_notes: bool = False

    # 구축 ────────────────────────────────────────────────────────────────
    @classmethod
    def build(cls, index: SourceIndex, raw: bytes | None = None, *,
              split_notes: bool = False) -> "DocModel":
        raw = index.raw if raw is None else raw
        spans = index.spans

        def is_block(sp) -> bool:
            if sp.tag not in BLOCK_TAGS:
                return False
            po = sp.parent_ordinal
            while po >= 0:
                anc = spans[po]
                if anc.tag in CELL_TAGS or anc.tag == "TABLE":
                    return False
                po = anc.parent_ordinal
            return True

        def content_text(sp) -> str:
            return merge_mixed_text(index, raw, sp.ordinal) if sp.child_ordinals else sp.text

        def build_table(sp) -> Table:
            table = Table(
                ordinal=sp.ordinal,
                eid=index.eid_of(sp.ordinal),
                aclass=sp.attrs.get("ACLASS", ""),
                line=sp.line,
            )

            def walk(o: int) -> None:
                for c in spans[o].child_ordinals:
                    ch = spans[c]
                    if ch.tag == "TABLE":      # 중첩 표(두 샘플 0건)는 이 표의 행이 아니다
                        continue
                    if ch.tag == "TR":
                        parent_tag = spans[ch.parent_ordinal].tag
                        row = Row(
                            ordinal=ch.ordinal,
                            eid=index.eid_of(ch.ordinal),
                            kind="THEAD-TR" if parent_tag == "THEAD" else "TR",
                            index=len(table.rows),
                            parent_tag=parent_tag,
                            line=ch.line,
                            acode=_str_attr(ch.attrs, "ACODE"),
                            alevel=_str_attr(ch.attrs, "ALEVEL"),
                            autosum=_str_attr(ch.attrs, "AUTOSUM"),
                        )
                        col = 0
                        for cc in ch.child_ordinals:
                            cs = spans[cc]
                            if cs.tag not in CELL_TAGS:
                                continue
                            mixed = bool(cs.child_ordinals)
                            text = content_text(cs)
                            k = classify(text, mixed=mixed)
                            row.cells.append(
                                Cell(
                                    ordinal=cs.ordinal,
                                    eid=index.eid_of(cs.ordinal),
                                    tag=cs.tag,
                                    col=col,
                                    row_index=row.index,
                                    text=text,
                                    mixed=mixed,
                                    cls=k.cls,
                                    value=k.value,
                                    pattern=k.pattern,
                                    colspan=_int_attr(cs.attrs, "COLSPAN"),
                                    rowspan=_int_attr(cs.attrs, "ROWSPAN"),
                                    line=cs.line,
                                    acode=_str_attr(cs.attrs, "ACODE"),
                                    alevel=_str_attr(cs.attrs, "ALEVEL"),
                                    adelim=_str_attr(cs.attrs, "ADELIM"),
                                    autosum=_str_attr(cs.attrs, "AUTOSUM"),
                                )
                            )
                            if cs.tag == "TU" and not table.aunit and _str_attr(cs.attrs, "AUNIT"):
                                table.aunit = _str_attr(cs.attrs, "AUNIT")
                                table.aunitvalue = _str_attr(cs.attrs, "AUNITVALUE")
                            col += 1
                        table.rows.append(row)
                        continue                # TR 안에 다시 TR 은 없다
                    walk(c)

            walk(sp.ordinal)
            return table

        model = cls(index=index, split_notes=split_notes)
        nested_tables = 0
        parent_sheet = ""          # 지금 섹션이 속한 `ATOC="Y"` 시트 이름 (주석 시트 이름의 앞머리)
        note_no = 0                # 그 안에서 마지막으로 인정한 주석 번호
        for sp in spans:
            if not is_block(sp):
                if sp.tag == "TABLE":
                    nested_tables += 1      # 다른 표·셀 안의 표 — 어느 시트에도 투영되지 않는다
                continue
            if sp.tag == "TABLE":
                block = Block(
                    ordinal=sp.ordinal,
                    eid=index.eid_of(sp.ordinal),
                    kind="TABLE",
                    tag=sp.tag,
                    table=build_table(sp),
                    line=sp.line,
                )
            else:
                block = Block(
                    ordinal=sp.ordinal,
                    eid=index.eid_of(sp.ordinal),
                    kind="TITLE" if sp.tag in TITLE_TAGS else "P",
                    tag=sp.tag,
                    text=content_text(sp),
                    mixed=bool(sp.child_ordinals),
                    atoc=sp.attrs.get("ATOC") if sp.tag in TITLE_TAGS else None,
                    line=sp.line,
                )
            model.blocks.append(block)

            heading = None
            if (split_notes and model.sections and not block.is_sheet_title
                    and block.kind in ("TITLE", "P")):
                h = note_heading(block.text)
                if h is not None and h[0] > note_no:
                    heading = h

            if block.is_sheet_title or heading is not None:
                used = [s.sheet_name for s in model.sections]
                if heading is None:
                    note_no, raw_name = 0, block.text
                else:
                    note_no = heading[0]
                    raw_name = f"{parent_sheet} {heading[0]} {heading[1]}"
                section = Section(
                    index=len(model.sections),
                    title=block,
                    sheet_name=make_sheet_name(raw_name, used),
                    blocks=[block],
                )
                model.sections.append(section)
                if heading is None:
                    parent_sheet = section.sheet_name
            elif model.sections:
                model.sections[-1].blocks.append(block)
            else:
                model.preamble.append(block)
        model.skipped = {
            "nested_tables": nested_tables,
            "preamble_blocks": len(model.preamble),
        }
        return model

    # 조회 ────────────────────────────────────────────────────────────────
    @property
    def sheet_blocks(self) -> list:
        """시트에 들어가는 블록 (머리말 제외)."""
        return [b for s in self.sections for b in s.blocks]

    def block_counts(self) -> dict:
        out = {"TITLE": 0, "P": 0, "TABLE": 0}
        for b in self.sheet_blocks:
            out[b.kind] += 1
        return out

    def mixed_ordinals(self) -> list:
        """혼합 콘텐츠 요소의 순번 — 블록 `P`/`TITLE` + 표 안의 셀 (문서 순서)."""
        out = []
        for b in self.blocks:
            if b.table is not None:
                out.extend(c.ordinal for r in b.table.rows for c in r.cells if c.mixed)
            elif b.mixed:
                out.append(b.ordinal)
        return sorted(out)

    def find_cell(self, ordinal: int):
        """순번으로 (Table, Row, Cell) 을 찾는다. 없으면 None."""
        for b in self.blocks:
            if b.table is None:
                continue
            for r in b.table.rows:
                for c in r.cells:
                    if c.ordinal == ordinal:
                        return b.table, r, c
        return None

    def find_cell_by_line(self, line: int):
        """1-기반 원본 행 번호로 (Table, Row, Cell) 을 찾는다 (테스트·`-v` 용)."""
        for b in self.blocks:
            if b.table is None:
                continue
            for r in b.table.rows:
                for c in r.cells:
                    if c.line == line:
                        return b.table, r, c
        return None


def _int_attr(attrs: dict, name: str, default: int = 1) -> int:
    try:
        return int(attrs.get(name, default))
    except (TypeError, ValueError):
        return default


def _str_attr(attrs: dict, name: str) -> str:
    """속성값을 문자열로. 없으면 빈 문자열 (`None` 과 빈칸을 구분하지 않는다)."""
    v = attrs.get(name)
    return "" if v is None else str(v)


def row_attrs(table: Table, row: Row) -> dict:
    """표 행 하나의 DART 속성 (`ATTR_KEYS` 순서, 값은 전부 문자열) — 1.2 WP2.

    - `alevel`/`acode`/`autosum` 은 `TR` 에 있으면 그것을, 없으면 **셀 중 처음으로 값을 가진 것**을
      쓴다 (두 샘플은 `ADELIM="0"` 인 첫 `TE` 에 들어 있다).
    - `tags` 는 셀 태그를 원문 순서로 (`TE|TE|TD`), `adelims` 는 셀마다의 `ADELIM` 을 같은 순서로
      (`0|1|2`, 없는 셀은 빈 칸) 이어 붙인다. 두 목록의 길이는 `_CELLS` 와 같다.
    - `aunit`/`aunitvalue` 는 **표 단위**(그 표 안의 첫 `TU`), `aclass` 는 표의 `ACLASS` 다.
    """
    return {
        "alevel": row.alevel or _first_nonempty(c.alevel for c in row.cells),
        "acode": row.acode or _first_nonempty(c.acode for c in row.cells),
        "tags": "|".join(c.tag for c in row.cells),
        "adelims": "|".join(c.adelim for c in row.cells),
        "autosum": row.autosum or _first_nonempty(c.autosum for c in row.cells),
        "aunit": table.aunit,
        "aunitvalue": table.aunitvalue,
        "aclass": table.aclass,
    }


def _first_nonempty(values) -> str:
    for v in values:
        if v:
            return v
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# 열 배치 (1.2 WP3)
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class TableLayout:
    """표 하나의 열 배치. `kind` 가 `"seq"` 면 나머지는 비어 있다.

    - `cols` — **본문 행 공통** 배치. `cols[i]` 는 `_CELLS` 의 i 번째 셀이 놓이는 0-기반 열
      오프셋(F=0)이며, `adelim` 배치에서는 곧 그 셀의 `ADELIM` 값이다.
    - `rowcols` — 본문 배치와 다른 행의 배치 `{행 eid: [열 오프셋, …]}`. 머리글(`THEAD`) 행만
      여기 들어간다. 행 하나가 `_CELLS` 순서로 어느 열에 놓이는지는 `cols_for(eid)` 로 얻는다.
    - `spans` — 병합 폭 `{행 eid: [폭, …]}`. `1` 뿐인 행은 담지 않는다(export 전용 정보).
    """

    kind: str = LAYOUT_SEQ
    cols: tuple = ()
    rowcols: dict = field(default_factory=dict)
    spans: dict = field(default_factory=dict)

    @property
    def is_adelim(self) -> bool:
        return self.kind == LAYOUT_ADELIM

    def cols_for(self, row_eid: str) -> list:
        """행 하나의 열 배치 (`_CELLS` 순서). `seq` 배치면 빈 목록 = 항등."""
        if not self.is_adelim:
            return []
        return list(self.rowcols.get(row_eid) or self.cols)


def _body_cols(row: Row):
    """본문 행의 `ADELIM` 목록. 조건을 어기면 `None`."""
    if not row.cells:
        return None
    out = []
    for c in row.cells:
        s = c.adelim.strip()
        if not s.isdigit():
            return None
        v = int(s)
        if v > MAX_LOGICAL_COL:
            return None
        out.append(v)
    return out if len(set(out)) == len(out) else None


def _header_placement(row: Row, body_cols: list):
    """머리글 행의 (열 오프셋, 병합 폭) — 본문 논리 열을 `COLSPAN` 만큼씩 소비한다."""
    cols, spans = [], []
    n = len(body_cols)
    p = 0
    for c in row.cells:
        if p < n:
            take = min(max(1, c.colspan), n - p)
            start, end = body_cols[p], body_cols[p + take - 1]
            cols.append(start)
            spans.append(end - start + 1)
            p += take
        else:
            # 본문 논리 열을 다 쓴 뒤의 여분 머리글 셀 — 마지막 논리 열 뒤에 차례로 놓는다.
            cols.append(body_cols[-1] + 1 + (p - n))
            spans.append(1)
            p += 1
    return cols, spans


def table_layout(table: Table) -> TableLayout:
    """표 하나의 열 배치를 정한다 (모듈 docstring 의 `adelim` 판정 조건)."""
    body = [r for r in table.rows if r.kind != "THEAD-TR"]
    if not body:
        return TableLayout()
    ref = None
    for r in body:
        cols = _body_cols(r)
        if cols is None:
            return TableLayout()
        if ref is None:
            ref = cols
        elif ref != cols:
            return TableLayout()
    rowcols, spans = {}, {}
    for r in table.rows:
        if r.kind != "THEAD-TR":
            continue
        cols, sp = _header_placement(r, ref)
        rowcols[r.eid] = cols
        if any(s > 1 for s in sp):
            spans[r.eid] = sp
    return TableLayout(LAYOUT_ADELIM, tuple(ref), rowcols, spans)


def table_layout_entry(table: Table, layout: TableLayout) -> dict:
    """`_dsdx.section_map[].tables[]` 항목 하나.

    `seq` 면 `{"id", "ncols", "layout": "seq"}` 뿐이고, `adelim` 이면 `cols`(본문 행 공통
    배치)와, 그와 다른 행만 담은 `rowcols`(머리글 행)가 붙는다. 배치를 **표 하나당 한 줄 +
    머리글 행 몇 개**로 적으므로 행 수가 아니라 표 수에만 비례한다 — 두 샘플의 `section_map`
    은 3.3 KB → 4.9 KB 로, 대부분은 표 86/43개마다 붙는 `layout` 값이다(Excel 한 셀 상한
    32,767자). import 는 이 항목만 보고 `열 → 셀 ID` 를 되짚는다 (`dsd_bridge.sheetstate`).
    """
    entry = {"id": table.eid, "ncols": table.ncols, "layout": layout.kind}
    if layout.is_adelim:
        entry["cols"] = list(layout.cols)
        if layout.rowcols:
            entry["rowcols"] = {k: list(v) for k, v in sorted(layout.rowcols.items())}
    return entry


# ─────────────────────────────────────────────────────────────────────────────
# export 투영 (계획 §4.4)
# ─────────────────────────────────────────────────────────────────────────────


def _cell_value(c: Cell) -> dict:
    v = {
        "id": c.eid,
        "cls": c.cls,
        "raw": c.text,
        "out": c.excel_out,
        "ro": c.mixed,
    }
    if c.numeric:
        v["t"] = "n"
        v["num"] = str(c.value)
        v["pat"] = c.pattern.as_dict()
        v["fmt"] = excel_number_format(c.pattern)
    else:
        v["t"] = "s"
    return v


def _block_value(b: Block) -> dict:
    return {"id": b.eid, "cls": "mixed" if b.mixed else "text", "raw": b.text,
            "out": b.excel_out, "ro": b.mixed, "t": "s"}


def project(model: DocModel) -> dict:
    """모델을 통합문서에 쓸 순수 데이터로 투영한다.

    반환값::

        {"workbook_format_version": "1.2",
         "sheets": [{"name", "title_id", "rows": [row, …]}, …],
         "section_map": [{"sheet", "title_id", "row_count",
                          "tables": [{"id", "ncols", "layout", "cols"?, "rowcols"?}, …]}, …],
         "skipped": {"nested_tables": n, "preamble_blocks": n}}

    행(row) 하나는 숨김 A~E열과 표시 값의 목록이다::

        {"id": _ID, "kind": _KIND, "table": _TABLE, "cells": _CELLS, "seq": _SEQ,
         "values": [value, …],                # `_CELLS` 와 같은 순서 (원문 셀 순서)
         "cols": [열 오프셋, …],              # `adelim` 배치의 표 행에만 (F=0)
         "spans": [병합 폭, …],               # `adelim` 배치에서 폭이 1 이 아닌 행에만
         "attrs": {…ATTR_KEYS…},              # 표 행(TR/THEAD-TR)에만 — `_dsdx_attrs` 시트의 자료
         "guide": {"extraction": bool, "rowspan": bool}}   # TABLE-START 행에만

    - 1행은 헤더 행: `id=HEADER_A1`, `kind=HEADER`, `seq=1`, `values=[HEADER_NOTE]`.
    - `_SEQ` 는 시트마다 1..N 을 빠짐없이 부여한다 (안내 행·헤더 행 포함).
    - 안내 행 규약: `TABLE-START` 의 `_ID` = 표 ID, `TABLE-END` 의 `_ID` = 빈 문자열,
      둘 다 `_TABLE` 보유.
    - 혼합 콘텐츠: 블록 행은 `_KIND` 에 `:RO`, 셀은 `_CELLS` 의 셀 ID 에 `:RO`.
    - `_KIND` 의 `TITLE` 은 XML 의 `TITLE` 과 `COVER-TITLE` 둘 다를 가리킨다.
    - `cols` 가 없으면 값은 F 열부터 순서대로 놓인다(`seq` 배치).

    같은 원본에 대해 언제나 같은 결과를 준다(결정성). 시간·경로 등 원본 밖의 정보는
    넣지 않는다 — 그런 값은 export 가 `_dsdx` 시트에 따로 쓴다.
    """
    sheets = []
    section_map = []
    for section in model.sections:
        rows = [
            {"id": HEADER_A1, "kind": HEADER_KIND, "table": "", "cells": "", "seq": 1,
             "values": [{"id": "", "cls": "text", "raw": HEADER_NOTE, "out": HEADER_NOTE,
                         "ro": True, "t": "s"}]}
        ]
        seq = 1
        layouts = {}
        for b in section.blocks:
            if b.table is None:
                seq += 1
                rows.append({
                    "id": b.eid,
                    "kind": b.kind + (RO_SUFFIX if b.mixed else ""),
                    "table": "",
                    "cells": "",
                    "seq": seq,
                    "values": [_block_value(b)],
                })
                continue
            t = b.table
            layout = table_layout(t)
            layouts[t.eid] = layout
            seq += 1
            rows.append({
                "id": t.eid, "kind": "TABLE-START", "table": t.eid, "cells": "", "seq": seq,
                "values": [],
                "guide": {"extraction": t.extraction, "rowspan": t.has_rowspan},
            })
            for r in t.rows:
                seq += 1
                row = {
                    "id": r.eid,
                    "kind": r.kind,
                    "table": t.eid,
                    "cells": "|".join(c.eid + (RO_SUFFIX if c.mixed else "") for c in r.cells),
                    "seq": seq,
                    "values": [_cell_value(c) for c in r.cells],
                    "attrs": row_attrs(t, r),
                }
                if layout.is_adelim:
                    row["cols"] = layout.cols_for(r.eid)
                    spans = layout.spans.get(r.eid)
                    if spans:
                        row["spans"] = list(spans)
                rows.append(row)
            seq += 1
            rows.append({
                "id": "", "kind": "TABLE-END", "table": t.eid, "cells": "", "seq": seq,
                "values": [],
            })
        sheets.append({"name": section.sheet_name, "title_id": section.title.eid, "rows": rows})
        section_map.append({
            "sheet": section.sheet_name,
            "title_id": section.title.eid,
            "row_count": section.row_count,
            "tables": [table_layout_entry(t, layouts[t.eid]) for t in section.tables],
        })
    return {
        "workbook_format_version": WORKBOOK_FORMAT_VERSION,
        "sheets": sheets,
        "section_map": section_map,
        "skipped": dict(model.skipped),
    }


#: `project_json` 이 빼는 키 — 통합문서에 쓰는 **내용**이 아니라 투영되지 않은 것에 대한
#: 진단 수치다. 규범 1 이 골든 갱신을 요구하는 것은 분류·포맷·블록 정의·열 배치이므로
#: 이 수치는 골든 해시에 넣지 않는다 (경고 문구가 바뀌어도 기존 통합문서는 유효하다).
NON_CONTENT_KEYS = ("skipped",)


def project_json(proj: dict) -> str:
    """`project()` 결과의 **정규 JSON** — 키 정렬·공백 없음. 골든 해시와 diff 의 기준."""
    body = {k: v for k, v in proj.items() if k not in NON_CONTENT_KEYS}
    return json.dumps(body, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def project_digest(proj: dict) -> str:
    """`project()` 결과의 SHA-256 (골든 회귀 테스트, 계획 §4.4 규범 1)."""
    return hashlib.sha256(project_json(proj).encode("utf-8")).hexdigest()


__all__ = [
    "WORKBOOK_FORMAT_VERSION",
    "CELL_TAGS",
    "TITLE_TAGS",
    "BLOCK_TAGS",
    "HIDDEN_COLUMNS",
    "HEADER_A1",
    "HEADER_KIND",
    "HEADER_NOTE",
    "RO_SUFFIX",
    "ATTR_KEYS",
    "MAX_LOGICAL_COL",
    "LAYOUT_SEQ",
    "LAYOUT_ADELIM",
    "NOTE_HEADING_RE",
    "note_heading",
    "NON_CONTENT_KEYS",
    "Cell",
    "Row",
    "Table",
    "Block",
    "Section",
    "DocModel",
    "TableLayout",
    "row_attrs",
    "table_layout",
    "table_layout_entry",
    "project",
    "project_json",
    "project_digest",
]
