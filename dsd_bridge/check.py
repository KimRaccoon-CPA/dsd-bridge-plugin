"""합계검증 `dsd-bridge check` (계획 1.2 WP4).

DART 편집기의 **재무제표 검사(F9)** 가 잡아내는 "계산 오류"(합계 불일치)를 `.dsd` 또는
내보낸 `.xlsx` 에서 미리 찾아 준다. 이름(대손충당금·감가상각누계액 …)이나 테두리 같은
겉모습이 아니라 DART 가 실제로 문서에 적어 둔 속성 **`ALEVEL`(계정 계층) / `ADELIM`(논리 열)
/ `ACODE`(표준계정코드) / `ACLASS`** 만 본다. 이 모듈은 **읽기 전용**이다 — 모델·export·import
는 건드리지 않는다(규범 2).

검사 대상 표
------------
1. 본문 행에 `ACODE` **또는** `ALEVEL` 이 있는 표 (재무상태표·손익계산서·현금흐름표 등
   표준계정 표). 두 샘플 실측: 샘플A 는 `ACODE`+`ALEVEL`, 샘플B 은 `ALEVEL` 만 있다.
2. `ACLASS="EXTRACTION"` 표 (DART 추출 대상). 이 표들은 대개 `ALEVEL` 이 없어 검사 항목이
   생기지 않지만, 있으면 같은 규칙으로 검사한다.

논리 열이 `0~4` 가 아닌 표는 표준 재무제표가 아니다. 표준 재무제표는 `0`=과목, `1·2`=당기
(세부·합계), `3·4`=전기(세부·합계) 다. 그런 표 중 **자본변동표**는 아래 "자본변동표" 규칙으로
따로 검사하고, 그 밖의 표는 제외한다.

금액 읽기
---------
행 하나의 **기간별 금액**은 논리 열 쌍에서 고른다 — 당기 `(1,2)`, 전기 `(3,4)`. 합계 열(짝수)에
값이 있으면 그것을, 없으면 세부 열(홀수)을 쓴다. 괄호는 음수, `-` 는 0, 빈칸·공백은 **값 없음**
이다. 값은 `Decimal` 로만 다뤄 원 단위 정수를 그대로 보존한다. `ADELIM` 이 없는 표는 셀 순서를
논리 열로 본다(첫 셀=과목=0, 그다음부터 1,2,3…).

한 기간에서 **두 열에 값이 다 있으면** 그 행이 다른 기간에서 쓴 열 종류를 따른다. 행의 열 배치는
기간이 달라도 같기 때문이다(당기 `1·2` 와 전기 `3·4` 는 같은 모양이다). 예: 현금흐름표의 한
가산 항목이 전기 금액을 세부 열(3)에 적고 합계 열(4)에 `-` 를 남겨 둔 경우, 합계 열을 그대로
믿으면 `0` 으로 읽혀 상위 영업활동 현금흐름까지 어긋난다.

검사 세 가지
------------
- **세로 합계** (`구조=하단합계` / `상단합계`): `ALEVEL` 트리에서 부모 행의 금액과 자식 행들의
  합을 견준다. 자식이 부모 **바로 아래**에 오면 `하단합계`, 부모 **바로 위**에 오면 `상단합계`
  다(둘 다 지원). 자식은 레벨이 정확히 부모+1 인 행들이며, 레벨이 부모 이하인 행을 만나면
  거기서 끊는다. 차감 항목은 DART 원문이 이미 괄호 음수라 그냥 더하면 된다.
- **총계 등식** (`구조=총계등식`): 공백을 지운 과목명이 `총계|합계|계` 로 끝나는 레벨 0 행은
  직전 총계(또는 표 시작) 이후의 레벨 0 행들과 견준다. 이때 각 구성원의 값은 **그 행의 말단
  하위 항목 합**이다 — DART 는 결손금처럼 차감 성격의 중간 합계를 **절댓값**으로 적기 때문이다
  (샘플A 재무상태표 `Ⅲ.결손금 1,234,567,890` ↔ 하위 `미처리결손금 (1,234,567,890)`).
  하위 항목이 없는 행은 자기 금액을 쓴다.
- **대차평형** (`구조=대차평형`): 재무상태표의 `자산총계 = 부채와자본총계`, 그 행이 없으면
  `자산총계 = 부채총계 + 자본총계`. 짝이 되는 행은 **과목명**(공백 제거)이 그 총계 이름과 정확히
  같은 레벨 0 행뿐이고, 없으면 대차평형은 아예 하지 않는다 — 다른 행으로 대신하지 않는다.
  손익계산서의 당기순이익 등은 규칙이 불명확하므로 시도하지 않는다.

자본변동표 (WP6)
----------------
열이 기간이 아니라 **자본 구성요소**(자본금 · 자본잉여금 · 이익잉여금 · 자본조정 · 총계)이고
행이 `기초 + 변동 = 기말` 인 표다. 시트(제목) 이름에 `자본변동표` 가 있거나 논리 열이 `0~4` 를
벗어나면(실측 `0,1,2,5,6` · `0,1,5,6` · `0,1,3,5,6`) 자본변동표로 보고, 아래 둘을 검사한다.
`ALEVEL` 은 쓰지 않는다 — 실측에서 잔액 행·변동 행이 모두 레벨 0 인 문서가 있다.

- **가로** (`구조=ce_row`): 머리글(`THEAD`)의 공백 지운 이름이 `총계|합계|계` 인 열이 총계 열이고,
  나머지 금액 열의 합과 견준다. 총계 열 머리글이 없으면 가로는 하지 않는다.
- **세로** (`구조=ce_col`): **잔액 행** 둘 사이의 변동 행을 잔액에 더해 다음 잔액과 견준다.
  잔액 행은 과목명에 `기초`/`기말` 이 있거나 (로마숫자·번호 접두 뒤에) 날짜로 시작하는 행이다 —
  예: `20X2.06.30(당기말)` · `20X1년 1월 1일 (제5기 기초)` · `I. 20X1.01.01 (전기초)` ·
  `20X1.01.01 (수정후)`(회계정책 변경 후 재작성) · `20X0.03.15 (설립일)`.
- **이월** (`구조=ce_carry`): 잔액 행 둘 사이에 변동 행이 없으면 `전기말 = 당기초` 이월이다.
  **날짜가 하루 차이일 때만** 검사한다 — 반기보고서는 `20X1년 6월 30일(전반기말)` 다음이
  `20X2년 1월 1일(당반기초)` 라 이월이 아니다.

빈칸·`-` 는 0 으로 센다(표에 그 칸이 없다는 뜻이므로). 판정은 아래와 같고 `부호주의` 는 쓰지 않는다.

판정
----
`차이 = 입력 − 계산` 으로

- `0` → **OK**
- `0 < |차이| ≤ tolerance` (기본 2) → **단수차**
- 그 밖이지만 **하위 하나(또는 전부)의 부호를 뒤집으면** 허용 오차 안에 들어오면 → **부호주의**.
  금액은 맞고 차감 항목의 부호 표기만 다르다는 뜻이다. DART 는 차감 계정을 `ACODE` 로 알지만
  문서에는 그 표시가 없어 자동으로는 구분할 수 없다. 예:
  현금흐름표 `2. 재무활동으로 인한 현금유출액 (1,100,000)` 의 하위가 `배당금지급 100,000` ·
  `차입금의 상환 1,000,000` 로 **양수**인 경우, 또는 `유입액 − 유출액` 구조인데 두 하위가 모두
  양수인 경우.
  불일치로 세지 않되 기본 출력에 이유와 함께 남긴다.
- 그 밖 → **불일치** (종료 코드 8)

부호주의 추정은 **세로 합계**에만 쓴다. 총계 등식·대차평형은 부호를 그대로 본다.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

from .attrs import ATTRS_SHEET, read_attrs_sheet
from .errors import FOOTING_MISMATCH_EXIT, InputError, OutputError
from .text import DEFAULT_PATTERN, classify, cr_to_lf, format_number

#: 기본 허용 단수차 (원). 실무에서 흔히 쓰는 단수차 허용치.
DEFAULT_TOLERANCE = 2

#: 기간 이름 — 논리 열 쌍 `(1,2)` = 당기, `(3,4)` = 전기.
PERIOD_NAMES = ("당기", "전기")
#: 표준 재무제표 표가 쓰는 논리 열 (`0`=과목, `1~4`=금액). 이 밖의 열이 있으면 검사 제외.
STD_COLS = frozenset({0, 1, 2, 3, 4})

#: 판정.
OK = "OK"
ROUNDING = "단수차"
SIGN_NOTE = "부호주의"
MISMATCH = "불일치"
#: 요약에 세는 순서.
VERDICTS = (OK, ROUNDING, SIGN_NOTE, MISMATCH)

#: 검사 구조 이름.
STRUCT_BELOW = "하단합계"      # 내역(자식)이 부모 **아래**
STRUCT_ABOVE = "상단합계"      # 내역(자식)이 부모 **위**
STRUCT_TOTAL = "총계등식"
STRUCT_BALANCE = "대차평형"
STRUCT_CE_ROW = "ce_row"        # 자본변동표 가로 — 구성요소 합 = 총계 열
STRUCT_CE_COL = "ce_col"        # 자본변동표 세로 — 기초 + 변동 = 기말
STRUCT_CE_CARRY = "ce_carry"    # 자본변동표 이월 — 전기말 = 당기초

#: 총계 행 판정 — 공백을 지운 과목명이 이것으로 끝나면 총계 행이다.
_RE_TOTAL = re.compile(r"(총계|합계|계)$")
_RE_WS = re.compile(r"[\s　]+")

#: 대차평형에 쓰는 과목명 규칙 — 공백을 지운 이름이 이것과 **정확히** 같아야 한다.
#: `ACODE` 는 쓰지 않는다: 표준계정코드가 과목과 어긋난 문서가 있다(예: 비유동부채 행이
#: 부채와자본총계 코드를 단 경우).
BS_LABELS = (
    ("자산총계", re.compile(r"^자산총계$")),
    ("부채총계", re.compile(r"^부채총계$")),
    ("자본총계", re.compile(r"^자본총계$")),
    ("부채와자본총계", re.compile(r"^부채(및|와)?자본의?총계$")),
)

#: 자본변동표 판정 — 시트(제목) 이름에 이것이 있으면 논리 열과 상관없이 자본변동표로 본다.
CE_TITLE = "자본변동표"
#: 자본변동표 총계 열 — 공백 지운 **머리글**이 이것과 정확히 같은 논리 열.
_RE_CE_TOTAL_HEAD = re.compile(r"^(총계|합계|계)$")
#: 잔액 행 — 과목명(공백 제거)에 `기초`/`기말` 이 있거나, 로마숫자·번호 접두 뒤 날짜로 시작.
_RE_CE_BALANCE = re.compile(r"기초|기말")
_RE_CE_DATE_HEAD = re.compile(r"^(?:[IVXivxⅠ-ⅿⅰ-ⅹ\d]{1,6}[.)])?\d{4}[.\-/년]")
#: 이월 판정에 쓰는 날짜 — 과목명 어디에 있어도 **처음 나온 것**을 쓴다.
_RE_CE_DATE = re.compile(r"(\d{4})[.\-/년](\d{1,2})[.\-/월](\d{1,2})")

#: 표 하나를 검사에서 제외하는 사유 — `ALEVEL` 이 없어 계층을 모를 때.
NO_LEVEL_REASON = "ALEVEL(계정 계층)이 없어 합계 구조를 알 수 없습니다."
#: 검사한 표가 하나도 없을 때 알리는 줄 (`{extra}` 자리에 제외한 표 수가 들어간다).
NOTE_NO_TABLE = (
    "안내: 합계검증을 한 표가 없습니다 — 표준계정 속성(ACODE/ALEVEL)이 있는 재무제표 표를 "
    "찾지 못했습니다{extra}."
)


# ─────────────────────────────────────────────────────────────────────────────
# 값 읽기
# ─────────────────────────────────────────────────────────────────────────────


def parse_amount(value):
    """셀 값 하나 → `Decimal` | `None`(값 없음).

    `.dsd` 는 원문 문자열, `.xlsx` 는 openpyxl 이 준 숫자/문자열이 들어온다. 문자열은
    `dsd_bridge.text.classify` 규칙(괄호 음수·콤마·소수)으로 읽고, 앞뒤 공백은 검사에서만
    무시한다(왕복 규칙은 그대로 — 이 모듈은 읽기 전용이다). `-` 는 0, 빈칸·공백은 값 없음.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(repr(value))
    if isinstance(value, Decimal):
        return value
    s = cr_to_lf(str(value)).strip()
    if not s:
        return None
    k = classify(s)
    if k.numeric:
        return k.value
    if k.cls == "dash":
        return Decimal(0)
    return None


def _label(text: str) -> str:
    """과목명 표시용 정리.

    공백을 하나로 줄이되, `자  산  총  계` 처럼 **모든 토막이 한 글자**면 공백을 아예 없앤다
    (DART 검사 창의 `자산총계` 표기와 같아진다). `1. 현금및현금성자산(주석3,14)` 처럼 한 글자가
    아닌 토막이 있으면 공백 하나로 둔다.
    """
    parts = [p for p in _RE_WS.split(cr_to_lf(text).strip()) if p]
    if not parts:
        return ""
    return "".join(parts) if all(len(p) == 1 for p in parts) else " ".join(parts)


def _nospace(text: str) -> str:
    return _RE_WS.sub("", cr_to_lf(text))


def _level(raw: str):
    s = str(raw).strip()
    return int(s) if s.isdigit() else None


# ─────────────────────────────────────────────────────────────────────────────
# 검사 입력 모델 (`.dsd` 와 `.xlsx` 가 같은 모양으로 모인다)
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class CheckRow:
    """검사에 쓰는 표 본문 행 하나."""

    eid: str
    label: str
    level: int | None
    acode: str
    vals: dict          # 논리 열 → Decimal (값이 없는 열은 키가 없다)

    @property
    def is_total(self) -> bool:
        return self.level == 0 and bool(_RE_TOTAL.search(_nospace(self.label)))

    @property
    def column_kind(self) -> str:
        """이 행이 금액을 적는 열 종류 — `"합계"`(짝수) / `"세부"`(홀수) / `""`(알 수 없음).

        한 기간에서 두 열 중 **하나만** 차 있으면 그것이 이 행의 열이다. 두 기간의 판단이
        엇갈리거나 어느 기간도 단서를 주지 않으면 빈 문자열이다.
        """
        kinds = set()
        for period in range(len(PERIOD_NAMES)):
            detail_col, total_col = 2 * period + 1, 2 * period + 2
            has_detail, has_total = detail_col in self.vals, total_col in self.vals
            if has_detail != has_total:
                kinds.add("합계" if has_total else "세부")
        return kinds.pop() if len(kinds) == 1 else ""

    def amount(self, period: int):
        """`period` (0=당기, 1=전기) 의 금액. 합계 열(짝수) 우선, 없으면 세부 열(홀수).

        두 열에 값이 다 있으면 이 행이 다른 기간에서 쓴 열(`column_kind`)을 먼저 본다
        (모듈 docstring 참고).
        """
        detail_col, total_col = 2 * period + 1, 2 * period + 2
        order = ((detail_col, total_col) if self.column_kind == "세부"
                 else (total_col, detail_col))
        for col in order:
            if col in self.vals:
                return self.vals[col]
        return None


@dataclass(frozen=True)
class CheckTable:
    """검사 대상 표 하나."""

    sheet: str
    eid: str
    aclass: str
    rows: tuple = ()
    excluded: str = ""      # 비어 있지 않으면 검사 제외 사유 (한국어)
    #: 머리글(`THEAD`) 셀 `(논리 열, 원문)` — 자본변동표의 열 이름에만 쓴다.
    headers: tuple = ()


def _make_rows(raw_rows) -> tuple:
    """`(eid, alevel, acode, [(adelim, 값), …])` 목록 → `CheckRow` 목록.

    `ADELIM` 이 모든 셀에 있으면 그것이 논리 열이고, 하나라도 없으면 **셀 순서**를 논리 열로
    본다(첫 셀=과목=0). 표 안에서 방식이 섞이지 않도록 표 단위로 정한다.
    """
    use_adelim = all(
        str(d).strip().isdigit() for _, _, _, cells in raw_rows for d, _ in cells
    ) and any(cells for _, _, _, cells in raw_rows)
    out = []
    for eid, alevel, acode, cells in raw_rows:
        vals, label = {}, ""
        for i, (adelim, value) in enumerate(cells):
            col = int(str(adelim).strip()) if use_adelim else i
            if col == 0 and not label:
                label = _label("" if value is None else str(value))
            amount = parse_amount(value)
            if amount is not None and col not in vals:
                vals[col] = amount
        out.append(CheckRow(eid, label, _level(alevel), str(acode or ""), vals))
    return tuple(out)


def _make_table(sheet: str, eid: str, aclass: str, raw_rows, raw_heads=()) -> CheckTable:
    """표 하나를 만들고 제외 사유까지 정한다 (`.dsd` 와 `.xlsx` 공통)."""
    table = CheckTable(sheet, eid, aclass, _make_rows(raw_rows), "",
                       tuple((int(col), "" if text is None else str(text))
                             for col, text in raw_heads))
    return CheckTable(sheet, eid, aclass, table.rows, _table_excluded(table), table.headers)


def _table_excluded(table: CheckTable) -> str:
    if _ce_plan(table) is not None:
        return ""
    rows = table.rows
    cols = {c for r in rows for c in r.vals} | {0}
    if not cols <= STD_COLS:
        extra = ", ".join(str(c) for c in sorted(cols - STD_COLS))
        return (
            f"논리 열 {extra} 은(는) 기간(당기 1·2 / 전기 3·4)이 아니어서 합계검증에서 "
            "제외합니다 (자본변동표 규칙에도 맞지 않습니다)."
        )
    if not any(r.level is not None for r in rows):
        return NO_LEVEL_REASON
    return ""


def model_from_dsd(path):
    """`.dsd` 하나를 열어 문서 모델을 만든다 (`tables_from_dsd` 와 상호대사가 함께 쓴다)."""
    from .container import DsdContainer
    from .model import DocModel
    from .xmlsplice import SourceIndex

    p = Path(path)
    if not p.is_file():
        raise InputError(f"오류: 입력 파일을 찾을 수 없습니다: {p}")
    raw = DsdContainer.read(p).contents_bytes
    return DocModel.build(SourceIndex.build(raw), raw)


def tables_from_dsd(path) -> list:
    """`.dsd` 에서 검사 대상 표를 모은다."""
    return tables_from_model(model_from_dsd(path))


def tables_from_model(model) -> list:
    """문서 모델에서 검사 대상 표를 모은다."""
    from .model import row_attrs, table_layout

    out = []
    for section in model.sections:
        for block in section.blocks:
            if block.table is None:
                continue
            t = block.table
            body = [r for r in t.rows if r.kind != "THEAD-TR"]
            has_attr = any(
                row_attrs(t, r)["alevel"] or row_attrs(t, r)["acode"] for r in body
            )
            if not (has_attr or t.extraction):
                continue
            layout = table_layout(t)
            heads = []
            for r in t.rows:
                if r.kind != "THEAD-TR":
                    continue
                cols = layout.cols_for(r.eid)
                heads.extend((cols[i] if i < len(cols) else i, c.text)
                             for i, c in enumerate(r.cells))
            out.append(_make_table(section.sheet_name, t.eid, t.aclass, [
                (r.eid, row_attrs(t, r)["alevel"], row_attrs(t, r)["acode"],
                 [(c.adelim, c.text) for c in r.cells])
                for r in body
            ], heads))
    return out


def tables_from_xlsx(path) -> list:
    """내보낸 `.xlsx` 에서 검사 대상 표를 모은다 (`_dsdx_attrs` + 시트 값)."""
    from .sheetstate import load_workbook_state

    attrs = read_attrs_sheet(path)
    if not attrs:
        raise InputError(
            f"오류: 통합문서에 '{ATTRS_SHEET}' 숨김 시트가 없어 합계검증을 할 수 없습니다: {path}\n"
            "  dsd-bridge 1.2 이상으로 `dsd-bridge export` 를 다시 실행해 통합문서를 새로 만드십시오."
        )
    by_id = {d["_ID"]: d for d in attrs if d.get("_ID")}
    state = load_workbook_state(path)

    out = []
    for sheet in state.sheets:
        pending: list = []
        heads: list = []
        table_eid = ""
        for r in sheet.rows:
            if r.base_kind == "TABLE-START":
                pending, heads, table_eid = [], [], r.table
                continue
            if r.base_kind == "TABLE-END":
                if pending:
                    _append_xlsx_table(out, sheet.name, table_eid, pending, heads)
                pending, heads, table_eid = [], [], ""
                continue
            if r.base_kind == "THEAD-TR" and r.id:
                heads.extend((r.col_of(i), r.value_for_cell(i))
                             for i in range(len(r.cell_ids)))
                continue
            if r.base_kind != "TR" or not r.id:
                continue
            a = by_id.get(r.id)
            if a is None:
                continue
            ids = r.cell_ids
            delims = a["adelims"].split("|") if a["adelims"] else [""] * len(ids)
            cells = [(delims[i] if i < len(delims) else "", r.value_for_cell(i))
                     for i in range(len(ids))]
            pending.append((a, r.id, cells))
        if pending:
            _append_xlsx_table(out, sheet.name, table_eid, pending, heads)
    return out


def _append_xlsx_table(out: list, sheet: str, table_eid: str, pending: list,
                       heads: list) -> None:
    aclass = pending[0][0]["aclass"]
    if not (any(a["alevel"] or a["acode"] for a, _, _ in pending) or aclass == "EXTRACTION"):
        return
    out.append(_make_table(
        sheet, table_eid, aclass,
        [(eid, a["alevel"], a["acode"], cells) for a, eid, cells in pending], heads,
    ))


# ─────────────────────────────────────────────────────────────────────────────
# 검사
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class CheckItem:
    """검사 한 건의 결과."""

    sheet: str
    table: str
    row: str
    label: str
    period: str
    expected: Decimal
    actual: Decimal
    diff: Decimal
    verdict: str
    structure: str
    note: str = ""            # `부호주의` 일 때 어느 하위를 차감으로 보면 맞는지

    def line_ko(self) -> str:
        line = (
            f"[{self.sheet}] {self.label} {self.period}: "
            f"입력 {format_number(self.expected, DEFAULT_PATTERN)} / "
            f"계산 {format_number(self.actual, DEFAULT_PATTERN)} / "
            f"차이 {format_number(self.diff, DEFAULT_PATTERN)} → {self.verdict}"
        )
        return f"{line} ({self.note})" if self.note else line

    def as_dict(self) -> dict:
        """`--json` 항목 — 금액은 정확성을 위해 **문자열**(정규 십진)이다."""
        return {
            "table": self.table,
            "sheet": self.sheet,
            "row": self.row,
            "label": self.label,
            "period": self.period,
            "expected": str(self.expected),
            "actual": str(self.actual),
            "diff": str(self.diff),
            "verdict": self.verdict,
            "structure": self.structure,
            "note": self.note,
        }


def _judge(table: CheckTable, row: CheckRow, period: str, expected: Decimal,
           actual: Decimal, structure: str, tolerance: int, alts=()) -> CheckItem:
    """검사 한 건을 판정한다.

    `alts` 는 `(대안 계산값, 사유)` 목록이다 — 하위의 부호를 뒤집었을 때의 합이며, 그중
    하나라도 허용 오차 안에 들면 `부호주의` 가 된다 (모듈 docstring 참고).
    """
    diff = expected - actual
    note = ""
    if diff == 0:
        verdict = OK
    elif abs(diff) <= tolerance:
        verdict = ROUNDING
    else:
        verdict = MISMATCH
        for alt, why in alts:
            if abs(expected - alt) <= tolerance:
                verdict, note = SIGN_NOTE, why
                break
    return CheckItem(table.sheet, table.eid, row.eid, row.label, period,
                     expected, actual, diff, verdict, structure, note)


def _sum(rows, period: int) -> Decimal:
    total = Decimal(0)
    for r in rows:
        v = r.amount(period)
        if v is not None:
            total += v
    return total


def _subtree(rows: tuple, i: int) -> list:
    """`rows[i]` 아래에 딸린 행 인덱스 (레벨이 더 깊은 연속 구간)."""
    base = rows[i].level
    j = i + 1
    while j < len(rows) and rows[j].level is not None and rows[j].level > base:
        j += 1
    return list(range(i + 1, j))


def _leaf_value(rows: tuple, i: int, period: int) -> Decimal:
    """총계 등식에서 레벨 0 행 하나가 기여하는 값 = **말단 하위 항목의 합**.

    하위가 없으면 그 행의 금액이다. 부모의 표시 금액 대신 말단을 쓰는 이유는 모듈 docstring
    참고(결손금 절댓값 표기).
    """
    sub = _subtree(rows, i)
    if not sub:
        return rows[i].amount(period) or Decimal(0)
    leaves = [k for k in sub
              if k == sub[-1] or rows[k + 1].level is None or rows[k + 1].level <= rows[k].level]
    return _sum([rows[k] for k in leaves], period)


def _tree_checks(table: CheckTable, tolerance: int) -> list:
    rows = table.rows
    n = len(rows)
    out = []
    for i, r in enumerate(rows):
        if r.level is None:
            continue
        kids, structure = [], ""
        if i + 1 < n and rows[i + 1].level == r.level + 1:
            structure = STRUCT_BELOW
            for k in _subtree(rows, i):
                if rows[k].level == r.level + 1:
                    kids.append(rows[k])
        elif i and rows[i - 1].level == r.level + 1 and not r.is_total:
            # 내역이 위에 있는 구조. 레벨 0 총계 행은 총계 등식이 따로 보므로 제외한다
            # (형제 소계를 이중으로 세지 않도록).
            j = i - 1
            while j >= 0 and rows[j].level is not None and rows[j].level > r.level:
                if rows[j].level == r.level + 1:
                    kids.append(rows[j])
                j -= 1
            if j >= 0 and rows[j].level == r.level:
                # 위쪽 내역이 **바로 앞 형제**의 것이다 (그 형제가 하단합계로 이미 가져간다).
                # 두 샘플의 `Ⅲ. 영업이익`(= 수익 − 비용)이나 `3. 현금의 유입이 없는 수익 등의
                # 차감` 이 여기 걸린다 — 앞 항목의 하위를 자기 합계로 오인하지 않도록 버린다.
                kids = []
            else:
                structure = STRUCT_ABOVE
                kids.reverse()
        if not kids:
            continue
        for p, name in enumerate(PERIOD_NAMES):
            expected = r.amount(p)
            if expected is None:
                continue
            actual = _sum(kids, p)
            out.append(_judge(table, r, name, expected, actual, structure, tolerance,
                              _sign_alternatives(kids, p, actual)))
    return out


def _sign_alternatives(kids, period: int, actual: Decimal) -> list:
    """차감 항목 추정용 대안 합 `(값, 사유)` — 하위 전부 뒤집기, 하위 하나씩 뒤집기."""
    alts = []
    if actual != 0:
        alts.append((-actual, "하위 합계의 부호를 뒤집으면 맞습니다"))
    for k in kids:
        v = k.amount(period)
        if v:
            alts.append((actual - 2 * v, f"하위 '{k.label}' 을(를) 차감으로 보면 맞습니다"))
    return alts


def _total_checks(table: CheckTable, tolerance: int) -> list:
    rows = table.rows
    out, group = [], []
    for i, r in enumerate(rows):
        if r.level != 0:
            continue
        if not r.is_total:
            group.append(i)
            continue
        if group:
            for p, name in enumerate(PERIOD_NAMES):
                expected = r.amount(p)
                if expected is None:
                    continue
                actual = sum((_leaf_value(rows, g, p) for g in group), Decimal(0))
                out.append(_judge(table, r, name, expected, actual, STRUCT_TOTAL, tolerance))
        group = []
    return out


def _bs_key(row: CheckRow) -> str:
    """대차평형에서 이 행이 어느 총계인지 — 과목명이 정확히 맞을 때만."""
    name = _nospace(row.label)
    for key, pattern in BS_LABELS:
        if pattern.match(name):
            return key
    return ""


def _balance_checks(table: CheckTable, tolerance: int) -> list:
    found: dict = {}
    for r in table.rows:
        if r.level != 0:
            continue
        key = _bs_key(r)
        if key and key not in found:
            found[key] = r
    assets = found.get("자산총계")
    if assets is None:
        return []
    both = found.get("부채와자본총계")
    parts = [found[k] for k in ("부채총계", "자본총계") if k in found]
    if both is not None:
        others, label = [both], f"{assets.label} = {both.label}"
    elif len(parts) == 2:
        others = parts
        label = f"{assets.label} = {parts[0].label} + {parts[1].label}"
    else:
        return []
    out = []
    for p, name in enumerate(PERIOD_NAMES):
        expected = assets.amount(p)
        if expected is None or any(o.amount(p) is None for o in others):
            continue
        row = CheckRow(assets.eid, label, assets.level, assets.acode, assets.vals)
        out.append(_judge(table, row, name, expected, _sum(others, p),
                          STRUCT_BALANCE, tolerance))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 자본변동표 (WP6)
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class CePlan:
    """자본변동표 검사 계획 — `_ce_plan` 이 만든다."""

    cols: tuple             # 금액 논리 열 (과목 열 0 제외), 오름차순
    total: int | None       # 총계 열. 없으면 가로 검사를 하지 않는다
    names: dict             # 논리 열 → 열 이름 (머리글, 공백 제거)
    anchors: tuple          # 잔액 행의 인덱스


def _ce_head_names(table: CheckTable) -> dict:
    """논리 열 → 머리글 이름 (공백 제거). 같은 열에 여럿이면 **먼저** 나온 것을 쓴다.

    `이  익&cr;잉여금` 처럼 글자를 벌려 놓거나 줄을 바꾼 머리글이 많아 `_label` 이 아니라
    공백을 아예 지운 이름을 쓴다 (`이익잉여금` · `총계`).
    """
    out: dict = {}
    for col, text in table.headers:
        name = _nospace(text)
        if name and col not in out:
            out[col] = name
    return out


def _is_ce_balance(row: CheckRow) -> bool:
    """잔액 행인가 — 과목명에 `기초`/`기말` 이 있거나 (번호 접두 뒤) 날짜로 시작한다."""
    name = _nospace(row.label)
    return bool(_RE_CE_BALANCE.search(name) or _RE_CE_DATE_HEAD.match(name))


def _ce_date(label: str):
    """과목명에서 처음 나오는 날짜 → `date` (없거나 말이 안 되면 `None`)."""
    m = _RE_CE_DATE.search(_nospace(label))
    if not m:
        return None
    try:
        return date(int(m[1]), int(m[2]), int(m[3]))
    except ValueError:
        return None


def _ce_plan(table: CheckTable):
    """이 표를 자본변동표로 검사할지 — 맞으면 `CePlan`, 아니면 `None`.

    시트(제목) 이름에 `자본변동표` 가 있어야 하고, 아니면 논리 열이 `0~4` 를 벗어나야 한다.
    그 위에 실제로 검사할 것이 있어야 한다 — 시트 이름이 말해 주면 **총계 열 머리글이나 잔액 행
    둘**(가로 또는 세로) 중 하나로 충분하지만, 논리 열만 보고 고른 표는 **둘 다** 요구한다.
    `외부감사 실시내용` 의 표들도 논리 열이 `0~4` 를 벗어나고 그 안에 날짜 셀이 있어서, 느슨하게
    잡으면 자본변동표로 오인된다(시험한 보고서 대부분에서 확인).
    """
    rows = table.rows
    cols = sorted({c for r in rows for c in r.vals if c != 0})
    if not cols:
        return None
    titled = CE_TITLE in _nospace(table.sheet)
    if not titled and set(cols) <= STD_COLS:
        return None
    names = _ce_head_names(table)
    total = next((c for c in cols if _RE_CE_TOTAL_HEAD.match(names.get(c, ""))), None)
    anchors = tuple(i for i, r in enumerate(rows) if _is_ce_balance(r))
    has_row, has_col = total is not None, len(anchors) >= 2
    enough = (has_row or has_col) if titled else (has_row and has_col)
    if not enough:
        return None
    return CePlan(tuple(cols), total, names, anchors)


def _ce_name(plan: CePlan, col: int) -> str:
    return plan.names.get(col) or f"열 {col}"


def _ce_sum(rows, col: int) -> Decimal:
    """자본변동표에서는 빈칸도 `-` 와 같이 0 으로 센다 (그 칸에 변동이 없다는 뜻)."""
    return sum((r.vals.get(col, Decimal(0)) for r in rows), Decimal(0))


def _ce_checks(table: CheckTable, plan: CePlan, tolerance: int) -> list:
    rows = table.rows
    out = []
    parts = [c for c in plan.cols if c != plan.total] if plan.total is not None else []
    if parts:
        name = _ce_name(plan, plan.total)
        for r in rows:
            expected = r.vals.get(plan.total)
            if expected is None:
                continue
            actual = sum((r.vals.get(c, Decimal(0)) for c in parts), Decimal(0))
            out.append(_judge(table, r, name, expected, actual, STRUCT_CE_ROW, tolerance))
    for a, b in zip(plan.anchors, plan.anchors[1:]):
        moves = rows[a + 1:b]
        if not moves:
            start, end = _ce_date(rows[a].label), _ce_date(rows[b].label)
            if start is None or end is None or end - start != timedelta(days=1):
                continue                      # 이어지지 않는 잔액 (반기보고서의 전반기말 → 당반기초)
        structure = STRUCT_CE_COL if moves else STRUCT_CE_CARRY
        for col in plan.cols:
            expected, base = rows[b].vals.get(col), rows[a].vals.get(col)
            if expected is None or base is None:
                continue
            out.append(_judge(table, rows[b], _ce_name(plan, col), expected,
                              base + _ce_sum(moves, col), structure, tolerance))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 보고
# ─────────────────────────────────────────────────────────────────────────────

#: 불일치가 하나라도 있을 때의 종료 코드 (README 종료 코드 표) — `dsd_bridge.errors` 가 원본이다.
MISMATCH_EXIT_CODE = FOOTING_MISMATCH_EXIT


@dataclass
class CheckReport:
    path: Path
    tolerance: int
    items: list = field(default_factory=list)
    #: 검사에서 제외한 표 `[(시트, 표 ID, 사유)]`.
    excluded: list = field(default_factory=list)
    tables: int = 0
    #: 상호대사 결과 (`--xref`). `None` 이면 하지 않았다는 뜻이며, 종료 코드에는 영향이 없다.
    xref: object = None
    #: 정산표 상호대사 결과 (`--settlement`, 1.3). `None` 이면 하지 않았다는 뜻이다.
    #: 금액 불일치가 있으면 합계검증과 마찬가지로 종료 코드 8 이다.
    settlement: object = None
    #: 확인 필요 문장 (0.4.0, `dsd_bridge.review`). 자동 판정이 아니며 종료 코드에 영향이 없다.
    review: list = field(default_factory=list)

    def counts(self) -> dict:
        out = {v: 0 for v in VERDICTS}
        for it in self.items:
            out[it.verdict] += 1
        return out

    @property
    def exit_code(self) -> int:
        if self.counts()[MISMATCH]:
            return MISMATCH_EXIT_CODE
        if self.settlement is not None and self.settlement.exit_code:
            return self.settlement.exit_code
        return 0

    def note_ko(self) -> str:
        """검사한 표가 하나도 없을 때 그 이유를 알리는 `안내:` 줄 (아니면 빈 문자열).

        `OK 0 / 단수차 0 / 불일치 0` 만 나오면 "문제 없음" 으로 오해하기 쉬워서, 검사를
        아예 못 했다는 사실을 기본 출력에도 남긴다. 판정이 아니므로 종료 코드는 0 이다.
        """
        if self.tables:
            return ""
        extra = (f" (검사에서 제외한 표 {len(self.excluded)}개 — 사유는 -v 로 볼 수 있습니다)"
                 if self.excluded else "")
        return NOTE_NO_TABLE.format(extra=extra)

    def summary_ko(self) -> str:
        c = self.counts()
        line = f"합계검증: OK {c[OK]} / 단수차 {c[ROUNDING]} / 불일치 {c[MISMATCH]}"
        if c[SIGN_NOTE]:
            line += f" / {SIGN_NOTE} {c[SIGN_NOTE]}"
        return line

    def lines_ko(self, verbose: bool = False) -> list:
        """출력할 줄. 기본은 OK 가 아닌 것만, `-v` 면 전부.

        합계검증 · 주석 상호대사(`--xref`) · 정산표 상호대사(`--settlement`)는 **서로 다른
        검사**이므로 빈 줄로 나눈 블록으로 낸다.
        """
        lines = ["── 합계검증 ──"] if self.settlement is not None else []
        if verbose:
            lines.append(
                f"검사 대상 표 {self.tables}개 / 검사 {len(self.items)}건 "
                f"(단수차 허용 ±{self.tolerance})"
            )
            for sheet, eid, why in self.excluded:
                lines.append(f"안내: [{sheet}] 표 {eid} — {why}")
        for it in self.items:
            if it.verdict == OK and not verbose:
                continue
            line = it.line_ko()
            if verbose:
                line += f"  ({it.structure})"
            lines.append(line)
        note = self.note_ko()
        if note:
            lines.append(note)
        lines.append(self.summary_ko())
        if self.xref is not None:
            lines.append("")
            lines.extend(self.xref.lines_ko(verbose))
        if self.settlement is not None:
            lines.append("")
            lines.extend(self.settlement.lines_ko(verbose))
        if self.review:
            lines.append("")
            lines.append(f"── 확인 필요 {len(self.review)}건 (자동 판정 아님) ──")
            lines.extend(f"확인: {r}" for r in self.review)
        return lines

    def json_obj(self):
        """`--json` 의 내용.

        상호대사를 하지 않았으면 **검사 항목 목록**(1.2 WP4 그대로)이고, `--xref` 나
        `--settlement` 을 주면 `{"items": […], "xref": …, "settlement": …}` 객체가 된다 —
        기존 결과의 모양은 그대로 두고 옆에 붙인다.
        """
        items = [it.as_dict() for it in self.items]
        if self.xref is None and self.settlement is None:
            return items
        obj = {"items": items}
        if self.review:
            obj["review"] = list(self.review)
        if self.xref is not None:
            obj["xref"] = self.xref.json_obj()
        if self.settlement is not None:
            obj["settlement"] = self.settlement.json_obj()
        return obj

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(
                json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except OSError as exc:
            raise OutputError(f"오류: 검사 결과를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


def check_tables(tables: list, path, tolerance: int = DEFAULT_TOLERANCE) -> CheckReport:
    report = CheckReport(Path(path), tolerance)
    for t in tables:
        if t.excluded:
            report.excluded.append((t.sheet, t.eid, t.excluded))
            continue
        report.tables += 1
        plan = _ce_plan(t)
        if plan is not None:
            # 자본변동표는 열이 기간이 아니라 자본 구성요소라 표준 검사(세로 합계·총계 등식·
            # 대차평형)를 쓰면 엉뚱한 값을 견준다 — 자본변동표 규칙만 쓴다.
            report.items.extend(_ce_checks(t, plan, tolerance))
            continue
        report.items.extend(_tree_checks(t, tolerance))
        report.items.extend(_total_checks(t, tolerance))
        report.items.extend(_balance_checks(t, tolerance))
    return report


def check(path, tolerance: int = DEFAULT_TOLERANCE, *, xref: bool = False,
          xref_min: int = 0) -> CheckReport:
    """`.dsd` 또는 `.xlsx` 하나를 합계검증한다.

    `xref=True` 면 주석 ↔ 재무제표 **상호대사**(1.2 WP7)를 덧붙인다. 참고용이라 종료 코드에는
    영향이 없으며, 주석의 구조(번호 제목·단위)를 읽어야 하므로 **`.dsd` 입력에서만** 된다.
    """
    if tolerance < 0:
        raise InputError(f"오류: --tolerance 는 0 이상이어야 합니다: {tolerance}")
    if xref_min < 0:
        raise InputError(f"오류: --xref-min 은 0 이상이어야 합니다: {xref_min}")
    p = Path(path)
    suffix = p.suffix.lower()
    model = None
    if suffix == ".xlsx":
        if xref:
            raise InputError(
                f"오류: 상호대사(--xref)는 .dsd 입력에서만 할 수 있습니다: {p}\n"
                "  통합문서에는 주석 번호 제목과 단위 표기를 가려낼 문서 구조가 남지 않습니다. "
                "원본 .dsd 를 검사하십시오."
            )
        tables = tables_from_xlsx(p)
    elif suffix == ".dsd":
        model = model_from_dsd(p)
        tables = tables_from_model(model)
    else:
        raise InputError(
            f"오류: 합계검증은 .dsd 또는 .xlsx 만 받습니다: {p}"
        )
    report = check_tables(tables, p, tolerance)
    try:
        from .docmap import load as load_map
        from .review import review

        report.review = review(load_map(p))
    except Exception:                       # pragma: no cover - 검토 목록은 덤이라 실패해도 합계검증은 낸다
        report.review = []
    if xref:
        from .xref import build as build_xref

        report.xref = build_xref(tables, model, tolerance, xref_min)
    return report


__all__ = [
    "DEFAULT_TOLERANCE",
    "MISMATCH_EXIT_CODE",
    "NOTE_NO_TABLE",
    "NO_LEVEL_REASON",
    "PERIOD_NAMES",
    "STD_COLS",
    "OK",
    "ROUNDING",
    "SIGN_NOTE",
    "MISMATCH",
    "VERDICTS",
    "STRUCT_BELOW",
    "STRUCT_ABOVE",
    "STRUCT_TOTAL",
    "STRUCT_BALANCE",
    "STRUCT_CE_ROW",
    "STRUCT_CE_COL",
    "STRUCT_CE_CARRY",
    "CE_TITLE",
    "CheckRow",
    "CheckTable",
    "CheckItem",
    "CheckReport",
    "parse_amount",
    "model_from_dsd",
    "tables_from_model",
    "tables_from_dsd",
    "tables_from_xlsx",
    "check_tables",
    "check",
]
