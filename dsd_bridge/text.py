"""텍스트·숫자 변환 규칙 (계획 §4.5, §4.6, §4.7 / AC-12′, 13, 14, 15, 16b, 21).

두 부분으로 이루어진다.

1. **텍스트 변환** — `cr_to_lf` / `lf_to_cr` / `escape_text` / `unescape` /
   `merge_mixed_text` / `count_bare_amp` (§4.6, §4.7).
2. **숫자 규칙** — `NumPattern` / `classify` / `format_number` / `excel_number_format` /
   `canonical_out` / `resolve_pattern` (§4.5).

변환 사슬 (계획 §4.6)::

    XML 바이트 ──unescape──▶ "…&cr;1. 회사의 개요…"
                ──cr_to_lf──▶ "…\\n1. 회사의 개요…"      (Excel 셀, wrap_text=True)
                ──lf_to_cr──▶ "…&cr;1. 회사의 개요…"
              ──escape_text──▶ "…&amp;cr;1. 회사의 개요…"

**규범 1 (계획 §4.4)**: `classify` / `format_number` / `canonical_out` / `merge_mixed_text`
중 어느 하나라도 동작이 바뀌면 `dsd_bridge.model.WORKBOOK_FORMAT_VERSION` 의 마이너를 올리고
`tests/golden_project.json` 을 갱신해야 한다. 버그 수정이라도 예외 없다.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Context, Decimal, InvalidOperation

#: DART 편집기가 셀·문단 안의 줄바꿈을 나타내는 데 쓰는 리터럴 4글자.
CR_TOKEN = "&cr;"

_CHARREF = re.compile(r"&(#[0-9]+|#[xX][0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]*);")

_NAMED = {
    "amp": "&",
    "lt": "<",
    "gt": ">",
    "quot": '"',
    "apos": "'",
}


def cr_to_lf(s: str) -> str:
    """`&cr;` 를 개행 `\\n` 으로 바꾼다 (AC-14).

    입력은 **이미 언이스케이프된** 문자열이다. 원본 바이트의 `&amp;cr;` 는
    `unescape()` 를 거쳐 `&cr;` 4글자가 되어 있어야 한다.
    """
    return s.replace(CR_TOKEN, "\n")


def lf_to_cr(s: str) -> str:
    """개행을 `&cr;` 로 되돌린다 (AC-14). `\\r\\n` → `\\n` → `&cr;` 순서."""
    return s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", CR_TOKEN)


def escape_text(s: str) -> str:
    """텍스트 노드 이스케이프 — `&`, `<`, `>` **만** 바꾼다 (AC-15).

    `"` 와 `'` 는 건드리지 않는다. 원본에 `&quot;`/`&apos;` 가 0건이고
    본문의 `'당사'`(샘플A:1723)·`"회사"`(샘플A:243)가 raw 이기 때문이다.
    `xml.sax.saxutils.escape(s)` 의 기본 동작과 같다.
    """
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def unescape(s: str) -> str:
    """XML 문자 참조를 되돌린다 (표준 5개 명명 엔티티 + 숫자 참조).

    `escape_text` 의 역함수이면서, 원본에 실제로 나타나는 형태
    (`&amp;` 198/143, `&gt;` 3/4, `&lt;` 1/2)를 모두 처리한다.
    한 번의 정규식 훑기로 처리하므로 `&amp;lt;` 는 `&lt;` 가 된다(이중 해제 없음).
    알 수 없는 참조는 원문 그대로 남긴다.
    """

    def sub(m: "re.Match[str]") -> str:
        ref = m.group(1)
        if ref.startswith("#"):
            try:
                code = int(ref[2:], 16) if ref[1] in "xX" else int(ref[1:])
            except ValueError:
                return m.group(0)
            try:
                return chr(code)
            except ValueError:
                return m.group(0)
        return _NAMED.get(ref, m.group(0))

    return _CHARREF.sub(sub, s)


_BARE_AMP = re.compile(r"&(?!cr;)")


def count_bare_amp(s: str) -> int:
    """언이스케이프된 텍스트에서 `&cr;` 가 **아닌** `&` 의 개수 (계획 §4.6, `-v` 보고용).

    두 샘플에서 0건이지만, 사용자가 `&` 를 직접 입력하면 `escape_text` 가 `&amp;` 로
    바꾸므로 왕복은 안전하다. 이 함수는 그 사실을 사용자에게 알리는 데만 쓴다.
    """
    return len(_BARE_AMP.findall(s))


def merge_mixed_text(index, raw: bytes, ordinal: int) -> str:
    """혼합 콘텐츠 요소의 표시 텍스트 — **순수 연접** (계획 §4.7, AC-16b).

    `text + Σ(child.text + child.tail)` 을 문서 순서 그대로 잇는다. 자식 사이에 `\\n`
    같은 구분자를 **넣지 않는다** — `샘플A:244` 의 `<SPAN>감사의견근거&amp;cr;</SPAN>우리는…`
    처럼 `SPAN` 은 제목 강조이고 그 직후 텍스트가 같은 문장의 연속이기 때문이다.

    구현은 원본 바이트를 `SourceIndex` 의 자식 구간으로 잘라 붙인다(자식 안의 자식도
    재귀적으로). 그런 다음 언이스케이프하고, expat 과 같은 시각이 되도록 CRLF 를 LF 로
    정규화한다(계획 §0.5 사실 0). 결과에는 `&cr;` 가 4글자 그대로 남아 있다 —
    Excel 표시용 `\\n` 변환은 호출 측이 `cr_to_lf` 로 한다.

    순수 텍스트 요소(자식 없음)에 불러도 동작하며 그때는 그 요소의 텍스트와 같다.
    """
    spans = index.spans

    def inner(o: int) -> bytes:
        sp = spans[o]
        if sp.self_closing:
            return b""
        parts = []
        pos = sp.start_gt
        for c in sp.child_ordinals:
            ch = spans[c]
            parts.append(raw[pos : ch.start_lt])
            parts.append(inner(c))
            pos = ch.end_gt
        parts.append(raw[pos : sp.end_lt])
        return b"".join(parts)

    text = unescape(inner(ordinal).decode("utf-8"))
    return text.replace("\r\n", "\n").replace("\r", "\n")


# ─────────────────────────────────────────────────────────────────────────────
# 숫자 규칙 (계획 §4.5)
# ─────────────────────────────────────────────────────────────────────────────

#: 절댓값이 이 값 이상이면 숫자로 내보내지 않는다 (AC-21, IEEE754 정확 표현 한계).
MAX_SAFE_ABS = 2**53
#: 소수 자릿수 상한 (AC-21, §4.5 `NUMPAT.dec` 범위 0..5).
MAX_DEC = 5

#: 숫자로 취급되는 분류 (§0.4 "숫자로 분류된 셀" = int + paren + dec, 그리고 코퍼스 0건인 minus).
NUMERIC_CLASSES = frozenset({"int", "paren", "minus", "dec"})

#: 분류 이름 전체 (표 출력·검증용, §4.5 표의 순서).
ALL_CLASSES = ("mixed", "empty", "dash", "ws", "paren", "minus", "dec", "int", "text")

# §4.5 분류 정규식 — `\s*` 를 넣지 않는다(MINOR-29: 포맷터가 `(1,234)` 를 내므로 파서가
# `( 1,234 )` 를 받으면 왕복 비대칭). 실측상 샘플A 에 `'2,030 '`(뒤 공백) 1건이 있는데 이 규칙대로
# `text` 다 — §0.4 의 int 209 는 그 셀을 뺀 값이다. **strip 하지 않는다**(바이트 보존 원칙).
# 자릿수 부분은 `int` 와 같이 `콤마 묶음 | 콤마 없는 연속 숫자` 를 모두 받는다 — 콤마 없이 쓴
# `1234.5` / `(1234.5)` / `-1234.5` / `(1234)` / `-1234` 도 숫자로 분류된다(1.2).
# `grp` 는 실제로 콤마가 있었는지로 정해지므로 복원 표기는 원문 그대로다.
_DIGITS = r"(\d{1,3}(?:,\d{3})*|\d+)"
_RE_DASH = re.compile("^[\\-–—－]+$")     # - – — －
_RE_WS = re.compile("^[\\s　]+$")                    # 공백류 + 전각 공백
_RE_PAREN = re.compile(r"^\(" + _DIGITS + r"(\.\d+)?\)$")
_RE_MINUS = re.compile(r"^-" + _DIGITS + r"(\.\d+)?$")
_RE_DEC = re.compile(r"^" + _DIGITS + r"(\.\d+)$")
_RE_INT = re.compile(r"^" + _DIGITS + r"$")

#: 숫자 변환에 쓰는 십진 컨텍스트 — 기본 정밀도(28)로도 충분하지만 2^53 근처의 큰 값에
#: 소수 5자리를 붙여도 자리가 모자라지 않도록 넉넉히 잡는다.
_CTX = Context(prec=60, rounding=ROUND_HALF_UP)
_QUANT_5 = Decimal("0.00001")


@dataclass(frozen=True)
class NumPattern:
    """원본 숫자 셀의 표기 패턴 `NUMPAT` (계획 §4.5).

    - `grp` — 천 단위 콤마 유무. 콤마 없는 순수 숫자(`5`, `7`)는 False 로 기록해
      복원 시 콤마를 넣지 않는다.
    - `dec` — 소수 자릿수 (0..5).
    - `neg` — 음수 표기 `"paren"` | `"minus"`. **양수 셀의 기본값은 `"paren"`**
      (양 코퍼스의 음수 표기가 100% 괄호이므로, MAJOR-6).
    """

    grp: bool
    dec: int
    neg: str = "paren"

    def __post_init__(self) -> None:
        if self.neg not in ("paren", "minus"):
            raise ValueError(f"neg 는 'paren' 또는 'minus' 여야 합니다: {self.neg!r}")
        if not 0 <= self.dec <= MAX_DEC:
            raise ValueError(f"dec 는 0..{MAX_DEC} 여야 합니다: {self.dec!r}")

    def as_dict(self) -> dict:
        return {"grp": self.grp, "dec": self.dec, "neg": self.neg}

    @classmethod
    def from_dict(cls, d: dict) -> "NumPattern":
        return cls(bool(d["grp"]), int(d["dec"]), str(d.get("neg", "paren")))


#: 복원 규칙 4 — 아무 단서도 없을 때의 패턴.
DEFAULT_PATTERN = NumPattern(grp=True, dec=0, neg="paren")


@dataclass(frozen=True)
class Classified:
    """`classify` 의 결과 — 분류 이름, 숫자값(숫자일 때만), 표기 패턴(숫자일 때만)."""

    cls: str
    value: Decimal | None = None
    pattern: NumPattern | None = None

    @property
    def numeric(self) -> bool:
        return self.cls in NUMERIC_CLASSES


def _to_decimal(value) -> Decimal:
    """int / float / str / Decimal → Decimal. float 은 `repr` 을 거쳐 잡음을 줄인다."""
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):          # bool 은 int 의 자손이지만 숫자로 보지 않는다
        raise TypeError("bool 은 숫자로 변환하지 않습니다.")
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(repr(value))
    if isinstance(value, str):
        return Decimal(value.strip())
    raise TypeError(f"숫자로 변환할 수 없는 값입니다: {type(value).__name__}")


def _decimal_places(d: Decimal) -> int:
    """유효한 소수 자릿수 (`1234567890.0` → 0, `2.50` → 1)."""
    exp = d.normalize(_CTX).as_tuple().exponent
    return max(0, -exp) if isinstance(exp, int) else 0


def classify(s: str, *, mixed: bool = False) -> Classified:
    """원문 텍스트 `s` 를 §4.5 표의 규칙으로 **위에서 아래로 첫 매치** 분류한다.

    `mixed=True` 는 요소가 자식 요소를 가진다는 뜻이며 규칙 1 로 즉시 `mixed` 다.
    규칙 5~8 에 매치했더라도 `abs(value) >= 2**53` 이거나 소수 5자리 초과면 `text` 로
    강등한다 (AC-21).
    """
    if mixed:
        return Classified("mixed")
    if s == "":
        return Classified("empty")
    if _RE_DASH.match(s):
        return Classified("dash")
    if _RE_WS.match(s):
        return Classified("ws")

    m = _RE_PAREN.match(s)
    if m:
        cls, sign, neg = "paren", -1, "paren"
    else:
        m = _RE_MINUS.match(s)
        if m:
            cls, sign, neg = "minus", -1, "minus"
        else:
            m = _RE_DEC.match(s)
            if m:
                cls, sign, neg = "dec", 1, "paren"
            else:
                m = _RE_INT.match(s)
                if m:
                    cls, sign, neg = "int", 1, "paren"
                else:
                    return Classified("text")

    digits = m.group(1)
    frac = m.group(2) if cls != "int" else None
    dec = len(frac) - 1 if frac else 0
    try:
        value = _CTX.create_decimal(digits.replace(",", "") + (frac or ""))
    except InvalidOperation:  # pragma: no cover - 정규식이 이미 형식을 보장한다
        return Classified("text")
    if sign < 0:
        value = -value
    if abs(value) >= MAX_SAFE_ABS or dec > MAX_DEC or not _float_exact(value):
        return Classified("text")
    return Classified(cls, value, NumPattern(grp="," in digits, dec=dec, neg=neg))


def _float_exact(d: Decimal) -> bool:
    """Excel(IEEE754 배정밀도)이 이 값을 **정확히** 담을 수 있는가.

    AC-21 의 `abs < 2**53` 은 정수부만 본다. `1,234,567,890,123,456.7` 은 2^53 미만이지만 소수까지 합쳐
    17자리라 float 으로 가면 `…456.8` 이 되고, import 가 그것을 "사용자 변경" 으로 오인해 원본을 다시 쓴다.
    그래서 float 왕복이 값을 보존하지 못하면 숫자로 내보내지 않는다(문자열 셀).
    """
    try:
        return Decimal(repr(float(d))) == d
    except (OverflowError, ValueError, InvalidOperation):
        return False


def format_number(value, pat: NumPattern) -> str:
    """숫자값을 `pat` 의 표기로 되돌린다 (§4.5 복원 규칙 5, AC-13).

    - 소수 자릿수: 값의 실제 자릿수가 `pat.dec` 보다 크면 실제 자릿수(최대 5, ROUND_HALF_UP),
      작으면 `pat.dec` 자리로 0 을 채운다.
    - 음수: `pat.neg` 가 `paren` 이면 `(1,234)`, `minus` 면 `-1,234`.
    - 콤마: `pat.grp` 일 때만.

    원본 숫자 셀 전부(샘플A 379 / 샘플B 297)에서 `format_number(value, pattern) == 원문` 이다
    (테스트로 강제).
    """
    d = _to_decimal(value)
    places = min(max(pat.dec, _decimal_places(d)), MAX_DEC)
    q = abs(d).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP, context=_CTX)
    body = f"{q:,.{places}f}" if pat.grp else f"{q:.{places}f}"
    if d < 0 and q != 0:
        return f"({body})" if pat.neg == "paren" else f"-{body}"
    return body


def excel_number_format(pat: NumPattern) -> str:
    """`pat` 에 대응하는 Excel `number_format` (AC-12: `#,##0` 계열)."""
    base = "#,##0" if pat.grp else "0"
    if pat.dec:
        base += "." + "0" * pat.dec
    return f"{base};({base})" if pat.neg == "paren" else f"{base};-{base}"


def canonical_out(value) -> str:
    """변경 판정용 표준형 문자열 `EXCEL_OUT` (§4.5 "변경 판정").

    - 숫자(int/float/Decimal) → `Decimal` 을 소수 5자리로 `quantize` 한 문자열.
      `1234567890` 과 `1234567890.0` 이 같은 값이 되어 부동소수 잡음이 무해해진다.
    - 문자열 → 개행을 `\\n` 으로 통일한 값.
    - `None` → 빈 문자열 (빈 Excel 셀).
    - 그 밖(bool, 날짜 등)은 `str()` 로 문자열 취급한다.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float, Decimal)):
        try:
            d = _to_decimal(value)
            if d == 0:
                d = abs(d)              # Excel 의 -0.0 이 "변경" 으로 집계되지 않도록
            return str(d.quantize(_QUANT_5, rounding=ROUND_HALF_UP, context=_CTX))
        except (InvalidOperation, ValueError):
            return str(value)
    s = str(value)
    return s.replace("\r\n", "\n").replace("\r", "\n")


@dataclass(frozen=True)
class PatternResolution:
    """`resolve_pattern` 의 결과 — 채택한 패턴과 그 출처 (SHOULD-6 `-v` 리포트용)."""

    pattern: NumPattern
    source: object | None        # 출처 셀 (`dsd_bridge.model.Cell`) — 표 최빈값/기본값이면 None
    level: str                   # "self" | "column" | "column-negative" | "table" | "default"

    def describe(self) -> str:
        base, plus, _ = self.level.partition("+")
        text = PatternResolution(self.pattern, self.source, base).describe() if plus else None
        if text is not None:
            return text + " (콤마는 같은 열·표에서 4자리 이상 숫자의 표기를 따름)"
        if self.level == "self":
            return "원본 셀의 표기 패턴"
        if self.level in ("column", "column-negative"):
            src = self.source
            where = f"{src.eid} (원본 {src.line}행)" if src is not None else "?"
            why = " — 실제 음수가 관측된 셀" if self.level == "column-negative" else ""
            return f"같은 열의 {where} 에서 상속{why}"
        if self.level == "table":
            return "같은 표의 최빈 표기 패턴"
        return "기본 패턴 (콤마, 소수 0자리, 괄호 음수)"


#: 천 단위 콤마가 나타날 수 있는 가장 작은 절댓값. 이보다 작은 원본 숫자(`7`, `(120)`)는 콤마가 없어도
#: 그 칸이 "콤마 없이 쓰는 칸" 이라는 증거가 되지 않는다.
GROUPING_MIN = Decimal(1000)


def _grouping_evidence(c) -> bool:
    """이 셀의 표기가 콤마 사용 여부를 실제로 보여 주는가 — 콤마가 있거나, 4자리 이상인데 콤마가 없을 때."""
    if c.pattern is None:
        return False
    if c.pattern.grp:
        return True
    return c.value is not None and abs(c.value) >= GROUPING_MIN


def resolve_pattern(cell, table, *, negative: bool = False, value=None) -> PatternResolution:
    """변경된 셀에 적용할 `NumPattern` 을 §4.5 복원 규칙 1~4 로 정한다 (AC-13).

    1. 그 셀에 패턴이 있으면(원본이 숫자였음) 그대로.
    2. 없으면 **같은 표, 같은 열 인덱스**에서 가장 가까운 패턴 보유 셀 — **위로 먼저,
       없으면 아래로**. `negative=True`(사용자가 음수를 입력) 면 그 열에서 실제 음수가
       관측된 셀(`paren`/`minus`)을 같은 순서로 먼저 찾는다.
    3. 없으면 같은 표 전체의 최빈 패턴.
    4. 없으면 `DEFAULT_PATTERN`.

    **콤마 보정 (0.3.0)**: `value`(새 값)를 주면, 위에서 고른 패턴이 콤마 없음인데 그 근거가
    3자리 이하 숫자뿐이고(`7` 은 콤마를 쓸 일이 없었을 뿐이다) 새 값이 1,000 이상이면, 콤마 여부만
    같은 열에서 가장 가까운 "근거 있는 셀"(콤마가 있거나 4자리 이상인 숫자)을 따른다. 열에 없으면
    표 전체의 다수결, 그것도 없으면 콤마를 쓴다. 소수 자릿수·음수 표기는 고른 패턴 그대로다.
    출처 표시는 `level` 에 `+grp` 를 붙인다. 원본 숫자 셀을 그대로 두는 왕복에는 영향이 없다.

    `cell` 은 `dsd_bridge.model.Cell`, `table` 은 `dsd_bridge.model.Table` 이다. 열 인덱스는 `TR` 내
    순서이므로 `ROWSPAN` 표에서는 시각 열과 다를 수 있다 — 그래서 출처를 돌려준다.
    """
    res = _resolve_base(cell, table, negative=negative)
    if value is None or res.pattern.grp:
        return res
    try:
        big = abs(_to_decimal(value)) >= GROUPING_MIN
    except (TypeError, InvalidOperation, ValueError):
        return res
    if not big:
        return res
    src = res.source
    if res.level == "default" or (src is not None and _grouping_evidence(src)):
        return res
    grp = _column_grouping(cell, table)
    if grp is None:
        grp = _table_grouping(table)
    if grp is None:
        grp = True
    if grp == res.pattern.grp:
        return res
    pat = NumPattern(grp=grp, dec=res.pattern.dec, neg=res.pattern.neg)
    return PatternResolution(pat, src, res.level + "+grp")


def _column_cells(cell, table):
    """같은 표·같은 열 인덱스의 다른 셀 — 위쪽을 가까운 순으로, 그다음 아래쪽을 가까운 순으로."""
    col = cell.col
    r = cell.row_index
    rows = table.rows
    for i in range(r - 1, -1, -1):
        cells = rows[i].cells
        if col < len(cells):
            yield cells[col]
    for i in range(r + 1, len(rows)):
        cells = rows[i].cells
        if col < len(cells):
            yield cells[col]


def _column_grouping(cell, table):
    for c in _column_cells(cell, table):
        if _grouping_evidence(c):
            return c.pattern.grp
    return None


def _table_grouping(table):
    yes = no = 0
    for row in table.rows:
        for c in row.cells:
            if _grouping_evidence(c):
                if c.pattern.grp:
                    yes += 1
                else:
                    no += 1
    if not yes and not no:
        return None
    return yes >= no


def _resolve_base(cell, table, *, negative: bool = False) -> PatternResolution:
    """§4.5 복원 규칙 1~4 (콤마 보정 전)."""
    if cell.pattern is not None:
        return PatternResolution(cell.pattern, cell, "self")

    def column_candidates():
        for c in _column_cells(cell, table):
            if c.pattern is not None:
                yield c

    if negative:
        for c in column_candidates():
            if c.cls in ("paren", "minus"):
                return PatternResolution(c.pattern, c, "column-negative")
    for c in column_candidates():
        return PatternResolution(c.pattern, c, "column")

    counter: Counter = Counter()
    for row in table.rows:
        for c in row.cells:
            if c.pattern is not None:
                counter[c.pattern] += 1
    if counter:
        # most_common 은 동률일 때 먼저 본 것을 앞에 둔다 — 결정적.
        return PatternResolution(counter.most_common(1)[0][0], None, "table")
    return PatternResolution(DEFAULT_PATTERN, None, "default")


__all__ = [
    "CR_TOKEN",
    "MAX_SAFE_ABS",
    "MAX_DEC",
    "NUMERIC_CLASSES",
    "ALL_CLASSES",
    "DEFAULT_PATTERN",
    "NumPattern",
    "Classified",
    "PatternResolution",
    "cr_to_lf",
    "lf_to_cr",
    "escape_text",
    "unescape",
    "count_bare_amp",
    "merge_mixed_text",
    "classify",
    "format_number",
    "excel_number_format",
    "canonical_out",
    "resolve_pattern",
]
