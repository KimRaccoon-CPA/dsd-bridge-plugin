"""`dsd-bridge check --settlement` — 보고서 ↔ 정산표 상호대사 (1.3 flow 도구).

보고서 형식에 맞춰 썼던 대사 스크립트를 일반화한 것이다.
보고서(`.dsd` 또는 내보낸 `.xlsx`)의 재무제표·주석 행을 **정규화 라벨**로 정산표 행과 짝지어
금액을 견준다.

짝짓기
------
- 보고서 쪽: 시트를 훑다가 `재무상태표`·`손익계산서`·`자본변동표`·`현금흐름표` 라는 **짧은**
  제목을 만나면 그다음 표 행들을 그 구분(BS/IS/CE/CF)으로 본다. 이름에 `주석` 이 든 시트는
  전부 `FN` 이다.
- 정산표 쪽: 기본 대응은 `DEFAULT_MAP` — BS·IS 는 이름에 `공시용`/`BS`/`IS` 가 든 시트,
  CE 는 `CE`, CF 는 `CF`, FN 은 `FN` 이다. 이름으로 못 찾으면 시트 **내용**에서
  `재무상태표`/`손익계산서`/`자본변동표`/`현금흐름표` 제목을 찾아 보충한다.
  `--settlement-map` JSON 으로 덮어쓸 수 있다::

      {"BS": ["공시용", "BS"], "IS": ["공시용", "IS"], "CE": ["CE"], "CF": ["CF"],
       "FN": ["FN"], "label_cols": {"CF": "BCD"}, "value_cols": {"CF": "FGHI"}}

  `label_cols`/`value_cols` 는 문자열 하나(모든 시트 공통)거나 `{시트: 열문자}` 다. 주지 않으면
  행마다 **첫 숫자 셀 앞의 문자 셀들**을 라벨로, **그 행의 숫자 셀 전부**를 금액으로 읽는다.

금액 비교
---------
정산표는 실수(천원 단위 소수)로 적히는 일이 많으므로 **보고서의 소수 자릿수로 반올림**해
견준다. `천원 ↔ 원`(×1000, ÷1000) 환산도 허용한다 (`amount_match`).

판정
----
`일치` / `불일치(금액)` / `오타 의심` / `라벨 미발견` / `미대사`.
`불일치(금액)` 가 하나라도 있으면 종료 코드 8 이고, `오타 의심`·`라벨 미발견` 은 종료 코드를
바꾸지 않는다 (계정과목 표기 차이는 대사 실패가 아니라 검토 거리다).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from .docmap import DocMap, load, norm_label, similarity
from .errors import FOOTING_MISMATCH_EXIT, InputError, OutputError
from .text import DEFAULT_PATTERN, classify, cr_to_lf, format_number

#: 재무제표 구분 → 보고서에서 그 구분을 여는 제목.
FS_TITLES = {
    "재무상태표": "BS",
    "손익계산서": "IS",
    "자본변동표": "CE",
    "현금흐름표": "CF",
}
#: 주석 구분.
FN = "FN"
#: 구분 → 정산표 시트 이름에 들어 있어야 할 조각 (기본 대응).
DEFAULT_MAP = {
    "BS": ["공시용", "BS"],
    "IS": ["공시용", "IS"],
    "CE": ["CE"],
    "CF": ["CF"],
    "FN": ["FN"],
}
#: 제목으로 인정할 최대 길이 (공백 제거 후) — 본문 문단이 제목으로 오인되지 않도록.
TITLE_MAX = 12

#: 판정.
MATCHED = "일치"
AMOUNT_MISMATCH = "불일치(금액)"
TYPO = "오타 의심"
LABEL_MISSING = "라벨 미발견"
UNRECONCILED = "미대사"
VERDICTS = (MATCHED, AMOUNT_MISMATCH, TYPO, LABEL_MISSING, UNRECONCILED)

#: 오타로 의심할 라벨 유사도 하한.
TYPO_RATIO = 0.6
#: 부분 일치(포함 관계)로 후보를 넓힐 때 요구하는 최소 라벨 길이.
MIN_PARTIAL = 3
#: 오타로 볼 최대 편집 거리 — 한 글자만 다르면 오타(`외상매줄금` ↔ `외상매출금`),
#: 두 글자 넘게 다르면 표기 차이로 보고 금액으로 대사한다(`배당금 지급액` ↔ `배당금의 지급`).
TYPO_DISTANCE = 1

_RE_WS = re.compile(r"[\s　]+")
#: `(1)` · `1.` 처럼 **번호만** 든 칸 — 정산표의 계정과목 번호다. `dsd_bridge.text.classify` 는
#: `(1)` 을 괄호 음수 `-1` 로 읽으므로, 금액으로 세지 않도록 여기서 걸러 낸다.
_RE_NUMBERING_ONLY = re.compile(r"^\(\s*\d{1,2}\s*\)$|^\d{1,2}\s*[.)]$")


def report_amount(value):
    """보고서 셀 하나 → 대사에 쓸 금액. `-`·빈칸·문자열은 `None`.

    `dsd_bridge.check.parse_amount` 와 달리 `-` 를 0 으로 보지 **않는다** — 정산표에는 그 칸이 아예
    없는 일이 흔해서, `-` 를 0 으로 세면 "정산표에 0 이 없다" 는 가짜 불일치가 생긴다.
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
    return k.value if k.numeric else None


def amount_match(a: Decimal, v: Decimal) -> bool:
    """보고서 금액 `a` 가 정산표 값 `v` 와 같은가 — `a` 의 소수 자릿수로 반올림해 견준다.

    정산표의 천원 표기(`1234.567`)와 보고서의 원 표기(`1,234,567`)가 섞이므로
    ×1000 · ÷1000 도 허용한다.
    """
    if a == v:
        return True
    q = a.as_tuple().exponent
    exp = Decimal(1).scaleb(q) if isinstance(q, int) and q < 0 else Decimal(1)
    for w in (v, v / 1000, v * 1000):
        try:
            if w.quantize(exp, rounding=ROUND_HALF_UP) == a:
                return True
        except Exception:                  # pragma: no cover - Decimal 범위 밖
            pass
    return False


# ─────────────────────────────────────────────────────────────────────────────
# 정산표 읽기
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class JsRow:
    """정산표 행 하나."""

    sheet: str
    row: int
    label: str
    labels: frozenset           # 정규화 라벨 후보 (칸마다 + 이어 붙인 것)
    values: tuple               # ((셀 좌표, Decimal), …)

    def amounts_ko(self) -> str:
        return ", ".join(f"{c}={format_number(v, DEFAULT_PATTERN)}" for c, v in self.values)


def _is_numbering(value) -> bool:
    """`(1)` · `1.` 처럼 계정과목 **번호만** 든 칸인가 (금액이 아니다)."""
    return isinstance(value, str) and bool(_RE_NUMBERING_ONLY.match(value.strip()))


def edit_distance(a: str, b: str, limit: int = 2) -> int:
    """두 문자열의 편집 거리 (`limit` 를 넘으면 `limit + 1` 로 끊는다)."""
    if abs(len(a) - len(b)) > limit:
        return limit + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        cur = [i]
        for j, cb in enumerate(b, start=1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        if min(cur) > limit:
            return limit + 1
        prev = cur
    return min(prev[-1], limit + 1)


def _letters(spec, sheet: str) -> str:
    """`label_cols`/`value_cols` 설정에서 이 시트의 열 문자를 꺼낸다. 없으면 빈 문자열."""
    if not spec:
        return ""
    if isinstance(spec, str):
        return spec.upper()
    if isinstance(spec, dict):
        return str(spec.get(sheet) or "").upper()
    return ""


def read_settlement_sheet(ws, *, label_cols: str = "", value_cols: str = "") -> list:
    """정산표 시트 하나에서 `JsRow` 목록을 읽는다.

    `label_cols`/`value_cols` 를 주지 않으면 행마다 **첫 숫자 셀 앞의 문자 셀들**을 라벨로,
    그 행의 숫자 셀 **전부**를 금액으로 본다. 라벨 칸이 `I.` · `1.` 처럼 번호만 있는 정산표
    (현금흐름표 시트 등)는 칸을 이어 붙여야 계정과목이 되므로, 칸마다의 라벨과 이어 붙인 라벨을
    **모두** 후보로 둔다.
    """
    from openpyxl.utils import get_column_letter

    out = []
    for row in ws.iter_rows():
        texts, values = [], []
        seen_number = False
        for c in row:
            letter = get_column_letter(c.column)
            v = c.value
            num = None if _is_numbering(v) else report_amount(v)
            if num is not None and (not value_cols or letter in value_cols):
                values.append((f"{letter}{c.row}", num))
                seen_number = True
                continue
            if num is not None:
                seen_number = True
                continue
            if isinstance(v, str) and v.strip():
                if label_cols:
                    if letter in label_cols:
                        texts.append(v.strip())
                elif not seen_number:
                    texts.append(v.strip())
        if not texts or not values:
            continue
        joined = " ".join(texts)
        labels = {norm_label(t) for t in texts} | {norm_label(joined)}
        labels.discard("")
        if not labels:
            continue
        out.append(JsRow(ws.title, row[0].row, joined, frozenset(labels), tuple(values)))
    return out


def _sheet_kinds(ws) -> set:
    """시트 **내용**에서 재무제표 제목을 찾아 구분을 추정한다 (이름으로 못 찾았을 때)."""
    found = set()
    for row in ws.iter_rows(max_row=min(ws.max_row, 40), values_only=True):
        for v in row:
            if not isinstance(v, str):
                continue
            name = _RE_WS.sub("", v)
            if len(name) > TITLE_MAX:
                continue
            for title, kind in FS_TITLES.items():
                if title in name:
                    found.add(kind)
    return found


def load_settlement(path, mapping: dict | None = None) -> dict:
    """정산표를 읽어 `{구분: [JsRow, …]}` 로. 시트 대응은 `mapping` (없으면 `DEFAULT_MAP`)."""
    from openpyxl import load_workbook

    p = Path(path)
    if not p.is_file():
        raise InputError(f"오류: 정산표를 찾을 수 없습니다: {p}")
    try:
        wb = load_workbook(p, data_only=True, read_only=False)
    except Exception as exc:
        raise InputError(
            f"오류: 정산표를 Excel 통합문서(.xlsx)로 열 수 없습니다: {p}\n"
            f"  사유: {type(exc).__name__}: {exc}"
        ) from exc
    conf = dict(mapping or {})
    label_spec = conf.pop("label_cols", None)
    value_spec = conf.pop("value_cols", None)
    names = {k: list(v) for k, v in (conf or DEFAULT_MAP).items()
             if k in DEFAULT_MAP or k in FS_TITLES.values() or k == FN}
    if not names:
        names = {k: list(v) for k, v in DEFAULT_MAP.items()}

    try:
        cache: dict = {}

        def rows_of(sheet_name: str) -> list:
            if sheet_name not in cache:
                cache[sheet_name] = read_settlement_sheet(
                    wb[sheet_name],
                    label_cols=_letters(label_spec, sheet_name),
                    value_cols=_letters(value_spec, sheet_name),
                )
            return cache[sheet_name]

        out: dict = {}
        for kind, needles in names.items():
            picked = [n for n in wb.sheetnames
                      if any(str(x).upper() in n.upper() for x in needles)]
            if not picked and kind in FS_TITLES.values():
                picked = [n for n in wb.sheetnames if kind in _sheet_kinds(wb[n])]
            rows = []
            for n in picked:
                rows.extend(rows_of(n))
            out[kind] = rows
        return out
    finally:
        wb.close()


# ─────────────────────────────────────────────────────────────────────────────
# 보고서 읽기
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class RepRow:
    """대사 대상이 되는 보고서 행 하나."""

    kind: str                   # BS / IS / CE / CF / FN
    sheet: str
    table: str
    row: int
    eid: str
    label: str
    norm: str
    amounts: tuple              # Decimal 목록 (`-`·빈칸은 뺀다)


def report_rows(doc: DocMap) -> list:
    """구조 지도에서 대사 대상 행을 모은다 (구분 판정은 모듈 docstring 참고)."""
    from .docmap import MapBlock, MapTable

    out = []
    for sheet in doc.sheets:
        note_sheet = "주석" in sheet.name
        if not note_sheet and not any(t.is_period for t in sheet.tables):
            continue
        kind = FN if note_sheet else ""
        for item in sheet.items:
            if isinstance(item, MapBlock):
                if not note_sheet:
                    kind = _title_kind(item.text) or kind
                continue
            if not isinstance(item, MapTable):      # pragma: no cover - 방어
                continue
            for r in item.rows:
                amounts = tuple(
                    a for a in (report_amount(c.value) for c in r.cells if c.col != 0)
                    if a is not None
                )
                if not note_sheet and not amounts:
                    k = _title_kind(r.label)
                    if k:
                        kind = k
                        continue
                if not kind or not r.label:
                    continue
                out.append(RepRow(kind, sheet.name, item.eid, r.excel_row, r.eid,
                                  r.label, norm_label(r.label), amounts))
    return out


def _title_kind(text: str) -> str:
    name = _RE_WS.sub("", cr_to_lf(str(text or "")).strip())
    if len(name) > TITLE_MAX:
        return ""
    for title, kind in FS_TITLES.items():
        if title in name:
            return kind
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# 대사
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Result:
    rep: RepRow
    verdict: str
    where: str = ""
    js_amounts: str = ""
    missing: tuple = ()
    suggestion: str = ""        # `오타 의심` 일 때 가장 비슷한 정산표 계정과목
    ratio: float = 0.0

    def line_ko(self) -> str:
        amts = ", ".join(format_number(a, DEFAULT_PATTERN) for a in self.rep.amounts)
        line = f"[{self.rep.kind}] {self.rep.label} | 보고서 {amts} | {self.verdict}"
        if self.where:
            line += f" | {self.where}"
        if self.suggestion:
            line += f" | 정산표에 비슷한 계정과목: '{self.suggestion}' (유사도 {self.ratio:.2f})"
        if self.missing:
            line += " | 정산표에 없는 금액: " + ", ".join(
                format_number(m, DEFAULT_PATTERN) for m in self.missing)
        return line

    def as_dict(self) -> dict:
        return {
            "kind": self.rep.kind, "sheet": self.rep.sheet, "table": self.rep.table,
            "row": self.rep.row, "id": self.rep.eid, "label": self.rep.label,
            "amounts": [str(a) for a in self.rep.amounts],
            "verdict": self.verdict, "where": self.where,
            "settlement_amounts": self.js_amounts,
            "missing": [str(m) for m in self.missing],
            "suggestion": self.suggestion, "ratio": round(self.ratio, 3),
        }


def _compare(amounts, js: JsRow):
    """보고서 금액 중 정산표 행에 없는 것."""
    vals = [v for _, v in js.values]
    return tuple(a for a in amounts if not any(amount_match(a, v) for v in vals))


def _candidates(rep: RepRow, rows: list) -> list:
    """`[(정산표 행 인덱스, JsRow), …]` — 정확히 같은 라벨이 있으면 그것만, 없으면 포함 관계."""
    exact = [(i, j) for i, j in enumerate(rows) if rep.norm in j.labels]
    if exact:
        return exact
    if len(rep.norm) < MIN_PARTIAL:
        return []
    return [(i, j) for i, j in enumerate(rows)
            if any(rep.norm in lab or lab in rep.norm for lab in j.labels if lab)]


def _closest(rep: RepRow, rows: list):
    """가장 비슷한 정산표 계정과목 `(표시 라벨, 유사도, 편집 거리)`."""
    best, ratio, dist = "", 0.0, TYPO_DISTANCE + 1
    for j in rows:
        for lab in j.labels:
            r = similarity(rep.norm, lab)
            if r > ratio:
                best, ratio = j.label, r
            dist = min(dist, edit_distance(rep.norm, lab, TYPO_DISTANCE))
    return best, ratio, dist


def _by_amount(rep: RepRow, rows: list, ptr: int) -> list:
    """계정과목 표기가 달라도 **금액이 모두 들어 있는** 정산표 행 `[(인덱스, JsRow)]`.

    예: 보고서 `배당금 지급액` ↔ 정산표 `(1) 배당금의 지급`. 오타(편집 거리 1)를
    먼저 걸러 낸 뒤에만 쓰므로, `외상매줄금` 같은 한 글자 오타가 여기서 조용히 넘어가지 않는다.
    """
    hits = [(i, j) for i, j in enumerate(rows) if not _compare(rep.amounts, j)]
    return sorted(hits, key=lambda t: (0 if t[0] >= ptr else 1, t[0]))


@dataclass
class SettlementReport:
    report_path: Path
    settlement_path: Path
    results: list = field(default_factory=list)
    #: 정산표에서 쓸 행을 못 찾은 구분 `[(구분, 사유)]`.
    notes: list = field(default_factory=list)

    def counts(self, kinds=None) -> dict:
        out = {v: 0 for v in VERDICTS}
        for r in self.results:
            if kinds is None or r.rep.kind in kinds:
                out[r.verdict] += 1
        return out

    @property
    def exit_code(self) -> int:
        return FOOTING_MISMATCH_EXIT if self.counts()[AMOUNT_MISMATCH] else 0

    def summary_ko(self) -> list:
        def line(name: str, kinds) -> str:
            c = self.counts(kinds)
            return (f"{name}: {MATCHED} {c[MATCHED]} / {AMOUNT_MISMATCH} {c[AMOUNT_MISMATCH]} / "
                    f"{TYPO} {c[TYPO]} / {LABEL_MISSING} {c[LABEL_MISSING]} / "
                    f"{UNRECONCILED} {c[UNRECONCILED]}")

        return [line("재무제표", {"BS", "IS", "CE", "CF"}), line("주석", {FN})]

    def lines_ko(self, verbose: bool = False) -> list:
        lines = [f"── 정산표 상호대사 ({self.settlement_path.name}) ──"]
        lines.extend(f"안내: {n}" for n in self.notes)
        for r in self.results:
            if r.verdict == MATCHED and not verbose:
                continue
            if r.verdict == UNRECONCILED and not verbose:
                continue
            lines.append(r.line_ko())
        lines.extend(self.summary_ko())
        return lines

    def json_obj(self) -> dict:
        return {
            "report": str(self.report_path),
            "settlement": str(self.settlement_path),
            "counts": {"재무제표": self.counts({"BS", "IS", "CE", "CF"}),
                       "주석": self.counts({FN})},
            "notes": list(self.notes),
            "items": [r.as_dict() for r in self.results],
        }

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n",
                         encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 대사 결과를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc

    def write_xlsx(self, path) -> None:
        """예시 스크립트와 같은 모양의 결과 통합문서 (판정 칸에 색)."""
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill

        fills = {MATCHED: "C6EFCE", AMOUNT_MISMATCH: "FFC7CE", TYPO: "FFEB9C",
                 LABEL_MISSING: "FFC7CE", UNRECONCILED: "D9D9D9"}
        header = ["구분", "시트", "행", "보고서 계정과목", "보고서 금액", "판정",
                  "정산표 위치", "정산표 금액", "정산표에 없는 금액", "비슷한 계정과목"]
        wb = Workbook()
        ws = wb.active
        ws.title = "재무제표"
        ws2 = wb.create_sheet("주석")
        for sheet in (ws, ws2):
            sheet.append(header)
            for c in sheet[1]:
                c.font = Font(bold=True)
            sheet.freeze_panes = "A2"
        for r in self.results:
            target = ws2 if r.rep.kind == FN else ws
            target.append([
                r.rep.kind, r.rep.sheet, r.rep.row or "", r.rep.label,
                ", ".join(format_number(a, DEFAULT_PATTERN) for a in r.rep.amounts),
                r.verdict, r.where, r.js_amounts,
                ", ".join(format_number(m, DEFAULT_PATTERN) for m in r.missing),
                f"{r.suggestion} ({r.ratio:.2f})" if r.suggestion else "",
            ])
            fill = fills.get(r.verdict)
            if fill:
                target.cell(target.max_row, 6).fill = PatternFill("solid", fgColor=fill)
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            wb.save(p)
        except OSError as exc:
            raise OutputError(f"오류: 대사 결과를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


def crosscheck(doc: DocMap, settlement_path, mapping: dict | None = None) -> SettlementReport:
    """구조 지도와 정산표를 상호대사한다."""
    js = load_settlement(settlement_path, mapping)
    report = SettlementReport(doc.path, Path(settlement_path))
    reps = report_rows(doc)
    for kind in sorted({r.kind for r in reps}):
        if not js.get(kind):
            report.notes.append(
                f"정산표에서 '{kind}' 에 해당하는 시트를 찾지 못했습니다 — "
                f"--settlement-map 으로 시트를 지정할 수 있습니다."
            )
    pointer: dict = {}
    for rep in reps:
        rows = js.get(rep.kind) or []
        if not rep.amounts:
            report.results.append(Result(rep, UNRECONCILED))
            continue
        ptr = pointer.get(rep.kind, 0)
        cands = _candidates(rep, rows)
        by_amount = False
        if not cands:
            best, ratio, dist = _closest(rep, rows)
            if dist <= TYPO_DISTANCE:
                # 한 글자 오타 — 금액이 맞더라도 계정과목을 고쳐야 하므로 여기서 멈춘다.
                report.results.append(Result(rep, TYPO, suggestion=best, ratio=ratio))
                continue
            cands = _by_amount(rep, rows, ptr)
            by_amount = bool(cands)
            if not cands:
                verdict = TYPO if ratio >= TYPO_RATIO else LABEL_MISSING
                report.results.append(Result(rep, verdict, suggestion=best if best else "",
                                             ratio=ratio))
                continue
        # 금액이 가장 많이 맞는 후보. 동률이면 앞서 쓴 행보다 **뒤에 있는** 것을 먼저 (표 순서).
        scored = []
        for i, j in cands:
            missing = _compare(rep.amounts, j)
            scored.append((-(len(rep.amounts) - len(missing)), 0 if i >= ptr else 1, i,
                           j, missing))
        scored.sort(key=lambda t: t[:3])
        _, _, idx, best_row, missing = scored[0]
        pointer[rep.kind] = idx + 1
        verdict = MATCHED if not missing else AMOUNT_MISMATCH
        where = f"{best_row.sheet}!{best_row.row}행 ({best_row.label[:24]})"
        if by_amount:
            where = "금액 대사 " + where
        report.results.append(Result(
            rep, verdict, where=where, js_amounts=best_row.amounts_ko(), missing=missing,
        ))
    n = sum(1 for r in report.results if r.where.startswith("금액 대사"))
    if n:
        report.notes.append(
            f"계정과목 표기가 달라 금액으로 대사한 행이 {n}개 있습니다 (-v 로 봅니다)."
        )
    return report


def load_mapping(path) -> dict:
    p = Path(path)
    if not p.is_file():
        raise InputError(f"오류: 대응표를 찾을 수 없습니다: {p}")
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise InputError(f"오류: 대응표를 읽을 수 없습니다: {p}\n  사유: {exc}") from exc
    if not isinstance(data, dict):
        raise InputError(f"오류: 대응표는 JSON 객체여야 합니다: {p}")
    return data


def check_settlement(report_path, settlement_path, mapping_path=None) -> SettlementReport:
    """`.dsd`/`.xlsx` 보고서 하나를 정산표와 상호대사한다 (CLI 진입점)."""
    mapping = load_mapping(mapping_path) if mapping_path else None
    return crosscheck(load(report_path), settlement_path, mapping)


__all__ = [
    "MATCHED",
    "AMOUNT_MISMATCH",
    "TYPO",
    "LABEL_MISSING",
    "UNRECONCILED",
    "VERDICTS",
    "DEFAULT_MAP",
    "FS_TITLES",
    "JsRow",
    "RepRow",
    "Result",
    "SettlementReport",
    "report_amount",
    "amount_match",
    "read_settlement_sheet",
    "load_settlement",
    "report_rows",
    "crosscheck",
    "check_settlement",
]
