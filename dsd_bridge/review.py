"""확인 필요 목록 — 자동 판정(OK/불일치)은 아니지만 사람이 꼭 봐야 하는 단서 (0.4.0).

`dsd-bridge check` 가 합계검증 뒤에 덧붙인다. `.dsd` 와 작업 엑셀(`.xlsx`) 모두에서 동작한다
(구조 지도 `docmap` 위에서 돈다). 종료 코드에는 영향을 주지 않는다.

1. **주석 참조 번호** — 재무제표 과목의 `(주석4,7)` 이 가리키는 번호가 주석 목록에 없다
   (주석을 지우거나 끼워 번호가 밀렸는데 참조를 안 고친 경우).
2. **주석 번호 순서** — 주석 번호가 건너뛰거나 겹친다 (`18.` 다음이 `20.`).
3. **중복 행** — 표준 재무제표 표에서 계정과목과 금액이 모두 같은 행이 두 번 나온다
   (행을 끼우다 같은 줄이 두 번 들어간 경우).
4. **과목 없는 금액 행** — 금액은 있는데 계정과목 칸이 비었다 (새 행에 라벨을 안 채운 경우).
5. **자리표시** — `○`, `XXXX년` 처럼 아직 채우지 않은 자리표시가 남아 있다.
"""

from __future__ import annotations

import re

from .docmap import DocMap, MapBlock, MapTable
from .model import note_heading
from .text import cr_to_lf

#: `(주석4,7)` · `<주석 3>` — 숫자가 있는 주석 참조.
_RE_NOTE_REF = re.compile(r"[<\(（]\s*주\s*석\s*([\d,\s]+)[>\)）]")
_RE_PLACEHOLDER = re.compile(r"○+|X{2,}\s*년|X{2}\s*월|X{2}\s*일")
#: 한 항목에 늘어놓을 위치 수.
SHOW = 8


def _is_note_sheet(name: str) -> bool:
    return "주석" in name


def _where(sheet: str, row: int) -> str:
    return f"{sheet} {row}행" if row else sheet


def note_duplicates(doc: DocMap) -> list:
    """바로 앞 주석과 번호가 같은 제목 (주석을 복제한 뒤 번호를 안 고친 경우) `[번호]`."""
    out = []
    last = 0
    for sheet in doc.sheets:
        if not _is_note_sheet(sheet.name):
            continue
        for item in sheet.items:
            if not isinstance(item, MapBlock):
                continue
            h = note_heading(item.text or "")
            if h is None:
                continue
            if h[0] == last:
                out.append(h[0])
            elif h[0] > last:
                last = h[0]
    return out


def note_numbers(doc: DocMap) -> list:
    """주석 번호 제목을 문서 순서로 — `dsd_bridge.model` 과 같은 규칙(첫 줄 `N. 제목`, 번호는 커져야 함)."""
    out = []
    last = 0
    for sheet in doc.sheets:
        if not _is_note_sheet(sheet.name):
            continue
        for item in sheet.items:
            if not isinstance(item, MapBlock):
                continue
            h = note_heading(item.text or "")
            if h is not None and h[0] > last:
                last = h[0]
                out.append((h[0], h[1], sheet.name, item.excel_row))
    return out


def note_refs(doc: DocMap) -> list:
    """재무제표(주석 시트가 아닌 기간표)의 과목이 가리키는 주석 번호 `[(번호, 시트, 행, 라벨)]`."""
    out = []
    for sheet in doc.sheets:
        if _is_note_sheet(sheet.name):
            continue
        for t in sheet.tables:
            if not t.is_period:
                continue
            for r in t.rows:
                raw = ""
                for c in r.cells:
                    if c.col == 0:
                        raw = cr_to_lf(c.text)
                        break
                for m in _RE_NOTE_REF.finditer(raw):
                    for tok in re.split(r"[,\s]+", m.group(1).strip()):
                        if tok.isdigit():
                            out.append((int(tok), sheet.name, r.excel_row, r.label))
    return out


def review(doc: DocMap) -> list:
    """확인 필요 문장 목록 (없으면 빈 목록)."""
    out = []

    notes = note_numbers(doc)
    numbers = [n for n, _, _, _ in notes]
    if numbers:
        known = set(numbers)
        missing = [(n, s, row, label) for n, s, row, label in note_refs(doc) if n not in known]
        if missing:
            shown = ", ".join(f"{label[:20]}→주석 {n}" for n, _, _, label in missing[:SHOW])
            more = f" 외 {len(missing) - SHOW}곳" if len(missing) > SHOW else ""
            out.append(
                f"주석 참조: 재무제표 과목이 가리키는 주석 번호가 주석 목록({numbers[0]}~{numbers[-1]}, "
                f"{len(numbers)}개)에 없습니다 — {shown}{more}. 주석 번호가 바뀌었으면 참조도 고치세요."
            )
        gaps = [f"{a} 다음이 {b}" for a, b in zip(numbers, numbers[1:]) if b != a + 1]
        gaps += [f"{n} 이(가) 두 번" for n in note_duplicates(doc)]
        if gaps:
            out.append("주석 번호: 번호가 이어지지 않습니다 — " + ", ".join(gaps[:SHOW])
                       + ". 주석을 지우거나 끼웠으면 뒤 번호와 재무제표의 주석 참조를 함께 고치세요.")

    dups, blanks = [], []
    for sheet in doc.sheets:
        if _is_note_sheet(sheet.name):
            continue
        for t in sheet.tables:
            if not isinstance(t, MapTable) or not t.is_period:
                continue
            seen: dict = {}
            for r in t.rows:
                vals = r.vals
                if not vals or not any(v != 0 for v in vals.values()):
                    continue
                if not r.norm:
                    blanks.append(_where(sheet.name, r.excel_row))
                    continue
                key = (r.norm, tuple(sorted(vals.items())))
                if key in seen:
                    dups.append(f"{_where(sheet.name, r.excel_row)} '{r.label[:20]}'"
                                f"(앞 {seen[key]}행과 같음)" if r.excel_row else f"{sheet.name} '{r.label[:20]}'")
                else:
                    seen[key] = r.excel_row
    if dups:
        more = f" 외 {len(dups) - SHOW}곳" if len(dups) > SHOW else ""
        out.append("중복 행: 계정과목과 금액이 모두 같은 행이 한 표에 두 번 있습니다 — "
                   + ", ".join(dups[:SHOW]) + more + ". 같은 줄이 두 번 들어갔는지 확인하세요.")
    if blanks:
        more = f" 외 {len(blanks) - SHOW}곳" if len(blanks) > SHOW else ""
        out.append("과목 없는 금액 행: 금액은 있는데 계정과목이 비어 있습니다 — "
                   + ", ".join(blanks[:SHOW]) + more + ".")

    places = []
    count = 0
    for sheet in doc.sheets:
        for item in sheet.items:
            if isinstance(item, MapBlock):
                n = len(_RE_PLACEHOLDER.findall(item.text or ""))
                if n:
                    count += n
                    places.append(_where(sheet.name, item.excel_row))
            elif isinstance(item, MapTable):
                for r in tuple(item.headers) + tuple(item.rows):
                    n = sum(len(_RE_PLACEHOLDER.findall(c.text)) for c in r.cells)
                    if n:
                        count += n
                        places.append(_where(sheet.name, r.excel_row))
    if count:
        uniq = list(dict.fromkeys(places))
        more = f" 외 {len(uniq) - SHOW}곳" if len(uniq) > SHOW else ""
        out.append(f"자리표시: 아직 채우지 않은 자리표시(○, XXXX년 등)가 {count}개 있습니다 — "
                   + ", ".join(uniq[:SHOW]) + more + ". 제출 전에 채우세요.")
    return out


__all__ = ["review", "note_numbers", "note_refs", "note_duplicates"]
