"""`dsd-bridge template` — 전기 `.dsd` 가 없어 틀 문서로 시작할 때의 안전장치 (0.5.0).

가장 안전한 틀은 DART 편집기 서식으로 만든 빈 문서다. 이전 내용이 들어 있는 `.dsd` 를 틀로 쓴다면
가장 큰 위험은 **틀의 이전 내용이 결과에 남는 것**이다 — 이름, 숨은 회사 고유번호, 법인등록번호,
요약값, 금액, 날짜, 주석 문장. 엑셀에 보이지 않는 곳(표지 머리말, 숨은 속성, `meta.xml`)에도 있다.

세 가지 일을 한다.

1. `template 틀.dsd` — 틀의 식별값이 어디에 무엇이 있는지 보여 준다 (읽기 전용).
2. `template 틀.dsd --scrub -o 새틀.dsd [--replace 옛=새 …]` — 숨은 식별값을 비운 새 틀을 만든다.
   비우는 것: `COMPANY-NAME` 글과 `AREGCIK`, `EXTRACTION` 의 법인등록번호·주총일·요약값,
   `meta.xml` 의 `regcik`/`regname` 과 같은 이름의 요약 항목. 편집기가 저장할 때 요약값을 지우고
   검사 때 다시 만드는 것과 같은 상태가 된다. `--replace` 는 본문·표지 전체에서 글자를 바꾼다
   (표지 머리말은 엑셀에 나타나지 않아 여기서만 바꿀 수 있다).
3. `template 틀.dsd --check 결과.dsd|작업.xlsx [--old 문자열 …]` — 결과에 틀의 것이 남았는지 본다.
   - **식별값 잔존**(종료 코드 8): 틀에 있던 이름·고유번호·법인등록번호가 본문·표지·숨은 값·`meta.xml` 에 있다.
   - **틀 금액 잔존**(확인 필요): 틀에 있던 5자리 이상 금액과 같은 금액이 결과의 표 칸에 있다.
     위치와 무관하게 값으로 견주므로 행을 옮겨도 잡힌다. 우연히 같은 금액일 수 있어 판정은 하지 않는다.
   - **틀 글 잔존**(확인 필요): 틀과 글자가 같은 문단. 금액·날짜가 든 것은 목록으로, 나머지는 건수만
     (`-v` 로 전부). 표준 문안이면 그대로 두어도 되지만 회사 고유 내용이면 고쳐야 한다.
   - **숨은 날짜·요약값이 틀과 같음**(`.dsd` 결과만).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from .container import DsdContainer, meta_file_info_size, update_meta_size
from .diff import RE_NUMBER_CLUE
from .docmap import MapBlock, MapTable, load
from .errors import FOOTING_MISMATCH_EXIT, InputError, OutputError
from .text import DEFAULT_PATTERN, escape_text, format_number, unescape
from .validate import check_wellformed, xsd_validate
from .xmlsplice import PatchSet, SourceIndex, attr_value_patch, set_text_patch

#: 틀의 이전 값이라 비워야 하는 `EXTRACTION`/`meta.xml ITEM` 이름 — 법인등록번호, 주총일, 요약값, 사업연도.
SCRUB_ITEMS = ("CRP_RGS_NO", "GMSH_DATE", "TOT_ASSETS", "TOT_DEBTS", "TOT_SALES", "TOT_EMPL",
               "BSN_BGN_DT", "BSN_END_DT")
#: 그대로 남으면 이전 문서의 식별값으로 읽히는 항목.
IDENTITY_ITEMS = ("CRP_RGS_NO",)
#: 날짜 속성 (import 의 `DATE_AUNITS` 와 같다).
DATE_AUNITS = frozenset({"PERIODFROM", "PERIODTO", "PERIODFROM2", "PERIODTO2",
                         "SUB_PERIODFROM", "SUB_PERIODTO", "COMM_DATE"})
#: 이 값 이상의 금액만 "틀 금액" 으로 견준다 (작은 수는 우연히 겹친다).
MIN_AMOUNT = 10000
#: 이 길이(공백 제외) 이상의 문단만 "틀 글" 로 견준다.
MIN_TEXT = 15
#: 기본 출력에서 보여 줄 항목 수.
SHOW = 40

_RE_WS = re.compile(r"[\s　]+")
_RE_TAG = re.compile(r"<[^>]+>")
_RE_CORP = re.compile(r"주식회사|유한책임회사|유한회사|\(주\)|㈜|\(유\)")


def _squash(text) -> str:
    return _RE_WS.sub("", "" if text is None else str(text))


def _read(path):
    p = Path(path)
    if p.suffix.lower() != ".dsd":
        raise InputError(f"오류: 틀은 .dsd 파일이어야 합니다: {p}")
    if not p.is_file():
        raise InputError(f"오류: 파일을 찾을 수 없습니다: {p}")
    container = DsdContainer.read(p)
    raw = container.contents_bytes
    return container, raw, SourceIndex.build(raw)


@dataclass
class Identity:
    """틀 `.dsd` 의 식별값."""

    path: Path
    company: str = ""
    cik: str = ""
    account_type: str = ""
    extraction: dict = field(default_factory=dict)      # EXTRACTION ACODE → 글
    meta: dict = field(default_factory=dict)            # regcik / regname / ITEM 이름 → 값
    dates: dict = field(default_factory=dict)           # AUNIT → [AUNITVALUE…]
    schema: str = ""
    editver: str = ""
    cover: list = field(default_factory=list)           # 표지 머리말(엑셀에 안 보임)의 글

    def needles(self, extra=()) -> list:
        """결과에서 찾을 문자열 `[(이름, 공백 없는 문자열)]` — 회사 이름·핵심 이름·고유번호·법인등록번호."""
        out = []

        def add(name, value):
            v = _squash(value)
            if len(v) >= 3 and all(v != x for _, x in out):
                out.append((name, v))

        for name in (self.company, self.meta.get("regname", "")):
            add("회사 이름", name)
            add("회사 이름(법인 형태 제외)", _RE_CORP.sub("", name or ""))
        for name, value in (("회사 고유번호", self.cik), ("회사 고유번호", self.meta.get("regcik", ""))):
            if len(_squash(value)) >= 6:
                add(name, value)
        reg = _squash(self.extraction.get("CRP_RGS_NO") or self.meta.get("CRP_RGS_NO") or "")
        if len(reg) == 13 and reg.isdigit():
            add("법인등록번호", reg)
            add("법인등록번호", f"{reg[:6]}-{reg[6:]}")
        for value in extra:
            add("지정한 문자열", value)
        return out

    def lines_ko(self) -> list:
        lines = [f"── 틀의 식별값 ({self.path.name}) ──",
                 f"서식: {self.schema or '?'} / 편집기 {self.editver or '?'}"]
        lines.append(f"회사 이름(COMPANY-NAME): {self.company or '(빈 값)'} / 고유번호(AREGCIK): {self.cik or '(빈 값)'}")
        if self.meta.get("regname") or self.meta.get("regcik"):
            lines.append(f"meta.xml: regname {self.meta.get('regname') or '(빈 값)'} / "
                         f"regcik {self.meta.get('regcik') or '(빈 값)'}")
        filled = {k: v for k, v in self.extraction.items() if v}
        if filled:
            lines.append("숨은 요약·식별 항목(EXTRACTION): "
                         + ", ".join(f"{k}={v}" for k, v in filled.items()))
        if self.dates:
            lines.append("숨은 날짜 값(AUNITVALUE): "
                         + ", ".join(f"{k}={'/'.join(v)}" for k, v in self.dates.items()))
        if self.cover:
            lines.append("표지 머리말(엑셀에 나타나지 않음, --replace 나 XML 치환으로만 바뀜): "
                         + " | ".join(t[:30] for t in self.cover[:8]))
        if not (self.company or self.cik or filled):
            lines.append("안내: 숨은 식별값이 비어 있는 파일입니다(등록 전 초안이거나 이미 --scrub 한 틀).")
        if self.schema and self.schema != "dart4":
            lines.append(f"안내: 서식이 {self.schema} 입니다. 지금 DART 편집기 서식(dart4)과 달라 편집기에서 열 때 "
                         "변환·정정 안내가 나올 수 있으니, 가능하면 최신 서식으로 만든 파일을 틀로 쓰세요.")
        return lines


def _preamble_end(index: SourceIndex) -> int:
    """첫 `ATOC="Y"` 제목 앞까지가 머리말(표지 포함) — 엑셀 시트에 나타나지 않는다."""
    for sp in index.spans:
        if sp.attrs.get("ATOC") == "Y":
            return sp.start_lt
    return len(index.raw)


def identity(path) -> Identity:
    container, raw, index = _read(path)
    out = Identity(Path(path), schema=container.schema() or "", editver=container.editver() or "")
    end = _preamble_end(index)
    for sp in index.spans:
        if sp.tag == "COMPANY-NAME":
            out.company = sp.text.strip()
            out.cik = sp.attrs.get("AREGCIK", "")
            out.account_type = sp.attrs.get("AACCOUNTTYPE", "")
        elif sp.tag == "EXTRACTION" and sp.attrs.get("ACODE"):
            out.extraction[sp.attrs["ACODE"]] = sp.text.strip()
        aunit = sp.attrs.get("AUNIT")
        if aunit in DATE_AUNITS and sp.attrs.get("AUNITVALUE"):
            out.dates.setdefault(aunit, [])
            if sp.attrs["AUNITVALUE"] not in out.dates[aunit]:
                out.dates[aunit].append(sp.attrs["AUNITVALUE"])
        if sp.start_lt < end and sp.tag in ("P", "TD", "TU", "TE", "TH") and sp.text.strip():
            if any(index[a].tag == "COVER" for a in _ancestors(index, sp.ordinal)):
                out.cover.append(" ".join(sp.text.split()))
    out.meta = _meta_values(container.meta_bytes)
    return out


def _ancestors(index: SourceIndex, ordinal: int):
    po = index[ordinal].parent_ordinal
    while po >= 0:
        yield po
        po = index[po].parent_ordinal


def _meta_values(meta: bytes) -> dict:
    s = meta.decode("utf-8", "replace")
    out = {}
    for key in ("regcik", "regname"):
        m = re.search(key + r'="([^"]*)"', s)
        if m:
            out[key] = unescape(m.group(1))
    for m in re.finditer(r'<ITEM\b[^>]*\bname="([^"]*)"[^>]*\bvalue="([^"]*)"', s):
        out[m.group(1)] = unescape(m.group(2))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# --scrub
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class ScrubReport:
    source: Path
    output: Path
    cleared: list = field(default_factory=list)
    replaced: list = field(default_factory=list)        # (옛, 새, 건수)
    xsd: str = ""

    def lines_ko(self) -> list:
        lines = [f"틀 정리 완료: {self.source.name} → {self.output}"]
        if self.cleared:
            lines.append("비운 숨은 값: " + ", ".join(self.cleared))
        else:
            lines.append("안내: 비울 숨은 식별값이 없었습니다.")
        for old, new, n in self.replaced:
            lines.append(f"바꾼 글: '{old}' → '{new}' {n}곳")
            if not n:
                lines.append(f"경고: '{old}' 을(를) 찾지 못했습니다. 글자 사이에 공백이나 서식이 끼어 있으면 더 짧게 지정하세요.")
        lines.append(f"XSD 검증: {self.xsd}")
        lines.append("다음: 이 파일을 export 해 새 회사 내용으로 채우고, 되돌린 뒤 "
                     "`dsd-bridge template <원래 틀.dsd> --check <결과.dsd>` 로 남은 것을 확인하세요. "
                     "회사 지정(제출인)은 DART 편집기에서 다시 해야 합니다.")
        return lines


def scrub(path, out, replacements=()) -> ScrubReport:
    """숨은 식별값을 비우고(선택) 글자를 바꾼 새 틀을 쓴다. 입력은 건드리지 않는다."""
    src, dst = Path(path), Path(out) if out else None
    if dst is None:
        raise InputError("오류: --scrub 에는 출력 경로(-o)가 필요합니다. 원본 틀은 덮어쓰지 않습니다.")
    if dst.resolve() == src.resolve():
        raise OutputError("오류: 출력 경로가 틀과 같습니다. 다른 이름으로 지정하세요.")
    container, raw, index = _read(src)
    report = ScrubReport(src, dst)
    patches = PatchSet()
    for sp in index.spans:
        if sp.tag == "COMPANY-NAME":
            if sp.text.strip() and not sp.self_closing:
                patches.add(set_text_patch(index, sp.ordinal, b""))
                report.cleared.append("COMPANY-NAME")
            if sp.attrs.get("AREGCIK"):
                patches.add(attr_value_patch(index, sp.ordinal, "AREGCIK", "")[0])
                report.cleared.append("AREGCIK")
        elif (sp.tag == "EXTRACTION" and sp.attrs.get("ACODE") in SCRUB_ITEMS
              and sp.text.strip() and not sp.self_closing):
            patches.add(set_text_patch(index, sp.ordinal, b""))
            report.cleared.append(sp.attrs["ACODE"])
    result = patches.apply(raw)
    text = result.decode("utf-8")
    for old, new in replacements:
        if not old:
            raise InputError("오류: --replace 의 옛 글이 비어 있습니다.")
        a, b = escape_text(old), escape_text(new)
        n = text.count(a)
        text = text.replace(a, b)
        report.replaced.append((old, new, n))
    result = text.encode("utf-8")
    check_wellformed(result)

    meta = container.meta_bytes.decode("utf-8")
    for key in ("regcik", "regname"):
        meta, n = re.subn(r'(\b' + key + r'=")[^"]+(")', r"\1\2", meta)
        if n:
            report.cleared.append(f"meta.{key}")
    for name in SCRUB_ITEMS:
        meta, n = re.subn(r'(<ITEM\b[^>]*\bname="' + name + r'"[^>]*\bvalue=")[^"]+(")', r"\1\2", meta)
        if n:
            report.cleared.append(f"meta.{name}")
    meta_bytes = meta.encode("utf-8")
    if len(result) != len(raw) and (meta_file_info_size(meta_bytes) or "") != "":
        meta_bytes = update_meta_size(meta_bytes, len(result))
    v = xsd_validate(result, container.schema())
    report.xsd = v.status_ko + (f" ({v.reason})" if v.status != "pass" and v.reason else "")
    try:
        container.write(dst, contents=result, meta=meta_bytes)
    except OSError as exc:
        raise OutputError(f"오류: 출력 파일을 쓸 수 없습니다: {dst}\n  사유: {exc}") from exc
    return report


# ─────────────────────────────────────────────────────────────────────────────
# --check
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class TemplateCheck:
    template: Path
    result: Path
    needles: list = field(default_factory=list)
    #: 식별값 잔존 `[(이름, 문자열, 어디, 건수)]`.
    identity_hits: list = field(default_factory=list)
    #: 틀 금액 잔존 `[(시트, 행, 과목, 열, 금액)]`.
    amount_hits: list = field(default_factory=list)
    #: 틀 글 잔존 — 금액·날짜가 든 것 / 그 밖 `[(시트, 행, 단서, 앞부분)]`.
    text_hits: list = field(default_factory=list)
    plain_hits: list = field(default_factory=list)
    #: 숨은 값이 틀과 같음 `[설명]`.
    hidden_hits: list = field(default_factory=list)
    notes: list = field(default_factory=list)

    @property
    def exit_code(self) -> int:
        return FOOTING_MISMATCH_EXIT if self.identity_hits else 0

    def summary_ko(self) -> str:
        return (f"틀 잔존 검사: 식별값 잔존 {len(self.identity_hits)} / 숨은 값이 틀과 같음 {len(self.hidden_hits)} / "
                f"틀 금액 잔존 {len(self.amount_hits)} / 금액·날짜가 든 틀 글 {len(self.text_hits)} / "
                f"그 밖 틀과 같은 글 {len(self.plain_hits)}")

    def lines_ko(self, verbose: bool = False) -> list:
        lines = [f"── 틀 잔존 검사 (틀 {self.template.name} → 결과 {self.result.name}) ──"]
        lines.extend(f"안내: {n}" for n in self.notes)
        if self.needles:
            lines.append("찾은 문자열: " + ", ".join(f"{name} '{v}'" for name, v in self.needles))
        for name, value, where, n in self.identity_hits:
            lines.append(f"[식별값 잔존] {name} '{value}' — {where} {n}곳")
        for h in self.hidden_hits:
            lines.append(f"[숨은 값이 틀과 같음] {h}")
        if self.amount_hits:
            lines.append(f"── 틀 금액 잔존 {len(self.amount_hits)}건 (틀에 있던 금액과 같은 금액, 우연인지 확인) ──")
            shown = self.amount_hits if verbose else self.amount_hits[:SHOW]
            for sheet, row, label, col, amount in shown:
                where = f"{sheet} {row}행 {col}열" if row else f"{sheet} '{label[:16]}'"
                lines.append(f"[틀 금액] {where} {label[:20]} | {format_number(amount, DEFAULT_PATTERN)}")
            if len(shown) < len(self.amount_hits):
                lines.append(f"… 그 밖 {len(self.amount_hits) - len(shown)}건 (-v 로 전부 봅니다)")
        if self.text_hits:
            lines.append(f"── 금액·날짜가 든 채 틀과 같은 글 {len(self.text_hits)}건 ──")
            shown = self.text_hits if verbose else self.text_hits[:SHOW]
            for sheet, row, clues, snippet in shown:
                where = f"{sheet} {row}행" if row else sheet
                lines.append(f"[틀 글] {where} — {', '.join(clues[:5])} | {snippet}")
            if len(shown) < len(self.text_hits):
                lines.append(f"… 그 밖 {len(self.text_hits) - len(shown)}건 (-v 로 전부 봅니다)")
        if self.plain_hits and verbose:
            lines.append(f"── 그 밖 틀과 같은 글 {len(self.plain_hits)}건 (표준 문안이면 그대로, 회사 고유 내용이면 수정) ──")
            for sheet, row, _clues, snippet in self.plain_hits:
                lines.append(f"[틀 글·일반] {sheet} {row}행 | {snippet}" if row else f"[틀 글·일반] {sheet} | {snippet}")
        lines.append(self.summary_ko())
        return lines

    def json_obj(self) -> dict:
        def rows(items):
            return [{"sheet": s, "row": r, "clues": list(c), "snippet": t} for s, r, c, t in items]

        return {
            "template": str(self.template), "result": str(self.result),
            "identity": [{"name": n, "value": v, "where": w, "count": c} for n, v, w, c in self.identity_hits],
            "hidden": list(self.hidden_hits),
            "amounts": [{"sheet": s, "row": r, "label": lab, "col": c, "amount": str(a)}
                        for s, r, lab, c, a in self.amount_hits],
            "texts": rows(self.text_hits), "plain_texts": rows(self.plain_hits),
            "notes": list(self.notes),
        }

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 검사 결과를 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


def _texts_and_amounts(doc):
    """구조 지도에서 (문단 글 목록, 표 금액 칸 목록, 표 글 목록)."""
    blocks, amounts, cells = [], [], []
    for sheet in doc.sheets:
        for item in sheet.items:
            if isinstance(item, MapBlock):
                blocks.append((sheet.name, item.excel_row, str(item.text or "")))
            elif isinstance(item, MapTable):
                for r in tuple(item.headers) + tuple(item.rows):
                    for c in r.cells:
                        a = c.amount
                        if a is not None:
                            amounts.append((sheet.name, r.excel_row, r.label, c.letter, a))
                        elif c.text.strip():
                            cells.append((sheet.name, r.excel_row, c.text))
    return blocks, amounts, cells


def check_against(template, result, old=()) -> TemplateCheck:
    """결과(`.dsd` 또는 작업 `.xlsx`)에 틀의 것이 남았는지 본다."""
    ident = identity(template)
    rp = Path(result)
    if not rp.is_file():
        raise InputError(f"오류: 파일을 찾을 수 없습니다: {rp}")
    out = TemplateCheck(Path(template), rp)
    out.needles = ident.needles(old)

    t_blocks, t_amounts, _ = _texts_and_amounts(load(template))
    r_doc = load(rp)
    r_blocks, r_amounts, r_cells = _texts_and_amounts(r_doc)

    # 1) 식별값 — 본문(시트)
    body = [(s, row, _squash(t)) for s, row, t in r_blocks] + [(s, row, _squash(t)) for s, row, t in r_cells]
    for name, needle in out.needles:
        places = [f"{s} {row}행" if row else s for s, row, t in body if needle in t]
        if places:
            uniq = list(dict.fromkeys(places))
            where = "본문(" + ", ".join(uniq[:6]) + (f" 외 {len(uniq) - 6}곳" if len(uniq) > 6 else "") + ")"
            out.identity_hits.append((name, needle, where, len(places)))

    if rp.suffix.lower() == ".dsd":
        container, raw, index = _read(rp)
        end = _preamble_end(index)
        pre = _squash(unescape(raw[:end].decode("utf-8", "replace")))
        meta = _squash(unescape(container.meta_bytes.decode("utf-8", "replace")))
        for name, needle in out.needles:
            n = pre.count(needle)
            if n:
                out.identity_hits.append((name, needle, "표지 머리말·숨은 값(엑셀에 안 보임)", n))
            n = meta.count(needle)
            if n:
                out.identity_hits.append((name, needle, "meta.xml", n))
        # 2) 숨은 값이 틀과 같음
        r_ident = identity(rp)
        same = []
        for k in SCRUB_ITEMS:
            tv = {v for v in (ident.extraction.get(k), ident.meta.get(k)) if v}
            rv = {v for v in (r_ident.extraction.get(k), r_ident.meta.get(k)) if v}
            if tv & rv:
                same.append(k)
        if same:
            out.hidden_hits.append("요약·식별 항목 " + ", ".join(same)
                                   + " — 틀 값 그대로입니다(--scrub 로 비우거나 DART 편집기에서 재추출)")
        template_dates = {v for vals in ident.dates.values() for v in vals}
        stale = [f"{k}={v}" for k, vals in r_ident.dates.items() for v in vals if v in template_dates]
        if stale:
            out.hidden_hits.append("숨은 날짜 값 " + ", ".join(stale)
                                   + " — 틀과 같은 날짜입니다(그 칸의 날짜 글을 고치면 같이 바뀝니다)")
    else:
        out.notes.append("작업 엑셀은 시트에 보이는 것만 검사합니다. 표지 머리말·숨은 값·meta.xml 은 "
                         "되돌린 .dsd 로 다시 --check 하세요.")

    # 3) 틀 금액 — 위치와 무관하게 값으로
    template_amounts = {a for _, _, _, _, a in t_amounts if abs(a) >= MIN_AMOUNT}
    out.amount_hits = [h for h in r_amounts if h[4] in template_amounts]

    # 4) 틀 글
    template_texts = {_squash(t) for _, _, t in t_blocks}
    for sheet, row, text in r_blocks:
        key = _squash(text)
        if len(key) < MIN_TEXT or key not in template_texts:
            continue
        flat = " ".join(text.split())
        clues = tuple(dict.fromkeys(m.group(0).strip() for m in RE_NUMBER_CLUE.finditer(flat)))
        (out.text_hits if clues else out.plain_hits).append((sheet, row, clues, flat[:60]))
    return out



# ─────────────────────────────────────────────────────────────────────────────
# --blank (0.7.0)
# ─────────────────────────────────────────────────────────────────────────────


#: --blank 이 이름 자리에 넣는 표식 (review 의 자리표시 ○ 로 잡힌다)
BLANK_NAME = "○회사명○"


@dataclass
class BlankReport:
    source: Path
    output: Path
    amounts: int = 0
    paragraphs: int = 0
    names: int = 0
    kept_mixed: list = field(default_factory=list)
    scrubbed: list = field(default_factory=list)
    xsd: str = ""

    def lines_ko(self) -> list:
        lines = [f"틀 비우기 완료: {self.source.name} → {self.output}",
                 f"비운 금액 칸 {self.amounts}개 / 비운 주석 문단 {self.paragraphs}개(번호 제목 줄은 남김) / "
                 f"이름을 {BLANK_NAME} 로 바꾼 곳 {self.names}곳",
                 "비운 숨은 값: " + (", ".join(self.scrubbed) or "없음")]
        if self.kept_mixed:
            lines.append(f"안내: 서식 조각이 섞인 주석 문단 {len(self.kept_mixed)}개는 그대로 두었습니다(엑셀에서 글자만 "
                         "고치면 됩니다): " + ", ".join(self.kept_mixed[:6]))
        lines.append(f"XSD 검증: {self.xsd}")
        lines.append("다음: 틀 모음 폴더(~/.dsd-bridge/templates/)에 보관해 두고, 쓸 때는 export 해 채운 뒤 "
                     "`template <원래 파일> --check <결과>` 로 남은 이전 내용을 확인하세요.")
        return lines


def blank(path, out) -> BlankReport:
    """구조만 남긴 빈 틀: 식별값을 비우고, 표 금액을 지우고, 주석 문단은 비우되 번호 제목 줄은 남기고, 이름은 ○회사명○ 로(자리표시로 잡히고, --scrub --replace "○회사명○=새 이름" 으로 바꾼다)."""
    import tempfile

    from .model import DocModel, note_heading
    from .text import CR_TOKEN

    src, dst = Path(path), Path(out) if out else None
    if dst is None:
        raise InputError("오류: --blank 에는 출력 경로(-o)가 필요합니다. 원본은 덮어쓰지 않습니다.")
    if dst.resolve() == src.resolve():
        raise OutputError("오류: 출력 경로가 원본과 같습니다. 다른 이름으로 지정하세요.")
    ident = identity(src)
    names = [n for n in {ident.company, ident.meta.get("regname", ""),
                         _RE_CORP.sub("", ident.company or "").strip()} if n and len(n) >= 2]
    with tempfile.TemporaryDirectory() as tmp:
        mid = Path(tmp) / "scrubbed.dsd"
        sr = scrub(src, mid, [(n, BLANK_NAME) for n in sorted(names, key=len, reverse=True)])
        container, raw, index = _read(mid)
    report = BlankReport(src, dst, names=sum(n for _, _, n in sr.replaced), scrubbed=sr.cleared)
    model = DocModel.build(index, raw)
    patches = PatchSet()
    for section in model.sections:
        in_notes = "주석" in section.sheet_name
        for block in section.blocks:
            if block.table is not None:
                for r in block.table.rows:
                    for c in r.cells:
                        if c.numeric and not c.mixed:
                            patches.add(set_text_patch(index, c.ordinal, b""))
                            report.amounts += 1
                continue
            if not in_notes or block.kind != "P":
                continue
            parts = block.text.split(CR_TOKEN)
            heading = note_heading(parts[0]) is not None
            keep = parts[0] if heading else ""
            if block.text == keep or not block.text.strip():
                continue
            if block.mixed:
                report.kept_mixed.append(" ".join(block.text.replace(CR_TOKEN, " ").split())[:20])
                continue
            patches.add(set_text_patch(index, block.ordinal, escape_text(keep).encode("utf-8")))
            report.paragraphs += 1
    result = patches.apply(raw)
    check_wellformed(result)
    meta = container.meta_bytes
    if len(result) != len(raw) and (meta_file_info_size(meta) or "") != "":
        meta = update_meta_size(meta, len(result))
    v = xsd_validate(result, container.schema())
    report.xsd = v.status_ko + (f" ({v.reason})" if v.status != "pass" and v.reason else "")
    try:
        container.write(dst, contents=result, meta=meta)
    except OSError as exc:
        raise OutputError(f"오류: 출력 파일을 쓸 수 없습니다: {dst}\n  사유: {exc}") from exc
    return report


__all__ = ["Identity", "ScrubReport", "TemplateCheck", "identity", "scrub", "check_against",
           "SCRUB_ITEMS", "MIN_AMOUNT", "BlankReport", "blank", "BLANK_NAME"]
