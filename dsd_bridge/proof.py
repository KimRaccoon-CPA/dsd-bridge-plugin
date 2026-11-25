"""`dsd-bridge proof` — 맞춤법·띄어쓰기 점검 도우미 (0.6.0).

두 가지를 한다.

1. **기계 점검**: 규칙으로 확실히 잡히는 표기 실수를 위치와 함께 낸다 — 같은 낱말 연속(`의 의`),
   괄호 짝, 낱자모(`ㅇ`), 문장부호 앞 공백, 마침표 두 번, 로마숫자 섞임(`Vl.`, `Ⅳ I`), 콤마 자리,
   `원원`. 자동으로 고치지 않는다.
2. **점검 목록**: 문단과 표의 글 칸을 위치와 함께 모은다. `--since 전.dsd` 를 주면 **바뀐 글만** 모은다.
   Claude 가 이 목록(`--json`)을 읽고 맞춤법·띄어쓰기 교정안을 `원문 / 교정안 / 위치` 로 작업 메모에
   적는다(스킬 3절). 고칠지는 사람이 정한다.

읽기 전용이며 종료 코드는 언제나 0 이다.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from .docmap import MapBlock, MapTable, load
from .errors import OutputError

#: 기계 점검 규칙 `(이름, 정규식)`.
RULES = (
    ("같은 낱말 연속", re.compile(r"(?<![가-힣A-Za-z])([가-힣]{1,8}|[A-Za-z]{2,12})\s+\1(?![가-힣A-Za-z])")),
    ("낱자모", re.compile(r"[ㄱ-ㅎㅏ-ㅣ]")),
    ("문장부호 앞 공백", re.compile(r"[가-힣A-Za-z0-9)]\s+[,.](?=\s|$)")),
    ("마침표 두 번", re.compile(r"(?<![.\d])\.\.(?![.\d])")),
    ("로마숫자 섞임", re.compile(r"(?:^|[\s(])(?:[IVX]+l|Vl+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][IVXl]|[IVXl][ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])[.)]")),
    ("콤마 자리", re.compile(r"\d{4,},\d{3}(?!\d)|(?<!\d)\d{1,3}(?:,\d{3})*,\d{4,}")),
    ("단위 중복", re.compile(r"(?:천원|백만원|원)원")),
)
_PAIRS = (("(", ")"), ("（", "）"), ("[", "]"), ("<", ">"), ("「", "」"))
#: 점검 목록에 넣을 표 글 칸의 최소 길이 (공백 제외).
MIN_CELL = 2


def _norm(text) -> str:
    return " ".join(str(text or "").split())


@dataclass(frozen=True)
class ProofItem:
    sheet: str
    row: int
    kind: str          # "문단" | "표"
    text: str

    @property
    def where(self) -> str:
        return f"{self.sheet} {self.row}행" if self.row else self.sheet


@dataclass(frozen=True)
class Finding:
    item: ProofItem
    rule: str
    hit: str

    def line_ko(self) -> str:
        t = self.item.text
        i = t.find(self.hit)
        snippet = t[max(0, i - 15): i + len(self.hit) + 15] if i >= 0 else t[:40]
        return f"[{self.rule}] {self.item.where} '{self.hit.strip()}' | …{snippet}…"


#: 괄호 짝을 셀 때 빼는 번호 매김 — `1)`, `가)`, `(1)` 은 짝이 없어도 정상이다.
_ENUM = re.compile(r"(?:^|(?<=\s))(?:\d{1,2}|[가나다라마바사아자차카타파하])\)")
#: 표 칸은 글자를 띄워 쓰는 머리글(`자 산 총 계`, `주 주`)이 많아 이 규칙을 쓰지 않는다.
TABLE_SKIP = frozenset({"같은 낱말 연속"})


def check_text(text: str, table: bool = False) -> list:
    """한 글의 기계 점검 결과 `[(규칙, 맞은 글)]`. `table=True` 면 표 칸 기준(띄어 쓴 머리글 허용)."""
    out = []
    for name, rx in RULES:
        if table and name in TABLE_SKIP:
            continue
        for m in rx.finditer(text):
            out.append((name, m.group(0)))
    bare = _ENUM.sub("", text)
    for a, b in _PAIRS:
        if bare.count(a) != bare.count(b):
            out.append(("괄호 짝", f"{a}{bare.count(a)} / {b}{bare.count(b)}"))
    return out


@dataclass
class ProofReport:
    path: Path
    since: Path | None = None
    items: list = field(default_factory=list)
    findings: list = field(default_factory=list)

    @property
    def exit_code(self) -> int:
        return 0

    def summary_ko(self) -> str:
        scope = f"바뀐 글 {len(self.items)}개" if self.since else f"글 {len(self.items)}개"
        return f"맞춤법 점검: {scope} / 기계 점검 후보 {len(self.findings)}건 (교정안은 --json 목록을 읽어 사람이 판단)"

    def lines_ko(self, verbose: bool = False) -> list:
        lines = [f.line_ko() for f in self.findings]
        if verbose:
            lines.extend(f"[점검 대상] {it.where} ({it.kind}) {it.text[:100]}" for it in self.items)
        lines.append(self.summary_ko())
        return lines

    def json_obj(self) -> dict:
        return {"file": str(self.path), "since": str(self.since) if self.since else "",
                "items": [{"sheet": i.sheet, "row": i.row, "kind": i.kind, "text": i.text} for i in self.items],
                "findings": [{"sheet": f.item.sheet, "row": f.item.row, "rule": f.rule, "hit": f.hit,
                              "text": f.item.text} for f in self.findings]}

    def write_json(self, path) -> None:
        p = Path(path)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(json.dumps(self.json_obj(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 점검 목록을 쓸 수 없습니다: {p}\n  사유: {exc}") from exc


def _items(doc) -> list:
    out = []
    for sheet in doc.sheets:
        for item in sheet.items:
            if isinstance(item, MapBlock):
                t = _norm(item.text)
                if t:
                    out.append(ProofItem(sheet.name, item.excel_row, "문단", t))
            elif isinstance(item, MapTable):
                for r in tuple(item.headers) + tuple(item.rows):
                    for c in r.cells:
                        if c.amount is not None:
                            continue
                        t = _norm(c.text)
                        if len(t.replace(" ", "")) >= MIN_CELL and any(ch.isalpha() for ch in t):
                            out.append(ProofItem(sheet.name, r.excel_row, "표", t))
    return out


def proof(path, since=None) -> ProofReport:
    doc = load(path)
    report = ProofReport(Path(path), Path(since) if since else None)
    items = _items(doc)
    if since:
        old = {it.text for it in _items(load(since))}
        items = [it for it in items if it.text not in old]
    seen = set()
    for it in items:
        if (it.sheet, it.row, it.text) in seen:
            continue
        seen.add((it.sheet, it.row, it.text))
        report.items.append(it)
        report.findings.extend(Finding(it, rule, hit) for rule, hit in check_text(it.text, it.kind == "표"))
    return report


__all__ = ["RULES", "ProofItem", "Finding", "ProofReport", "check_text", "proof"]
