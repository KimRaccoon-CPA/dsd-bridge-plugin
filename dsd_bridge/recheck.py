"""되돌린 결과 값 재확인 (0.7.0).

import 가 만든 contents.xml 을 **다시 펼쳐**(export 와 같은 투영) 작업 엑셀의 보이는 값과 칸마다 맞춰 본다.
바이트 패치·구조 동등성 검사는 "요소가 제자리에 있는지" 를 보장하지만, 이 검사는 "엑셀에 적은 값이 그대로
들어갔는지" 를 결과 쪽에서 한 번 더 본다. 소수점이 표 형식에 맞춰 잘린 칸, 글자로 적은 숫자, 서식 조각이 섞인
칸의 글자처럼 사람이 눈치채기 어려운 차이를 찾는다.

판정은 경고뿐이고 출력은 막지 않는다. 비교 규칙:
- 숫자는 값으로 (1,234 / 1234 / 1234.0 은 같다), 글자는 공백을 빼고 비교한다.
- 섹션 순서로 시트를 짝짓고, 시트 안에서는 매핑 행(문단·표 행)의 순서로 짝짓는다. 행 수가 다르면 그 시트는
  칸 비교를 건너뛰고 행 수 차이만 알린다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation

from .model import DocModel, project
from .sheetstate import MAPPED_KINDS, RO_SUFFIX
from .text import canonical_out
from .xmlsplice import SourceIndex


@dataclass
class Recheck:
    checked: int = 0
    mismatches: list = field(default_factory=list)     # (위치, 엑셀 값, 결과 값)
    row_count: list = field(default_factory=list)      # (시트, 엑셀 행 수, 결과 행 수)

    def line_ko(self) -> str:
        if not self.mismatches and not self.row_count:
            return f"[재확인] 되돌린 결과를 다시 펼쳐 엑셀 값 {self.checked:,}칸과 맞춰 봄: 모두 일치"
        return (f"[재확인] 되돌린 결과를 다시 펼쳐 엑셀 값 {self.checked:,}칸과 맞춰 봄: 다른 칸 "
                f"{len(self.mismatches)}개" + (f", 행 수가 다른 시트 {len(self.row_count)}개" if self.row_count else ""))

    def warnings_ko(self) -> list:
        out = []
        if self.mismatches:
            show = "; ".join(f"{w} 엑셀 '{a[:20]}' → 결과 '{b[:20]}'" for w, a, b in self.mismatches[:6])
            more = f" 외 {len(self.mismatches) - 6}곳" if len(self.mismatches) > 6 else ""
            out.append(f"경고: 되돌린 결과의 값이 엑셀과 다른 칸이 {len(self.mismatches)}개 있습니다({show}{more}). "
                       "소수점이 표 형식(자릿수)에 맞춰 반올림됐거나 숫자를 글자로 적은 경우가 많습니다. 의도한 값인지 "
                       "확인하세요.")
        for sheet, a, b in self.row_count:
            out.append(f"경고: 시트 '{sheet}' 의 문단·표 행 수가 엑셀 {a}개, 결과 {b}개로 다릅니다. 행 끼우기·지우기가 "
                       "의도대로 됐는지 DART 편집기에서 확인하세요.")
        return out


def _num(s: str):
    try:
        t = s.replace(",", "").strip()
        if t.startswith("(") and t.endswith(")"):
            t = "-" + t[1:-1]
        return Decimal(t)
    except (InvalidOperation, ValueError):
        return None


def _same(excel: str, result: str) -> bool:
    if excel == result:
        return True
    a, b = _num(excel), _num(result)
    if a is not None and b is not None:
        return a == b
    return "".join(excel.split()) == "".join(result.split())


def _rows(proj_sheet: dict) -> list:
    return [r for r in proj_sheet["rows"] if r["kind"].removesuffix(RO_SUFFIX) in MAPPED_KINDS]


def recheck(state, result: bytes, section_sheet) -> Recheck:
    """`section_sheet(entry)` = 섹션 지도 항목 → 통합문서 시트(없으면 None). import 의 `_section_sheet`."""
    proj = project(DocModel.build(SourceIndex.build(result), result, split_notes=state.split_notes))
    rc = Recheck()
    for entry, ps in zip(state.section_map, proj["sheets"]):
        sheet = section_sheet(entry)
        if sheet is None:
            continue
        wb_rows = [r for r in sheet.rows
                   if not r.is_header and not r.is_guide
                   and ((r.seq is not None and r.base_kind in MAPPED_KINDS and r.id)
                        or (r.seq is None and (r.id or r.kind or r.table or r.cells or r.values)))]
        res_rows = _rows(ps)
        if len(wb_rows) != len(res_rows):
            rc.row_count.append((sheet.name, len(wb_rows), len(res_rows)))
            continue
        for wr, pr in zip(wb_rows, res_rows):
            cols = pr.get("cols")
            for i, v in enumerate(pr["values"]):
                col = cols[i] if cols and i < len(cols) else i
                excel = canonical_out(wr.value_at(col))
                rc.checked += 1
                if not _same(excel, v["out"]):
                    rc.mismatches.append((f"{sheet.name} {wr.row}행 {_col(col)}열", excel, v["out"]))
    return rc


def _col(i: int) -> str:
    n, s = i + 6, ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


__all__ = ["Recheck", "recheck"]
