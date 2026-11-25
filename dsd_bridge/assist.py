"""작은 도우미 명령 (0.7.0).

- `dsd-bridge round` — 단위 바꾸기 반올림 도우미. 원 단위 금액들을 천원·백만원으로 반올림할 때 **합계가 맞도록**
  가장 적게 고칠 칸을 고른다(최대 잔여 방식). 주석을 천원 단위로 쓸 때 "구성요소 반올림 합 ≠ 합계 반올림"
  문제를 미리 푼다.
- `dsd-bridge replace` — 보고서 기본정보 한꺼번에 바꾸기. 작업 엑셀의 보이는 칸(회색 칸 포함)에서 회사 이름·
  대표이사·주소 같은 글자를 한 번에 바꾼다. 기본은 미리보기, `--apply` 로 저장.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from .errors import InputError, OutputError
from .text import DEFAULT_PATTERN, format_number

UNITS = {"원": 1, "천원": 1000, "백만원": 1000000}


def _dec(v) -> Decimal:
    try:
        return Decimal(str(v).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        raise InputError(f"오류: 숫자가 아닙니다: {v}")


@dataclass
class RoundResult:
    unit: str
    amounts: list
    simple: list
    result: list
    target: Decimal

    @property
    def adjusted(self) -> list:
        return [i for i, (a, b) in enumerate(zip(self.simple, self.result)) if a != b]

    def lines_ko(self) -> list:
        lines = [f"단위 {self.unit} 반올림 (합계 {format_number(self.target, DEFAULT_PATTERN)} 에 맞춤)",
                 "번호 | 원 단위 | 단순 반올림 | 결과 | 조정"]
        for i, (a, s, r) in enumerate(zip(self.amounts, self.simple, self.result), 1):
            mark = f"{r - s:+}" if r != s else ""
            lines.append(f"{i} | {format_number(a, DEFAULT_PATTERN)} | {format_number(s, DEFAULT_PATTERN)} | "
                         f"{format_number(r, DEFAULT_PATTERN)} | {mark}")
        lines.append(f"단순 반올림 합 {format_number(sum(self.simple), DEFAULT_PATTERN)} / 결과 합 "
                     f"{format_number(sum(self.result), DEFAULT_PATTERN)} / 고친 칸 {len(self.adjusted)}개")
        return lines


def round_to_unit(amounts, unit: str = "천원", total=None) -> RoundResult:
    """원 단위 금액들 → 단위 반올림. 결과 합이 `total`(없으면 금액 합)을 반올림한 값과 같도록 조정한다."""
    if unit not in UNITS:
        raise InputError(f"오류: 단위는 {', '.join(UNITS)} 중 하나여야 합니다: {unit}")
    k = Decimal(UNITS[unit])
    vals = [_dec(a) for a in amounts]
    if not vals:
        raise InputError("오류: 금액을 하나 이상 주세요.")
    simple = [(v / k).quantize(Decimal(1), rounding=ROUND_HALF_UP) for v in vals]
    base = _dec(total) if total is not None else sum(vals)
    target = (base / k).quantize(Decimal(1), rounding=ROUND_HALF_UP)
    result = list(simple)
    diff = int(target - sum(simple))
    if diff:
        # 반올림으로 잃은(또는 얻은) 몫이 큰 칸부터 1 씩 고친다.
        loss = [(v / k) - s for v, s in zip(vals, simple)]
        order = sorted(range(len(vals)), key=lambda i: -loss[i] if diff > 0 else loss[i])
        step = 1 if diff > 0 else -1
        for i in order[:abs(diff)]:
            result[i] += step
    return RoundResult(unit, vals, simple, result, target)


def read_range(path, ref: str) -> list:
    """`시트!C5:C12` 범위의 숫자 (빈칸·글자는 건너뜀)."""
    from openpyxl import load_workbook

    if "!" not in ref:
        raise InputError("오류: 범위는 '시트!C5:C12' 꼴로 적어 주세요.")
    sheet, cells = ref.rsplit("!", 1)
    sheet = sheet.strip("'")
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise InputError(f"오류: 시트 '{sheet}' 가 없습니다.")
        out = []
        for row in wb[sheet][cells]:
            for c in (row if isinstance(row, tuple) else (row,)):
                if isinstance(c.value, (int, float, Decimal)) and not isinstance(c.value, bool):
                    out.append(Decimal(str(c.value)))
        return out
    finally:
        wb.close()


# ─────────────────────────────────────────────────────────────────────────────
# replace
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class ReplaceReport:
    path: Path
    pairs: list
    hits: list = field(default_factory=list)       # (위치, 전, 후)
    applied: bool = False

    def lines_ko(self) -> list:
        lines = [f"{where}\n   전: {old[:150]}\n   후: {new[:150]}" for where, old, new in self.hits]
        counts = ", ".join(f"'{o}'→'{n}'" for o, n in self.pairs)
        lines.append(f"바꿀 칸 {len(self.hits)}개 ({counts})"
                     + (" — 저장했습니다" if self.applied else " (미리보기, 반영하려면 --apply)"))
        return lines


def replace_text(path, pairs, apply: bool = False) -> ReplaceReport:
    """작업 엑셀의 보이는 칸(F 열부터, 숨김 시트 제외)에서 글자를 바꾼다. 긴 것부터 한 번에 치환한다."""
    import re

    from openpyxl import load_workbook

    from .sheetstate import is_internal_sheet

    p = Path(path)
    if p.suffix.lower() != ".xlsx":
        raise InputError("오류: replace 는 작업 엑셀(.xlsx)에서만 씁니다.")
    if not pairs:
        raise InputError("오류: --set \"옛=새\" 를 하나 이상 주세요.")
    table = dict(pairs)
    rx = re.compile("|".join(re.escape(k) for k in sorted(table, key=len, reverse=True)))
    wb = load_workbook(p)
    report = ReplaceReport(p, list(pairs))
    for ws in wb.worksheets:
        if is_internal_sheet(ws.title):
            continue
        for row in ws.iter_rows(min_col=6):
            for c in row:
                if not isinstance(c.value, str):
                    continue
                new = rx.sub(lambda m: table[m.group(0)], c.value)
                if new != c.value:
                    report.hits.append((f"{ws.title}!{c.coordinate}", c.value, new))
                    if apply:
                        c.value = new
    if apply and report.hits:
        try:
            wb.save(p)
        except OSError as exc:
            raise OutputError(f"오류: 저장할 수 없습니다(엑셀에서 열려 있나요?): {p}\n  사유: {exc}") from exc
        report.applied = True
    return report


__all__ = ["UNITS", "RoundResult", "round_to_unit", "read_range", "ReplaceReport", "replace_text"]
