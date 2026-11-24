"""검증 게이트 — AC-23 / AC-28 / AC-29 / AC-30 / AC-31 의 판정 함수 (계획 §3.3, §4.8).

각 함수는 **판정만** 하고 예외를 던지지 않는다. `GateResult.violations` 가 비어 있으면 "발화하지
않음"이다. import 파이프라인(`dsd_bridge.import_xlsx`)이 순서대로 불러 첫 위반에서 해당 종료 코드로
중단하고, `scripts/roundtrip_check.py` 는 다섯 게이트의 상태를 한 줄로 출력한다.

**규범 2 (계획 §4.4)**: 새 게이트는 두 샘플의 무편집 왕복(AC-32)에서 발화하지 않음을 단언하는
테스트와 함께만 추가한다. `tests/test_export.py` (export 직후 사전 점검)와 단계 5 의 전체 왕복
테스트가 그 단언이다.

검사 대상 행(AC-28 / AC-29(1)(2)(4) / AC-30) = `_KIND` 에서 `:RO` 를 뗀 값이 `{TR, THEAD-TR, P, TITLE}`
이고 `_SEQ != ""` 인 행. 헤더 행·안내 행·`_SEQ` 가 빈 삽입 행은 제외한다. **단 AC-29(3) `_SEQ` 중복
검사만은 시트의 모든 행**(안내·헤더 포함, `_SEQ` 가 빈 행만 제외)에 적용한다.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from .model import WORKBOOK_FORMAT_VERSION
from .sheetstate import MAPPED_KINDS, SheetRow, SheetState, WorkbookState, anchor_table_row

#: AC-29(1) — `_ID` 형식. 순번은 `E{순번:05d}` 로 쓰므로 보통 5자리지만, 요소가 10만 개를
#: 넘는 문서에서는 자연히 6~7자리로 늘어난다 (두 샘플은 5자리).
EID_RE = re.compile(r"^E(\d{5,7})\.[0-9a-f]{4}$")


@dataclass(frozen=True)
class Violation:
    sheet: str
    row: int                 # 1-기반 Excel 행 번호 (시트 무관이면 0)
    code: int                # 종료 코드
    message: str


@dataclass
class GateResult:
    name: str                # "AC-23" 등
    violations: list = field(default_factory=list)
    checked: int = 0         # 실제로 검사한 행 수
    note: str = ""           # 요약 출력용 부가 정보

    @property
    def ok(self) -> bool:
        return not self.violations

    @property
    def status_ko(self) -> str:
        if self.ok:
            return "통과"
        v = self.violations[0]
        return f"발화({v.sheet} {v.row}행)" if v.row else "발화"


def checked_rows(sheet: SheetState) -> list:
    """AC-28 / AC-29(1)(2)(4) / AC-30 의 검사 대상 행."""
    return [r for r in sheet.rows if r.is_mapped_kind and r.seq is not None]


# ─────────────────────────────────────────────────────────────────────────────
# AC-31 — _dsdx 메타
# ─────────────────────────────────────────────────────────────────────────────


def gate_meta(state: WorkbookState) -> GateResult:
    res = GateResult("AC-31")
    if state.meta is None:
        res.violations.append(Violation("", 0, 2, (
            "오류: 이 통합문서에는 dsd-bridge 메타 시트(_dsdx)가 없습니다. dsd-bridge export 로 만든 .xlsx 가 아니거나, "
            "CSV/xls 로 저장하면서 숨김 시트가 사라졌을 수 있습니다. 원본 .dsd 를 다시 export 하세요."
        )))
        return res
    res.checked = 1
    got = str(state.meta.get("workbook_format_version", "") or "")
    if got != WORKBOOK_FORMAT_VERSION:
        res.violations.append(Violation("", 0, 2, (
            f"오류: 이 통합문서는 형식 버전 {got or '(없음)'} 로 만들어졌고 현재는 "
            f"{WORKBOOK_FORMAT_VERSION} 입니다. 같은 버전으로 다시 export 하세요."
        )))
    return res


# ─────────────────────────────────────────────────────────────────────────────
# AC-29(3) — _SEQ 중복 (모든 행 대상, 제외보다 먼저)
# ─────────────────────────────────────────────────────────────────────────────


def gate_seq_duplicates(sheet: SheetState) -> GateResult:
    res = GateResult("AC-29(3)")
    counts = Counter(r.seq for r in sheet.rows if r.seq is not None)
    res.checked = sum(counts.values())
    seen = set()
    for r in sheet.rows:
        if r.seq is None or counts[r.seq] < 2 or r.seq in seen:
            continue
        seen.add(r.seq)
        res.violations.append(Violation(sheet.name, r.row, 4, (
            f"오류: 시트 '{sheet.name}' {r.row}행 — 채우기 핸들로 행이 복제되었습니다"
            f"(순번 {r.seq} 중복 {counts[r.seq]}회). Ctrl+Z 로 되돌리세요. "
            "행을 새로 넣으려면 빈 행을 삽입한 뒤 값을 입력하세요."
        )))
    return res


# ─────────────────────────────────────────────────────────────────────────────
# AC-29(1)(2)(4) — _ID 형식 · 체크섬 · 중복 (+ 숨김 열 손상)
# ─────────────────────────────────────────────────────────────────────────────


def gate_id_integrity(sheet: SheetState, index) -> GateResult:
    """`index` 는 원본의 `SourceIndex` — 체크섬은 `index.eid_of(순번)` 과 비교한다."""
    res = GateResult("AC-29")
    rows = checked_rows(sheet)
    res.checked = len(rows)
    # (0) 숨김 열 손상 — 코드 6. 검사 집합 밖으로 조용히 빠져 삭제로 오분류되는 일을 막는다.
    for r in sheet.rows:
        if r.seq_corrupt:
            res.violations.append(Violation(sheet.name, r.row, 6, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 숨김 열 _SEQ 값이 정수가 아닙니다({r.seq_raw!r}). "
                "숨김 A~E열은 편집하면 안 됩니다. export 를 다시 하거나 그 행을 원래대로 되돌리세요."
            )))
            continue
        if r.hidden_formula:
            res.violations.append(Violation(sheet.name, r.row, 6, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 숨김 A~E열에 수식이 들어 있습니다. "
                "숨김 열은 편집하면 안 됩니다. Ctrl+Z 로 되돌리거나 export 를 다시 하세요."
            )))
            continue
        if r.seq is not None and not r.is_header and not r.is_guide and not r.is_mapped_kind:
            # _SEQ 는 있는데 _KIND 가 비었거나(예: 'tr', '') 알 수 없는 값 — 그대로 두면 plan_rows 가
            # 이 행을 건너뛰어 "시트에서 사라진 행" 으로 오분류하고 요소를 조용히 삭제한다.
            res.violations.append(Violation(sheet.name, r.row, 6, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 숨김 열 _KIND 값이 올바르지 않습니다({r.kind!r}). "
                "숨김 A~E열은 편집하면 안 됩니다. Ctrl+Z 로 되돌리거나 export 를 다시 하세요."
            )))
    for r in rows:
        if not r.id:
            # §4.4 행 분류 표: _SEQ 값 있음 + _ID 빈 값 → 손상
            res.violations.append(Violation(sheet.name, r.row, 6, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 순번(_SEQ)은 있는데 요소 ID(_ID)가 비어 있습니다. "
                "숨김 열이 손상되었습니다. Ctrl+Z 로 되돌리거나 export 를 다시 하세요."
            )))
            continue
        m = EID_RE.match(r.id)
        if not m:
            res.violations.append(Violation(sheet.name, r.row, 6, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 요소 ID 형식이 올바르지 않습니다({r.id!r}). "
                "숨김 A~E열은 편집하면 안 됩니다."
            )))
            continue
        ordinal = int(m.group(1))
        if ordinal >= len(index) or index.eid_of(ordinal) != r.id:
            res.violations.append(Violation(sheet.name, r.row, 4, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 요소 ID {r.id} 가 원본의 요소와 맞지 않습니다"
                "(체크섬 불일치). 숨김 열이 바뀌었거나 다른 문서의 통합문서입니다. Ctrl+Z 로 되돌리세요."
            )))
    # (4) _SEQ 가 채워진 행들 사이의 _ID 중복
    counts = Counter(r.id for r in rows if r.id)
    seen = set()
    for r in rows:
        if not r.id or counts[r.id] < 2 or r.id in seen:
            continue
        seen.add(r.id)
        res.violations.append(Violation(sheet.name, r.row, 4, (
            f"오류: 시트 '{sheet.name}' {r.row}행 — 같은 요소를 가리키는 행이 {counts[r.id]}개 있습니다"
            f"(요소 {r.id}). 채우기 핸들로 숨김 열이 끌려 내려간 것으로 보입니다. Ctrl+Z 로 되돌리세요."
        )))
    return res


# ─────────────────────────────────────────────────────────────────────────────
# AC-28 — 단조성 (_SEQ 기준)
# ─────────────────────────────────────────────────────────────────────────────


def gate_monotonic(sheet: SheetState) -> GateResult:
    res = GateResult("AC-28")
    rows = checked_rows(sheet)
    res.checked = len(rows)
    prev = None
    for r in rows:
        if prev is not None and r.seq <= prev.seq:
            res.violations.append(Violation(sheet.name, r.row, 4, (
                f"오류: 시트 '{sheet.name}' 의 행 순서가 원본과 다릅니다({r.row}행). "
                "이 도구는 행 재배열을 지원하지 않습니다. Excel에서 Ctrl+Z 로 정렬을 되돌린 뒤 다시 저장하세요."
            )))
            break
        prev = r
    return res


# ─────────────────────────────────────────────────────────────────────────────
# AC-30 (a) — 열 삽입 (비대칭 술어)
# ─────────────────────────────────────────────────────────────────────────────


def _anchor_cell_count(sheet: SheetState, i: int) -> int | None:
    """삽입 행의 셀 수 상한 — 바로 위(없으면 아래)의 원본 표 행의 `len(_CELLS)`."""
    anchor = anchor_table_row(sheet.rows, i)
    return None if anchor is None else len(anchor.cell_ids)


def gate_column_overflow(sheet: SheetState) -> GateResult:
    """**셀이 없는 열에 값이 있을 때만** 발화한다. 빈 셀은 검사하지 않는다.

    `seq` 배치에서는 예전과 같은 판정이다 — 셀이 놓이는 열은 `0..len(_CELLS)-1` 이므로 값이 든
    최우측 열이 `len(_CELLS)` 이상이면 발화한다. `adelim` 배치(1.2 WP3)에서는 셀이 놓이는 열이
    `_CELLS` 순서의 논리 열 목록(`SheetRow.cols`)이므로, 표 오른쪽 바깥뿐 아니라 **가운데의 빈
    논리 열**에 값을 넣은 것도 같은 코드 5 로 잡는다.
    """
    res = GateResult("AC-30")
    for i, r in enumerate(sheet.rows):
        if r.is_header or r.is_guide:
            continue
        if r.seq is not None:
            if r.base_kind not in ("TR", "THEAD-TR"):
                continue
            limit = len(r.cell_ids)
        else:
            # 삽입 행 — 표 안이면 앵커의 셀 수, 표 밖(P/TITLE 사이)이면 검사하지 않는다
            limit = _anchor_cell_count(sheet, i)
            if limit is None:
                continue
        res.checked += 1
        if limit == 0 and r.seq is not None:
            res.violations.append(Violation(sheet.name, r.row, 6, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 숨김 열 _CELLS 가 비어 있습니다(표 행인데 셀 목록이 없음). "
                "숨김 A~E열은 편집하면 안 됩니다. Ctrl+Z 로 되돌리거나 export 를 다시 하세요."
            )))
            continue
        idx = r.unmapped_value_col(limit)
        if idx < 0:
            continue
        mapped = sorted(r.mapped_cols(limit))
        if idx > mapped[-1]:
            res.violations.append(Violation(sheet.name, r.row, 5, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 표 오른쪽에 열이 추가되어 값이 매핑 범위 밖에 "
                f"있습니다(값이 든 열 {idx + 1}번째, 원본 셀 {limit}개). 이 도구는 열 추가·삽입을 지원하지 "
                "않습니다. 추가한 열을 지우고 다시 저장하세요."
            )))
        else:
            res.violations.append(Violation(sheet.name, r.row, 5, (
                f"오류: 시트 '{sheet.name}' {r.row}행 — 이 행에 셀이 없는 열에 값이 있습니다"
                f"(값이 든 열 {idx + 1}번째, 이 행의 셀 {limit}개가 놓이는 열 "
                f"{', '.join(str(c + 1) for c in mapped)}번째). 표의 빈 열에는 값을 넣을 수 없습니다 — "
                "이 도구는 열 추가·삽입을 지원하지 않습니다. 넣은 값을 지우고 다시 저장하세요."
            )))
    return res


# ─────────────────────────────────────────────────────────────────────────────
# AC-23 — 대량 소실 (section_map 순회)
# ─────────────────────────────────────────────────────────────────────────────


def gate_bulk_delete(state: WorkbookState) -> GateResult:
    """`section_map` 을 순회한다 — 시트를 통째로 지우면 그 섹션은 100% 소실이다.

    섹션과 시트의 대응은 이름이 아니라 `title_id` 로 정한다 (`WorkbookState.section_sheet`) —
    시트 이름을 바꿔도 섹션이 유지되고, 이름만 같고 제목이 다른 시트는 오류다.
    """
    res = GateResult("AC-23")
    for entry in state.section_map:
        name = entry["sheet"]
        total = int(entry["row_count"])
        found = state.section_sheet(entry)
        if found.conflict:
            res.checked += 1
            res.violations.append(Violation(name, 0, 4, found.conflict))
            continue
        sheet = found.sheet
        present = len(checked_rows(sheet)) if sheet is not None else 0
        deleted = max(0, total - present)
        res.checked += 1
        if total and deleted * 2 > total:
            label = sheet.name if sheet is not None else name
            where = "시트가 없습니다" if sheet is None else f"원본 행 {total}개 중 {deleted}개가 사라졌습니다"
            res.violations.append(Violation(label, 0, 4, (
                f"오류: 시트 '{label}' — {where}(50% 초과). 의도한 삭제라면 --allow-bulk-delete 를 사용하세요."
            )))
    return res


# ─────────────────────────────────────────────────────────────────────────────
# 묶음 실행 (roundtrip_check · 테스트용)
# ─────────────────────────────────────────────────────────────────────────────


def run_sheet_gates(sheet: SheetState, index) -> dict:
    """한 시트에 대한 AC-29(3) → AC-29 → AC-28 → AC-30 결과."""
    return {
        "AC-29(3)": gate_seq_duplicates(sheet),
        "AC-29": gate_id_integrity(sheet, index),
        "AC-28": gate_monotonic(sheet),
        "AC-30": gate_column_overflow(sheet),
    }


def run_all_gates(state: WorkbookState, index) -> dict:
    """AC-31 / AC-23 / AC-28 / AC-29 / AC-30 을 전부 돌려 게이트 이름 → `GateResult` 로 모은다.

    시트별 게이트는 모든 시트의 위반을 합친다(첫 위반에서 멈추지 않는다 — 진단용).
    """
    out = {"AC-31": gate_meta(state)}
    merged = defaultdict(lambda: None)
    for name in ("AC-29(3)", "AC-29", "AC-28", "AC-30"):
        merged[name] = GateResult(name)
    for sheet in state.sheets:
        for name, res in run_sheet_gates(sheet, index).items():
            merged[name].violations.extend(res.violations)
            merged[name].checked += res.checked
    out.update(merged)
    out["AC-23"] = gate_bulk_delete(state) if state.meta is not None else GateResult("AC-23", note="건너뜀(_dsdx 없음)")
    return out


__all__ = [
    "EID_RE",
    "Violation",
    "GateResult",
    "checked_rows",
    "gate_meta",
    "gate_seq_duplicates",
    "gate_id_integrity",
    "gate_monotonic",
    "gate_column_overflow",
    "gate_bulk_delete",
    "run_sheet_gates",
    "run_all_gates",
]
