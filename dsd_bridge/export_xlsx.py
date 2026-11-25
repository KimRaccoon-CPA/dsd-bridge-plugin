"""`.dsd` → 편집용 Excel 통합문서 (계획 §4.4, §5 단계 4 / AC-1, 2′, 3′, 12, 16, 17, 26, 27).

레이아웃(시트마다):

- **A~E 열 숨김** — `_ID` / `_KIND` / `_TABLE` / `_CELLS` / `_SEQ`. 1행 A1 은 `⚠ 편집 금지 (dsd-bridge 내부용)`.
- **F 열부터 표시** — 열 위치는 `project()` 가 표마다 정한다(`dsd_bridge.model` 의 열 배치):
  `seq` 배치면 F=0번째 셀, G=1번째 … (`TR` 내 순서, AC-17), `adelim` 배치면 F + `ADELIM`
  (1.2 WP3, 행의 `cols`). `P`/`TITLE` 행은 F 에 텍스트.
- `freeze_panes = "F1"`.
- 숫자 셀은 숫자형 + `#,##0` 계열 서식(AC-12), 그 밖은 `'@'`(AC-26). `&cr;` 는 `\\n` 으로 보이고
  `wrap_text=True`(AC-14a).
- 혼합 콘텐츠(읽기 전용)는 회색 `FFEFEFEF` (AC-16). `TITLE` 행은 굵게 + `FFD9E1F2`.
- `seq` 배치에서 `COLSPAN` 은 **다음 셀과 충돌하지 않을 때만** 병합한다. 셀은 `TR` 순서대로
  놓이므로 뒤에 셀이 더 있는 `COLSPAN` 셀을 병합하면 그 뒤 셀이 가려진다. 건너뛴 수는 리포트에
  남긴다. `adelim` 배치에서는 배치가 이미 겹치지 않으므로 행의 `spans` 대로 그냥 병합한다
  (머리글 `TH COLSPAN=2` 가 논리 열 2칸을 덮는다). `merge=False` 면 병합을 전혀 하지 않는다
  (테스트용, MAJOR-10). `ROWSPAN` 은 세로 병합하지 않는다(AC-17).
- `_dsdx` 시트(`sheet_state="hidden"`) — key/value 메타 (`contents_sha256`, `element_count`,
  `tag_digest`, `workbook_format_version`, `section_map` …). `_dsdx_raw` 는 만들지 않는다.
- `_dsdx_attrs` 시트(`sheet_state="hidden"`, `_dsdx` 다음) — 표 행마다 원문 DART 속성 한 줄
  (`dsd_bridge.attrs`). **참고용**이라 import 는 읽지 않고, 지워도 되돌리기에 영향이 없다.

시트에 쓰는 **내용**은 전부 `dsd_bridge.model.project()` 가 정한다 — import 는 원본에서 같은 투영을
다시 계산해 기준값으로 쓴다.
"""

from __future__ import annotations

import datetime as _dt
import json
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import __version__
from .attrs import ATTRS_HEADER, ATTRS_SHEET, attrs_rows
from .container import DsdContainer
from .errors import OutputError
from .model import (
    HEADER_KIND,
    HIDDEN_COLUMNS,
    RO_SUFFIX,
    WORKBOOK_FORMAT_VERSION,
    DocModel,
    project,
)
from .text import cr_to_lf
from .xmlsplice import SourceIndex, check_rewrite_identity

#: 숨김 메타 시트 이름.
META_SHEET = "_dsdx"
#: 표시 열이 시작하는 Excel 열 번호 (F).
FIRST_VALUE_COL = len(HIDDEN_COLUMNS) + 1
#: AC-27 — DART 추출 대상 표 안내 문구 (README 와 문자열 일치).
EXTRACTION_NOTE = "※ DART 추출 대상 표 — 값 변경 시 DART 편집기에서 재추출 필요"
#: AC-17 — ROWSPAN 표 안내 문구.
ROWSPAN_NOTE = "※ 세로 병합이 있어 열 위치가 원문과 다르게 보일 수 있습니다"
TABLE_START_TEXT = "▶ 표 시작"
TABLE_END_TEXT = "◀ 표 끝"

_FILL_RO = PatternFill("solid", fgColor="FFEFEFEF")
_FILL_TITLE = PatternFill("solid", fgColor="FFD9E1F2")
_FILL_GUIDE = PatternFill("solid", fgColor="FFF7F7F7")
_FONT_TITLE = Font(bold=True)
_FONT_GUIDE = Font(italic=True, color="FF808080")
_FONT_HEADER = Font(italic=True, color="FF9C0006")
_ALIGN_TOP = Alignment(vertical="top")
_ALIGN_WRAP = Alignment(wrap_text=True, vertical="top")
_ALIGN_NUM = Alignment(horizontal="right", vertical="top")


@dataclass
class ExportReport:
    """`export()` 의 결과 요약 (CLI 출력·테스트용)."""

    source: Path
    output: Path
    sheets: int = 0
    blocks: dict = field(default_factory=dict)      # {"TITLE": n, "P": n, "TABLE": n}
    rows: int = 0                                   # 시트에 매핑된 TR 수
    attr_rows: int = 0                              # `_dsdx_attrs` 본문 행 수 (= TR 수)
    mixed: int = 0
    merges_applied: int = 0
    merges_skipped: int = 0
    element_count: int = 0
    contents_sha256: str = ""
    warnings: list = field(default_factory=list)     # `경고:` 로 시작하는 한국어 문장
    notices: list = field(default_factory=list)      # `안내:` 로 시작하는 한국어 문장
    skipped: dict = field(default_factory=dict)      # `project()` 의 미투영 수치

    def summary_ko(self) -> str:
        b = self.blocks
        return (
            f"시트 {self.sheets} / 블록 {sum(b.values())} (TITLE {b.get('TITLE', 0)}, "
            f"P {b.get('P', 0)}, TABLE {b.get('TABLE', 0)}) / TR {self.rows} / 혼합 {self.mixed}"
            f"  →  {self.output}"
        )


def _table_guide_text(table) -> str:
    """`TABLE-START` 안내 행의 문구 (AC-27 / AC-17 표시)."""
    parts = [TABLE_START_TEXT]
    if table.extraction:
        parts.append(EXTRACTION_NOTE)
    if table.has_rowspan:
        parts.append(ROWSPAN_NOTE)
    return "  ".join(parts)


def _guide_text_from_flags(flags: dict) -> str:
    parts = [TABLE_START_TEXT]
    if flags.get("extraction"):
        parts.append(EXTRACTION_NOTE)
    if flags.get("rowspan"):
        parts.append(ROWSPAN_NOTE)
    return "  ".join(parts)


def _number_value(num: str):
    """`project()` 의 십진 문자열을 Excel 에 쓸 int/float 로."""
    d = Decimal(num)
    if d == d.to_integral_value():
        return int(d)
    return float(d)


def _set_text(cell, text: str) -> None:
    """문자열 셀 — 수식으로 오인되지 않게 하고 `'@'` 서식을 건다 (AC-26)."""
    cell.value = text
    if cell.data_type == "f":          # "=..." 로 시작하는 텍스트
        cell.data_type = "s"
    cell.number_format = "@"


def _apply_cell_style(cell, value: dict) -> None:
    """값 dict(`project()` 의 value) 하나를 셀에 쓴다."""
    if value["t"] == "n":
        cell.value = _number_value(value["num"])
        cell.number_format = value["fmt"]
        cell.alignment = _ALIGN_NUM
    else:
        text = cr_to_lf(value["raw"])
        _set_text(cell, text)
        cell.alignment = _ALIGN_WRAP if "\n" in text else _ALIGN_TOP
    if value.get("ro"):
        cell.fill = _FILL_RO


def _write_row(ws, r: int, row: dict, colspans: dict, merge: bool, report: ExportReport) -> None:
    kind = row["kind"]
    base = kind.removesuffix(RO_SUFFIX)
    ws.cell(row=r, column=1, value=row["id"])
    ws.cell(row=r, column=2, value=kind)
    ws.cell(row=r, column=3, value=row["table"])
    ws.cell(row=r, column=4, value=row["cells"])
    ws.cell(row=r, column=5, value=row["seq"])
    for col in range(1, FIRST_VALUE_COL):
        ws.cell(row=r, column=col).number_format = "@" if col != 5 else "0"

    if base == HEADER_KIND:
        c = ws.cell(row=r, column=FIRST_VALUE_COL)
        _set_text(c, row["values"][0]["raw"])
        c.font = _FONT_HEADER
        ws.cell(row=r, column=1).font = _FONT_HEADER
        return
    if base == "TABLE-START":
        c = ws.cell(row=r, column=FIRST_VALUE_COL)
        _set_text(c, _guide_text_from_flags(row.get("guide", {})))
        c.font = _FONT_GUIDE
        c.fill = _FILL_GUIDE
        return
    if base == "TABLE-END":
        c = ws.cell(row=r, column=FIRST_VALUE_COL)
        _set_text(c, TABLE_END_TEXT)
        c.font = _FONT_GUIDE
        c.fill = _FILL_GUIDE
        return
    if base in ("P", "TITLE"):
        c = ws.cell(row=r, column=FIRST_VALUE_COL)
        _apply_cell_style(c, row["values"][0])
        if base == "TITLE":
            c.font = _FONT_TITLE
            if not row["values"][0].get("ro"):
                c.fill = _FILL_TITLE
        return

    # 표 행 (TR / THEAD-TR)
    values = row["values"]
    n = len(values)
    cols = row.get("cols")
    if cols:
        # ADELIM 배치 (1.2 WP3) — 셀마다 정해진 논리 열에 놓는다. 빈 논리 열은 손대지 않는다.
        for v, c in zip(values, cols):
            _apply_cell_style(ws.cell(row=r, column=FIRST_VALUE_COL + c), v)
        if not merge:
            return
        for c, span in zip(cols, row.get("spans") or []):
            if span <= 1:
                continue
            # 배치가 이미 겹치지 않게 정해져 있으므로 건너뛸 일이 없다 (머리글 COLSPAN 병합).
            ws.merge_cells(start_row=r, start_column=FIRST_VALUE_COL + c,
                           end_row=r, end_column=FIRST_VALUE_COL + c + span - 1)
            report.merges_applied += 1
        return
    for i, v in enumerate(values):
        _apply_cell_style(ws.cell(row=r, column=FIRST_VALUE_COL + i), v)
    if not merge:
        return
    for i, v in enumerate(values):
        span = colspans.get(v["id"], 1)
        if span <= 1:
            continue
        # 다음 셀과 충돌하면 건너뛴다 — 셀은 TR 순서대로 놓이므로 뒤에 셀이 있으면 반드시 충돌한다.
        # 즉 병합할 수 있는 것은 행의 **마지막 셀**뿐이다.
        if i == n - 1:
            ws.merge_cells(
                start_row=r, start_column=FIRST_VALUE_COL + i,
                end_row=r, end_column=FIRST_VALUE_COL + i + span - 1,
            )
            report.merges_applied += 1
        else:
            report.merges_skipped += 1


def _write_section_sheet(ws, sheet: dict, colspans: dict, merge: bool, report: ExportReport) -> None:
    for letter in ("A", "B", "C", "D", "E"):
        ws.column_dimensions[letter].hidden = True
    ws.column_dimensions["F"].width = 34
    for col in range(FIRST_VALUE_COL + 1, FIRST_VALUE_COL + 12):
        ws.column_dimensions[get_column_letter(col)].width = 16
    for r, row in enumerate(sheet["rows"], start=1):
        _write_row(ws, r, row, colspans, merge, report)
    ws.freeze_panes = "F1"


def _write_meta_sheet(wb, meta: dict) -> None:
    ws = wb.create_sheet(META_SHEET)
    ws.sheet_state = "hidden"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 80
    ws.append(["key", "value"])
    for k, v in meta.items():
        ws.append([k, v if isinstance(v, (int, float)) else str(v)])
        c = ws.cell(row=ws.max_row, column=2)
        if isinstance(v, str):
            c.number_format = "@"


def _write_attrs_sheet(wb, proj: dict) -> int:
    """`_dsdx_attrs` 숨김 시트 — 표 행마다 원문 DART 속성 한 줄 (WP2). 행 수를 돌려준다.

    참고용이라 import 는 읽지 않는다. 값은 전부 문자열로 쓴다(`ACODE` 가 숫자로 바뀌면
    앞자리 0 이 사라진다).
    """
    ws = wb.create_sheet(ATTRS_SHEET)
    ws.sheet_state = "hidden"
    ws.append(list(ATTRS_HEADER))
    rows = attrs_rows(proj)
    for values in rows:
        ws.append(values)
    for col in range(1, len(ATTRS_HEADER) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 18
        for c in ws[letter]:
            c.number_format = "@"
    return len(rows)


def _colspans(model: DocModel) -> dict:
    """셀 eid → COLSPAN (1 초과만)."""
    out = {}
    for b in model.blocks:
        if b.table is None:
            continue
        for r in b.table.rows:
            for c in r.cells:
                if c.colspan > 1:
                    out[c.eid] = c.colspan
    return out


#: 중첩 표 경고 (문자열 일치 테스트 대상). 두 샘플은 0건이라 무편집 왕복에서 발화하지 않는다 (규범 2).
NESTED_TABLE_WARNING = (
    "경고: 중첩 표 {n}개는 Excel 에 나타나지 않습니다(편집 불가). 그 표의 내용은 되돌릴 때 "
    "원본 그대로 보존됩니다."
)
#: 머리말 안내 — 두 샘플에 모두 있으므로 **경고가 아니라 안내**다 (규범 2).
PREAMBLE_NOTICE = (
    "안내: 목차 앞 머리말 블록 {n}개는 어느 시트에도 나타나지 않습니다(편집 불가). "
    "되돌릴 때 원본 그대로 보존됩니다."
)


def _report_skipped(skipped: dict, report: ExportReport) -> None:
    """투영되지 않은 것을 사용자에게 알린다 — 중첩 표는 `경고:`, 머리말은 `안내:`."""
    report.skipped = dict(skipped)
    n_nested = int(skipped.get("nested_tables", 0) or 0)
    if n_nested:
        report.warnings.append(NESTED_TABLE_WARNING.format(n=n_nested))
    n_preamble = int(skipped.get("preamble_blocks", 0) or 0)
    if n_preamble:
        report.notices.append(PREAMBLE_NOTICE.format(n=n_preamble))


def _build_workbook(model: DocModel, container: DsdContainer, source: Path, *, merge: bool,
                    report: ExportReport) -> Workbook:
    proj = project(model)
    _report_skipped(proj.get("skipped", {}), report)
    colspans = _colspans(model)
    wb = Workbook()
    wb.remove(wb.active)
    for sheet in proj["sheets"]:
        ws = wb.create_sheet(sheet["name"])
        _write_section_sheet(ws, sheet, colspans, merge, report)

    index = model.index
    meta = {
        "dsdx_version": __version__,
        "workbook_format_version": WORKBOOK_FORMAT_VERSION,
        "created_at": _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat(),
        # 0.3.0: 전체 경로(PC 사용자 이름·작업 폴더 이름)는 남기지 않는다. 어디에서도 읽지 않는 값이었고
        # 작업 엑셀이 밖으로 나가면 함께 새어 나간다. 옛 작업 엑셀의 `source_path` 는 import 가 무시한다.
        "source_name": source.name,
        "contents_sha256": container.contents_sha256(),
        "meta_sha256": container.meta_sha256(),
        "contents_size": len(container.contents_bytes),
        "element_count": len(index),
        "tag_digest": index.tag_digest(),
        "schema": container.schema() or "",
        "editver": container.editver() or "",
        # 1.2 WP5 — 주석 번호 제목마다 시트를 나눴는지. import 는 이 값으로 같은 규칙을 다시 적용해
        # 섹션을 만들고 section_map 과 대조한다 (다르면 코드 2).
        "split_notes": 1 if model.split_notes else 0,
        "section_map": json.dumps(proj["section_map"], ensure_ascii=False, separators=(",", ":")),
    }
    _write_meta_sheet(wb, meta)
    report.attr_rows = _write_attrs_sheet(wb, proj)

    report.sheets = len(proj["sheets"])
    report.blocks = model.block_counts()
    report.rows = sum(len(t.rows) for s in model.sections for t in s.tables)
    report.mixed = len(model.mixed_ordinals())
    report.element_count = len(index)
    report.contents_sha256 = meta["contents_sha256"]
    return wb


def default_output_path(dsd_path) -> Path:
    """`-o` 생략 시: 입력과 같은 폴더의 `<입력명>.xlsx`."""
    p = Path(dsd_path)
    return p.with_suffix(".xlsx") if p.suffix.lower() == ".dsd" else p.with_name(p.name + ".xlsx")


def export(dsd_path, out_path=None, *, merge: bool = True, split_notes: bool = False,
           log=None) -> ExportReport:
    """`.dsd` 를 읽어 `.xlsx` 를 쓴다. 입력 문제는 `InputError`(2), 출력 문제는 `OutputError`(7).

    `split_notes=True` 면 주석 번호 제목(`dsd_bridge.model.note_heading`)마다 시트를 나눈다 (1.2 WP5).
    기본값 `False` 는 1.1 까지와 완전히 같은 투영이다.
    """
    log = log or (lambda *_: None)
    source = Path(dsd_path)
    output = Path(out_path) if out_path else default_output_path(source)

    container = DsdContainer.read(source)
    log(f"[1/4] 컨테이너      contents.xml {len(container.contents_bytes):,} B / "
        f"meta.xml {len(container.meta_bytes):,} B / schema={container.schema()} editver={container.editver()}")
    index = SourceIndex.build(container.contents_bytes)
    # 오프셋 산술이 이 문서에서 성립하는지 — import 와 같은 검사를 먼저 한다. 실패하면
    # 되돌릴 수 없는 통합문서를 만들게 되므로 여기서 코드 6 으로 중단한다.
    check_rewrite_identity(container.contents_bytes, index)
    log(f"[2/4] 오프셋 인덱스  요소 {len(index):,}개 / REWRITE_IDENTITY OK")
    model = DocModel.build(index, container.contents_bytes, split_notes=split_notes)
    report = ExportReport(source=source, output=output)
    wb = _build_workbook(model, container, source, merge=merge, report=report)
    log(f"[3/4] 통합문서      {report.summary_ko().split('  →')[0]}"
        f" / 병합 {report.merges_applied} (건너뜀 {report.merges_skipped})")
    if report.merges_skipped:
        # 구조적으로 늘 생기는 일(뒤에 셀이 있는 COLSPAN)이라 기본 출력을 어지럽히지 않고 -v 에만 남긴다.
        log(f"      COLSPAN 병합 {report.merges_skipped}개는 뒤에 셀이 있어 건너뛰었습니다"
            " (열 위치는 원문의 셀 순서를 따릅니다, AC-17).")
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
    except (OSError, PermissionError) as exc:
        raise OutputError(
            f"오류: 출력 파일을 쓸 수 없습니다: {output}\n  사유: {exc}\n"
            "  같은 이름의 파일이 Excel 에서 열려 있으면 닫고 다시 시도하세요."
        ) from exc
    log(f"[4/4] 저장          {output}")
    return report


__all__ = [
    "META_SHEET",
    "ATTRS_SHEET",
    "FIRST_VALUE_COL",
    "EXTRACTION_NOTE",
    "ROWSPAN_NOTE",
    "NESTED_TABLE_WARNING",
    "PREAMBLE_NOTICE",
    "TABLE_START_TEXT",
    "TABLE_END_TEXT",
    "ExportReport",
    "default_output_path",
    "export",
]
