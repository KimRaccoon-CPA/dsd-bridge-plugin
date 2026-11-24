"""편집한 `.xlsx` → 원본 `.dsd` 에 되돌리기 (계획 §4.1, §4.5, §4.7~4.10, §5 단계 5).

파이프라인 (반드시 이 순서)::

     1. _dsdx 검증                (AC-31)
     2. 해시·구조 검증            (AC-18: contents_sha256 / element_count / tag_digest)
     3. _SEQ 중복 검사            (AC-29(3)) — 모든 행, 제외보다 먼저
     4. 안내·헤더 행 제외
     5. _ID 형식·체크섬·중복      (AC-29(1)(2)(4))
     6. 행 분류                   (유지 / 삽입 / 복제 삽입 / 삭제 — rowedit.plan_rows)
     7. 단조성                    (AC-28)
     8. 열 검증                   (AC-30 (a))
     9. 대량 소실                 (AC-23)
    10. 기준값 재계산             (원본 → project(), `_dsdx.split_notes` 반영 후 섹션 지도 대조)
    11. 혼합 콘텐츠 게이트        (AC-16 / --allow-lossy-mixed → planned_deletions)
    12. 셀 diff → 패치 → 겹침 검사 → 적용
    13. §4.10 사후 검사 1~4       (well-formed / 요소 수 / 구조 동등성 / XSD)
    14. meta.xml size (AC-19) → ZIP 재포장 (AC-20)

**직렬화하지 않는다.** 원본 바이트에 패치를 스플라이스할 뿐이다.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from openpyxl.utils import get_column_letter

from .container import DsdContainer, meta_file_info_size, update_meta_size
from .errors import HashMismatch, InputError, OutputError, PatchIntegrityError, UnsupportedEdit
from .gates import (
    GateResult,
    gate_bulk_delete,
    gate_column_overflow,
    gate_id_integrity,
    gate_meta,
    gate_monotonic,
    gate_seq_duplicates,
)
from .model import CELL_TAGS, RO_SUFFIX, DocModel, project
from .rowedit import RowPlan, clone_p_bytes, clone_tr_bytes, delete_span, insert_patch, plan_rows
from .sheetstate import FIRST_VALUE_COL, WorkbookState, load_workbook_state
from .text import (
    canonical_out,
    count_bare_amp,
    escape_text,
    format_number,
    lf_to_cr,
    merge_mixed_text,
    resolve_pattern,
)
from .validate import (
    ValidationResult,
    check_wellformed,
    structural_equivalence,
    subtree_size,
    xsd_validate,
)
from .xmlsplice import PatchSet, SourceIndex, check_rewrite_identity, set_text_patch

#: AC-27 경고 (문자열 일치 테스트 대상).
EXTRACTION_WARNING = (
    "경고: DART 추출 대상 표의 값 {n}개가 변경되었습니다. 총자산/총부채/매출액 등 요약값(SUMMARY, "
    "meta.xml EXTRACT)은 자동으로 갱신되지 않습니다. DART 편집기에서 재추출한 뒤 제출하세요. 변경된 표: {tables}"
)
#: AC-16 오류 (문자열 일치 테스트 대상).
MIXED_EDIT_ERROR = (
    "오류: 시트 '{sheet}' {row}행 {col}열 (요소 {id}) 은(는) 서식 조각(SPAN/P)을 포함한 문단·셀입니다. "
    "지금 저장하면 그 서식이 사라집니다. 원래 값으로 되돌리거나, 손실을 감수하려면 --allow-lossy-mixed 를 사용하세요."
)


@dataclass
class CellChange:
    sheet: str
    row: int
    col: str                   # Excel 열 문자
    id: str
    ordinal: int
    old: str                   # 원본 원문 (unescape, &cr; 유지)
    new: str                   # 되돌릴 원문 (escape 전, &cr; 유지)
    kind: str                  # "cell" | "block"
    whitespace_only: bool = False
    mixed: bool = False
    inherited: object = None   # PatternResolution (숫자 상속이 있었을 때)
    table_id: str = ""


@dataclass
class ImportReport:
    xlsx: Path
    dsd: Path
    output: Path
    changes: list = field(default_factory=list)          # CellChange
    inserted: int = 0
    deleted: int = 0
    warnings: list = field(default_factory=list)
    gates: dict = field(default_factory=dict)            # 이름 → GateResult
    xsd_original: ValidationResult | None = None
    xsd_result: ValidationResult | None = None
    element_count_before: int = 0
    element_count_after: int = 0
    patch_stats: dict = field(default_factory=dict)
    force_hash_used: bool = False
    meta_size_updated: bool = False
    bytes_before: int = 0
    bytes_after: int = 0

    @property
    def changed_cells(self) -> int:
        return sum(1 for c in self.changes if not c.whitespace_only and c.kind in ("cell", "block"))

    @property
    def whitespace_only(self) -> list:
        return [c for c in self.changes if c.whitespace_only]

    @property
    def inherited(self) -> list:
        return [c for c in self.changes if c.inherited is not None and c.inherited.level != "self"]

    def gate_line_ko(self) -> str:
        order = ("AC-23", "AC-28", "AC-29", "AC-30", "AC-31")
        return " · ".join(f"{k} {self.gates[k].status_ko}" for k in order if k in self.gates)

    def xsd_line_ko(self) -> str:
        o, r = self.xsd_original, self.xsd_result
        if o is None or r is None:
            return "[6/6] XSD 검증: 건너뜀 (사유: 실행되지 않음)"
        if o.status == "skipped":
            return f"[6/6] XSD 검증: 건너뜀 (사유: {o.reason})"
        where = f"  ({Path(o.xsd_path).name})" if o.xsd_path else ""
        return f"[6/6] XSD 검증: 원본 {o.status_ko} / 결과 {r.status_ko}{where}"

    def summary_ko(self) -> str:
        lines = [
            f"되돌리기 완료: 변경 셀 {self.changed_cells} / 삽입 {self.inserted} / 삭제 {self.deleted} / "
            f"공백만 변경 {len(self.whitespace_only)}  →  {self.output}"
        ]
        for c in self.inherited:
            lines.append(
                f"  · {c.sheet} {c.row}행 {c.col}열 (요소 {c.id}) 표기 패턴: {c.inherited.describe()}"
            )
        for w in self.warnings:
            lines.append(w if w.startswith("경고") else f"경고: {w}")
        lines.append(self.xsd_line_ko())
        return "\n".join(lines)


def default_output_path(xlsx_path) -> Path:
    p = Path(xlsx_path)
    return p.with_name(p.stem + ".out.dsd")


def _raise_first(res: GateResult, *, allow: bool = False, code_override: int | None = None) -> None:
    if res.ok or allow:
        return
    v = res.violations[0]
    code = code_override if code_override is not None else v.code
    exc = {2: "InputError", 3: "HashMismatch", 4: "SheetIntegrityError", 5: "UnsupportedEdit",
           6: "PatchIntegrityError"}.get(code)
    from . import errors as _e

    cls = getattr(_e, exc) if exc else _e.DsdBridgeError
    raise cls(v.message, exit_code=code)


def _excel_col(i: int) -> str:
    return get_column_letter(FIRST_VALUE_COL + i)


def _is_number(v) -> bool:
    return isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)


def _whitespace_only(a: str, b: str) -> bool:
    return "".join(a.split()) == "".join(b.split())


def import_(xlsx, dsd, out=None, *, force_hash: bool = False, allow_bulk_delete: bool = False,
            allow_lossy_mixed: bool = False, no_meta_size: bool = False, log=None,
            dart_root=None) -> ImportReport:
    log = log or (lambda *_: None)
    xlsx_path = Path(xlsx)
    dsd_path = Path(dsd)
    out_path = Path(out) if out else default_output_path(xlsx_path)
    report = ImportReport(xlsx=xlsx_path, dsd=dsd_path, output=out_path)
    if out_path.resolve() == dsd_path.resolve():
        raise OutputError("오류: 출력 경로가 원본 .dsd 와 같습니다. 원본을 덮어쓰지 않도록 다른 경로를 지정하세요.")

    # 1. _dsdx
    state = load_workbook_state(xlsx_path)
    report.gates["AC-31"] = gate_meta(state)
    _raise_first(report.gates["AC-31"])
    meta = state.meta
    _check_formula_cells(state, report)
    log(f"[1/6] 통합문서      시트 {len(state.sheets)} / 형식 버전 {meta.get('workbook_format_version')}"
        f" / dsd-bridge {meta.get('dsdx_version')}")

    # 2. 원본·해시
    container = DsdContainer.read(dsd_path)
    raw = container.contents_bytes
    index = SourceIndex.build(raw)
    check_rewrite_identity(raw, index)          # 오프셋 산술이 이 문서에서 성립하는지
    report.element_count_before = len(index)
    report.bytes_before = len(raw)
    recorded = str(meta.get("contents_sha256") or "")
    actual = container.contents_sha256()
    if recorded != actual:
        conds = {
            "contents_size": str(meta.get("contents_size")) == str(len(raw)),
            "element_count": str(meta.get("element_count")) == str(len(index)),
            "tag_digest": str(meta.get("tag_digest") or "") == index.tag_digest(),
        }
        if not force_hash or not all(conds.values()):
            extra = ""
            if force_hash:
                failed = ", ".join(k for k, v in conds.items() if not v)
                extra = f"\n  --force-hash 조건 미충족: {failed}"
            raise HashMismatch(
                f"오류: 원본 DSD가 내보내기 당시와 다릅니다. (기록 {recorded[:12]}…, 실제 {actual[:12]}…) "
                "내보내기에 사용한 원본을 지정하거나 export를 다시 실행하세요." + extra
            )
        report.force_hash_used = True
        report.warnings.append(
            "경고: 원본 해시가 다르지만 --force-hash 의 3조건(크기·요소 수·태그 순서)이 맞아 진행합니다. "
            "요소별 체크섬(AC-29)이 최종 방어선입니다."
        )
    elif str(meta.get("element_count")) != str(len(index)):
        raise HashMismatch(
            f"오류: 원본의 요소 수({len(index):,})가 통합문서 기록({meta.get('element_count')})과 다릅니다. "
            "export 를 다시 실행하세요."
        )
    log(f"[2/6] 원본          contents.xml {len(raw):,} B / 요소 {len(index):,}개 / 해시 "
        f"{'일치' if not report.force_hash_used else '불일치(--force-hash)'}")

    # 3~5, 7~8. 시트별 게이트 (분류 6 은 게이트 뒤에)
    merged = {k: GateResult(k) for k in ("AC-29(3)", "AC-29", "AC-28", "AC-30")}
    for sheet in state.sheets:
        for name, res in (
            ("AC-29(3)", gate_seq_duplicates(sheet)),
            ("AC-29", gate_id_integrity(sheet, index)),
            ("AC-28", gate_monotonic(sheet)),
            ("AC-30", gate_column_overflow(sheet)),
        ):
            merged[name].violations.extend(res.violations)
            merged[name].checked += res.checked
    report.gates.update(merged)
    _raise_first(merged["AC-29(3)"])
    _raise_first(merged["AC-29"])

    # 10. 기준값 재계산 (분류에도 필요하므로 먼저 계산한다 — 순서상 게이트 뒤·혼합 게이트 앞)
    #     `_dsdx.split_notes` (1.2 WP5) 를 그대로 적용해야 섹션 구성이 통합문서와 같아진다.
    model = DocModel.build(index, raw, split_notes=state.split_notes)
    baseline = project(model)
    _check_section_map(state, baseline)

    # 6. 행 분류
    plan = plan_rows(state, model, baseline, allow_bulk_delete=allow_bulk_delete)
    report.warnings.extend(f"경고: {w}" for w in plan.warnings)
    # 7, 8
    _raise_first(merged["AC-28"])
    _raise_first(merged["AC-30"])
    # 9
    report.gates["AC-23"] = gate_bulk_delete(state)
    _raise_first(report.gates["AC-23"], allow=allow_bulk_delete)
    log(f"[3/6] 행 검증        {report.gate_line_ko()} / 유지 {len(plan.kept)} / 삽입 {len(plan.inserted)}"
        f" / 삭제 {len(plan.deleted)}")

    # 11~12. diff → 패치
    patches = PatchSet()
    _diff_and_patch(state, model, baseline, plan, index, raw, patches, report,
                    allow_lossy_mixed=allow_lossy_mixed)
    _build_insert_patches(model, plan, index, raw, patches, report)
    bare = (sum(count_bare_amp(c.new) for c in report.changes)
            + sum(count_bare_amp(t) for ins in plan.inserted for t in ins.new_texts if t))
    if bare:
        report.warnings.append(f"경고: 입력 텍스트에 `&cr;` 가 아닌 `&` 가 {bare}개 있어 `&amp;` 로 이스케이프합니다.")
    for d in plan.deleted:
        patches.add(delete_span(index, d.ordinal))
    report.deleted = len(plan.deleted)
    report.inserted = len(plan.inserted)
    _warn_extraction_and_bulk_change(model, report)
    result = patches.apply(raw)
    report.patch_stats = patches.stats()
    log(f"[4/6] 패치          {report.patch_stats['count']}개 (치환 {report.patch_stats['replacements']}, "
        f"삽입 {report.patch_stats['inserts']}, 삭제 {report.patch_stats['deletes']}) / 변경 셀 {report.changed_cells}")

    # 13. 사후 검사 (서브트리 단위, §4.10)
    check_wellformed(result)
    n_deleted = sum(subtree_size(index, d.ordinal) for d in plan.deleted)
    n_flattened = sum(subtree_size(index, o) for o in plan.planned_deletions)
    n_inserted = sum(
        subtree_size(index, ins.clone_ordinal) - sum(subtree_size(index, f) for f in ins.flattened)
        for ins in plan.inserted
    )
    expected = len(index) + n_inserted - n_deleted - n_flattened
    after = structural_equivalence(index, result, plan)
    if after != expected:
        raise PatchIntegrityError(
            f"오류: 요소 총수가 기대와 다릅니다 (기대 {expected:,} / 실제 {after:,}). 출력 파일을 만들지 않았습니다."
        )
    report.element_count_after = after
    report.bytes_after = len(result)
    log(f"[5/6] 정합성        well-formed OK / 요소 {after:,} = {len(index):,} + 삽입 {n_inserted:,} − 삭제 "
        f"{n_deleted:,} − 평탄화 {n_flattened:,} / 구조 동등성 OK")
    schema = container.schema()
    report.xsd_original = xsd_validate(raw, schema, dart_root)
    report.xsd_result = (xsd_validate(result, schema, dart_root)
                         if report.xsd_original.status != "skipped" else report.xsd_original)
    if report.xsd_original.status == "pass" and report.xsd_result.status == "fail":
        report.warnings.append(
            "경고: 되돌린 contents.xml 이 DART XSD 검증에 실패했습니다 (원본은 통과). "
            f"첫 오류: {report.xsd_result.reason}. DART 편집기에서 열어 확인하세요."
        )
    # XSD 결과 줄은 -v 와 무관하게 summary_ko() 가 항상 출력한다 (AC-25 조용한 skip 금지)

    # 14. meta.xml size → ZIP
    meta_bytes = container.meta_bytes
    if not no_meta_size and len(result) != len(raw) and (meta_file_info_size(meta_bytes) or "") != "":
        meta_bytes = update_meta_size(meta_bytes, len(result))
        report.meta_size_updated = True
    container.write(out_path, contents=result, meta=meta_bytes)
    return report


def _check_section_map(state: WorkbookState, baseline: dict) -> None:
    """통합문서에 적힌 섹션 지도와 **원본에서 다시 만든** 섹션이 같은지 본다 (1.2 WP5).

    시트 이름은 바뀔 수 있으므로(이름을 바꿔도 섹션은 유지된다) 기록된 `section_map` 자체를
    비교한다 — 기록은 export 당시의 값이고 재계산은 같은 원본·같은 `split_notes` 에서 나오므로
    둘은 반드시 같아야 한다. 다르면 `_dsdx.split_notes` 가 지워졌거나 통합문서가 다른 원본의
    것이라는 뜻이며, 그대로 진행하면 섹션과 시트의 짝이 어긋난 채 행을 지우게 된다.
    """
    recorded = [(str(e.get("sheet") or ""), str(e.get("title_id") or ""))
                for e in state.section_map]
    computed = [(e["sheet"], e["title_id"]) for e in baseline["section_map"]]
    if recorded == computed:
        return
    flag = "1" if state.split_notes else "0"
    diff = next((f"{i + 1}번째 섹션 — 기록 '{a[0]}'({a[1] or '없음'}) / 계산 '{b[0]}'({b[1] or '없음'})"
                 for i, (a, b) in enumerate(zip(recorded, computed)) if a != b),
                f"섹션 수 — 기록 {len(recorded)}개 / 계산 {len(computed)}개")
    raise InputError(
        f"오류: 통합문서의 섹션 지도(_dsdx.section_map)가 원본에서 다시 만든 섹션과 다릅니다"
        f"(split_notes={flag}). {diff}. "
        "숨김 시트 _dsdx 를 고쳤거나 다른 옵션으로 만든 통합문서입니다. 같은 옵션으로 export 를 "
        "다시 실행하세요."
    )


def _diff_and_patch(state: WorkbookState, model: DocModel, baseline: dict, plan: RowPlan,
                    index: SourceIndex, raw: bytes, patches: PatchSet, report: ImportReport, *,
                    allow_lossy_mixed: bool) -> None:
    by_eid_cell = {}
    table_of_cell = {}
    for b in model.blocks:
        if b.table is None:
            continue
        for r in b.table.rows:
            for c in r.cells:
                by_eid_cell[c.eid] = c
                table_of_cell[c.eid] = b.table
    by_eid_block = {b.eid: b for b in model.blocks}
    base_cell_out = {}
    base_block_out = {}
    for s in baseline["sheets"]:
        for r in s["rows"]:
            base = r["kind"].removesuffix(RO_SUFFIX)
            if base in ("TR", "THEAD-TR"):
                for v in r["values"]:
                    base_cell_out[v["id"]] = v["out"]
            elif base in ("P", "TITLE"):
                base_block_out[r["id"]] = r["values"][0]["out"]

    mixed_edits = []

    def flatten(ordinal: int, new_text: str, ch: CellChange) -> None:
        """--allow-lossy-mixed: 내부 구간 전체를 평문으로 치환하고 자식 서브트리를 계획된 삭제로 등록."""
        sp = index[ordinal]
        for c in sp.child_ordinals:
            plan.planned_deletions.add(c)
        promoted = merge_mixed_text(index, raw, ordinal).count("\n")
        report.warnings.append(
            f"경고: {ch.sheet} {ch.row}행 {ch.col}열 (요소 {ch.id}) 의 서식 조각 {len(sp.child_ordinals)}개를 "
            f"제거하고 평문으로 저장합니다. 원본의 들여쓰기 개행 {promoted}개가 &cr; 로 승격될 수 있습니다."
        )

    for k in plan.kept:
        sheet_name, r = k.sheet, k.row
        if k.kind in ("TR", "THEAD-TR"):
            ids = r.cell_ids
            for i, cid in enumerate(ids):
                mc = by_eid_cell.get(cid)
                if mc is None:
                    raise PatchIntegrityError(
                        f"오류: 시트 '{sheet_name}' {r.row}행 — 셀 ID {cid} 를 원본에서 찾을 수 없습니다."
                    )
                excel_val = r.value_for_cell(i)
                cur = canonical_out(excel_val)
                if cur == base_cell_out.get(cid, mc.excel_out):
                    continue
                ch = CellChange(sheet_name, r.row, _excel_col(r.col_of(i)), cid, mc.ordinal, mc.text, "",
                                "cell", mixed=mc.mixed, table_id=r.table)
                new_text = _new_text_for_cell(mc, table_of_cell[cid], excel_val, ch)
                ch.new = new_text
                ch.whitespace_only = _whitespace_only(canonical_out(mc.display), cur)
                if mc.mixed:
                    if not allow_lossy_mixed:
                        mixed_edits.append(ch)
                        continue
                    flatten(mc.ordinal, new_text, ch)
                patches.add(set_text_patch(index, mc.ordinal, escape_text(new_text).encode("utf-8")))
                report.changes.append(ch)
        else:
            b = by_eid_block[r.id]
            excel_val = r.value_at(0)
            cur = canonical_out(excel_val)
            if cur == base_block_out.get(r.id, b.excel_out):
                continue
            new_text = lf_to_cr("" if excel_val is None else str(excel_val))
            ch = CellChange(sheet_name, r.row, _excel_col(0), r.id, b.ordinal, b.text, new_text, "block",
                            mixed=b.mixed)
            ch.whitespace_only = _whitespace_only(canonical_out(b.display), cur)
            if b.mixed:
                if not allow_lossy_mixed:
                    mixed_edits.append(ch)
                    continue
                flatten(b.ordinal, new_text, ch)
            patches.add(set_text_patch(index, b.ordinal, escape_text(new_text).encode("utf-8")))
            report.changes.append(ch)
            if b.kind == "TITLE" and b.atoc == "Y":
                report.warnings.append(
                    f"경고: 시트 '{sheet_name}' 의 섹션 제목(요소 {r.id})이 바뀌었습니다. 목차(TOC)는 자동으로 "
                    "갱신되지 않으므로 DART 편집기에서 목차를 확인하세요."
                )

    if mixed_edits:
        ch = mixed_edits[0]
        raise UnsupportedEdit(MIXED_EDIT_ERROR.format(sheet=ch.sheet, row=ch.row, col=ch.col, id=ch.id))


def _check_formula_cells(state: WorkbookState, report: ImportReport) -> None:
    """수식 셀은 **마지막 계산값**으로만 되돌릴 수 있다. 캐시값이 없거나(`None`) 오류값(`#REF!` 등)이면
    원본 텍스트를 빈 문자열/오류 문자열로 덮어쓰게 되므로 코드 2 로 중단한다."""
    used = 0
    for sheet in state.sheets:
        for r in sheet.rows:
            for i in sorted(r.formulas):
                v = r.value_at(i)
                col = _excel_col(i)
                if v is None:
                    raise InputError(
                        f"오류: 시트 '{sheet.name}' {r.row}행 {col}열에 수식이 있는데 계산된 값이 저장되어 있지 "
                        "않습니다. Excel 에서 파일을 열어 다시 계산·저장(F9 후 저장)하거나, 수식을 값으로 바꾼 뒤 "
                        "다시 시도하세요."
                    )
                if isinstance(v, str) and v.startswith("#"):
                    raise InputError(
                        f"오류: 시트 '{sheet.name}' {r.row}행 {col}열의 수식 결과가 오류값({v})입니다. "
                        "수식을 고치거나 값으로 바꾼 뒤 다시 시도하세요."
                    )
                used += 1
    if used:
        report.warnings.append(
            f"경고: 수식이 든 셀 {used}개는 마지막 계산값으로 되돌립니다(DSD 에는 수식이 저장되지 않습니다)."
        )


def _build_insert_patches(model: DocModel, plan: RowPlan, index: SourceIndex, raw: bytes,
                          patches: PatchSet, report: ImportReport) -> None:
    """삽입 행 → 앵커(복제 원본) 바이트 복제 패치 (§4.8, §4.9, AC-7, AC-24)."""
    if not plan.inserted:
        return
    by_eid_cell = {}
    table_of_cell = {}
    for b in model.blocks:
        if b.table is None:
            continue
        for r in b.table.rows:
            for c in r.cells:
                by_eid_cell[c.eid] = c
                table_of_cell[c.eid] = b.table
    by_eid_block = {b.eid: b for b in model.blocks}

    for ins in plan.inserted:
        r = ins.row
        where = f"시트 '{ins.sheet}' {r.row}행"
        if ins.kind in ("TR", "THEAD-TR"):
            clone_sp = index[ins.clone_ordinal]
            cells = [c for c in clone_sp.child_ordinals if index[c].tag in CELL_TAGS]
            # 앵커의 열 배치(`SheetRow.cols`) 밖에 값이 있으면 짝지을 셀이 없다 (AC-30 과 같은 판정).
            stray = r.unmapped_value_col(len(cells))
            if stray >= 0:
                raise UnsupportedEdit(
                    f"오류: {where} — 삽입 행의 값이 셀이 없는 열({_excel_col(stray)}열)에 있습니다"
                    f"(앵커 행의 셀 {len(cells)}개). 이 도구는 열 추가·삽입을 지원하지 않습니다."
                )
            new_texts = []
            keep = set()
            for i, c in enumerate(cells):
                mc = by_eid_cell[index.eid_of(c)]
                v = r.value_for_cell(i)
                cur = canonical_out(v)
                if mc.mixed and cur == mc.excel_out:
                    keep.add(c)
                    new_texts.append(None)
                    continue
                if mc.mixed:
                    report.warnings.append(
                        f"경고: {where} {_excel_col(r.col_of(i))}열 — 복제 원본 셀(요소 {mc.eid})의 서식 조각을 새 행에서는 "
                        "평문으로 대체했습니다(새 행이라 원본 손실은 없습니다)."
                    )
                if v is None:
                    new_texts.append("")
                elif _is_number(v):
                    d = Decimal(str(v)) if not isinstance(v, Decimal) else v
                    res = resolve_pattern(mc, table_of_cell[mc.eid], negative=d < 0)
                    new_texts.append(format_number(d, res.pattern))
                    if res.level != "self":
                        report.changes.append(CellChange(ins.sheet, r.row, _excel_col(r.col_of(i)), "",
                                                         mc.ordinal, "", new_texts[-1], "insert",
                                                         inherited=res, table_id=ins.table_id))
                else:
                    new_texts.append(lf_to_cr(str(v)))
            payload, flattened = clone_tr_bytes(index, raw, ins.clone_ordinal, new_texts, keep)
            ins.flattened = flattened
            ins.new_texts = new_texts
        elif ins.kind == "P":
            b = by_eid_block[index.eid_of(ins.clone_ordinal)]
            v = r.value_at(0)
            cur = canonical_out(v)
            if b.mixed:
                if cur != b.excel_out:
                    raise UnsupportedEdit(
                        f"오류: {where} — 서식 조각(SPAN)을 포함한 문단(요소 {b.eid}, P:RO)을 복제하면서 "
                        "텍스트를 바꾸는 것은 지원하지 않습니다. 값을 그대로 둔 순수 복제만 허용됩니다."
                    )
                payload = clone_p_bytes(index, raw, ins.clone_ordinal, None)
            else:
                payload = clone_p_bytes(index, raw, ins.clone_ordinal, lf_to_cr("" if v is None else str(v)))
        else:
            raise UnsupportedEdit(f"오류: {where} — {ins.kind} 행은 삽입할 수 없습니다.")
        patches.add(insert_patch(index, ins, payload))


def _new_text_for_cell(mc, table, excel_val, ch: CellChange) -> str:
    """Excel 값 → 되돌릴 원문 (§4.5 복원 규칙, AC-13)."""
    if excel_val is None:
        return ""
    if _is_number(excel_val):
        d = Decimal(str(excel_val)) if not isinstance(excel_val, Decimal) else excel_val
        res = resolve_pattern(mc, table, negative=d < 0)
        ch.inherited = res
        return format_number(d, res.pattern)
    return lf_to_cr(str(excel_val))


def _warn_extraction_and_bulk_change(model: DocModel, report: ImportReport) -> None:
    """AC-27 (추출 대상 표 변경) + R-23 (한 표 50% 초과 변경)."""
    tables = {b.table.eid: (b, b.table) for b in model.blocks if b.table is not None}
    sheet_of_table = {}
    for s in model.sections:
        for t in s.tables:
            sheet_of_table[t.eid] = s.sheet_name
    changed_by_table: dict = {}
    for c in report.changes:
        if c.kind == "cell" and c.table_id and not c.whitespace_only:
            changed_by_table.setdefault(c.table_id, []).append(c)
    extraction_n = 0
    extraction_tables = []
    for tid, changes in changed_by_table.items():
        if tid not in tables:          # 숨김 _TABLE 이 훼손된 경우 — 경고 집계만 건너뛴다 (게이트가 이미 판정)
            continue
        block, table = tables[tid]
        if table.extraction:
            extraction_n += len(changes)
            extraction_tables.append(f"{sheet_of_table.get(tid, '?')} (원본 {table.line}행, {tid})")
        ncells = sum(len(r.cells) for r in table.rows)
        if ncells and len(changes) * 2 > ncells:
            report.warnings.append(
                f"경고: 시트 '{sheet_of_table.get(tid, '?')}' 의 표(원본 {table.line}행, {tid})에서 셀 {ncells}개 중 "
                f"{len(changes)}개가 바뀌었습니다(50% 초과). 정렬이나 열 이동으로 값의 짝이 어긋나지 않았는지 확인하세요."
            )
    if extraction_n:
        report.warnings.append(EXTRACTION_WARNING.format(n=extraction_n, tables=", ".join(extraction_tables)))


__all__ = [
    "EXTRACTION_WARNING",
    "MIXED_EDIT_ERROR",
    "CellChange",
    "ImportReport",
    "default_output_path",
    "import_",
]
