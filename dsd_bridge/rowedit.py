"""행 분류와 행 편집 (계획 §4.4 행 분류 표·앵커 결정 규칙, §4.8, §4.9).

- `plan_rows` — 통합문서의 행을 **유지 / 삽입 / 복제 삽입 / 삭제** 로 분류해 `RowPlan` 을 만든다
  (단계 5 골격). 게이트(AC-28/29/30)는 이보다 앞에서 돌아 형식이 보장된 상태를 전제한다.
- `delete_span` — 행/블록 하나를 지우는 패치 (`[선행공백_시작, end_gt)`).
- `clone_tr_bytes` / `clone_p_bytes` — 앵커 바이트 복제 삽입 (단계 6).

행 분류 규칙 (정본, §4.4)::

    _SEQ 값 있음, 유일 | _ID 값 있음 → 원본 행 (유지)
    _SEQ 빈 값        | _ID 빈 값   → 평범한 삽입 행 — 앵커 = 바로 위 행
    _SEQ 빈 값        | _ID 값 있음 → 복제 삽입 행 (AC-24) — 복제 원본 = 그 요소
    _SEQ 중복                       → 채우기 핸들 사고 (AC-29(3), 게이트가 먼저 잡는다)
    _SEQ 값 있음      | _ID 빈 값   → 손상 (AC-29 게이트, 코드 6)
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .errors import SheetIntegrityError, UnsupportedEdit
from .gates import EID_RE
from .model import CELL_TAGS, RO_SUFFIX
from .sheetstate import MAPPED_KINDS, SheetRow, WorkbookState
from .text import escape_text
from .xmlsplice import Patch, SourceIndex, set_attr_value


@dataclass
class KeptRow:
    sheet: str
    row: SheetRow
    ordinal: int
    kind: str                    # TR / THEAD-TR / P / TITLE (:RO 제거)


@dataclass
class InsertedRow:
    sheet: str
    row: SheetRow
    kind: str                    # 상속한 종류 (TR / THEAD-TR / P)
    anchor_ordinal: int          # 위치 앵커 — 이 요소 뒤(before=True 면 앞)에 넣는다
    clone_ordinal: int           # 바이트 복제 원본 (평범한 삽입이면 == anchor_ordinal)
    before: bool = False         # TABLE-START 아래 삽입: 표의 첫 원본 행 **앞**
    table_id: str = ""
    duplicate: bool = False      # AC-24 복제 삽입 (_ID 있음)
    flattened: set = field(default_factory=set)   # 사본에서 평문화된 자식(혼합 셀의 자식) 서브트리 루트 순번
    new_texts: list = field(default_factory=list) # 사본 셀에 넣은 원문 (진단용)
    attr_overrides: dict = field(default_factory=dict)  # 0.3.0 사본에서 속성을 바꾼 요소: 원본 순번 → attr_bytes

    @property
    def position_ordinal(self) -> int:
        return self.anchor_ordinal


@dataclass
class DeletedRow:
    sheet: str
    ordinal: int
    kind: str
    seq: int | None = None
    whole_table: bool = False    # TABLE 블록 통째 삭제 (시트 통삭제)


@dataclass
class RowPlan:
    kept: list = field(default_factory=list)
    inserted: list = field(default_factory=list)
    deleted: list = field(default_factory=list)
    planned_deletions: set = field(default_factory=set)   # 평탄화로 사라지는 자식 서브트리 루트 순번
    attr_overrides: dict = field(default_factory=dict)    # 0.3.0 속성값을 바꾼 원본 요소: 순번 → attr_bytes
    warnings: list = field(default_factory=list)

    def deleted_ordinals(self) -> set:
        return {d.ordinal for d in self.deleted}

    @property
    def kept_by_ordinal(self) -> dict:
        return {k.ordinal: k for k in self.kept}


def eid_ordinal(eid: str) -> int:
    m = EID_RE.match(eid)
    if not m:
        raise ValueError(f"요소 ID 형식이 아닙니다: {eid!r}")
    return int(m.group(1))


# ─────────────────────────────────────────────────────────────────────────────
# 분류
# ─────────────────────────────────────────────────────────────────────────────


def _row_ref(sheet: str, r: SheetRow) -> str:
    return f"시트 '{sheet}' {r.row}행"


def _section_sheet(state: WorkbookState, entry: dict):
    """`section_map` 항목의 시트 — `title_id` 기준 (AC-23). 짝이 어긋나면 코드 4 로 중단한다.

    행 분류(6)가 AC-23 게이트(9)보다 먼저 돌므로 게이트와 같은 문구를 여기서 낸다.
    """
    found = state.section_sheet(entry)
    if found.conflict:
        raise SheetIntegrityError(found.conflict)
    return found.sheet


def _anchor_for_insert(sheet_name: str, rows: list, i: int, deleted_eids: set, model_rows: dict,
                       title_ids: set = frozenset()):
    """§4.4 앵커 결정 규칙. 반환 (anchor_ordinal, kind, table_id, before).

    `title_ids` 는 `section_map` 의 섹션 제목 요소 ID 다 — `--split-notes` (1.2 WP5) 로 나눈
    주석 시트의 제목은 `TITLE` 이 아니라 `P` 라서, 그 아래 삽입도 `ATOC="Y"` 제목과 똑같이
    거절하려면 종류가 아니라 ID 로 봐야 한다.
    """
    j = i - 1
    while j >= 0:
        above = rows[j]
        base = above.base_kind
        if above.is_header:
            raise UnsupportedEdit(
                f"오류: {_row_ref(sheet_name, rows[i])} — 시트 맨 위에는 행을 넣을 수 없습니다."
            )
        if base == "TABLE-END":
            raise UnsupportedEdit(
                f"오류: {_row_ref(sheet_name, rows[i])} — 표 밖에는 행을 넣을 수 없습니다. "
                "표 안의 기존 행 아래에 삽입하세요."
            )
        if base == "TABLE-START":
            # 그 표의 첫 원본 행 앞에 삽입 — 종류·_TABLE 은 첫 원본 행에서 상속
            first = next((mr for mr in model_rows.get(above.table, []) ), None)
            if first is None:
                raise UnsupportedEdit(
                    f"오류: {_row_ref(sheet_name, rows[i])} — 이 표에는 원본 행이 없어 앞에 넣을 수 없습니다."
                )
            return first["ordinal"], first["kind"], above.table, True
        if base == "TITLE" or (above.seq is not None and above.id and above.id in title_ids):
            raise UnsupportedEdit(
                f"오류: {_row_ref(sheet_name, rows[i])} — 제목 아래에 새 블록을 만드는 것은 섹션 구조 변경이라 "
                "지원하지 않습니다. 기존 문단이나 표의 행 아래에 삽입하세요."
            )
        if above.seq is None:
            # 위 행도 삽입 행 — 그 위로 계속 올라간다 (같은 앵커 뒤에 순서대로 붙는다)
            j -= 1
            continue
        if base in ("TR", "THEAD-TR", "P"):
            if above.id in deleted_eids:
                j -= 1
                continue
            return eid_ordinal(above.id), base, above.table, False
        j -= 1
    raise UnsupportedEdit(
        f"오류: {_row_ref(sheet_name, rows[i])} — 삽입 위치 위쪽에 유지되는 원본 행이 없습니다."
    )


def plan_rows(state: WorkbookState, model, baseline: dict, *, allow_bulk_delete: bool = False) -> RowPlan:
    """행 분류. `baseline` 은 `project(model)` 결과.

    - `section_map` 을 순회하며 섹션의 시트를 **`title_id`** 로 찾는다(이름이 아니다 — 시트 이름을
      바꿔도 섹션이 유지된다). 시트가 없으면 그 섹션은 100% 소실이며, `allow_bulk_delete` 일 때만
      본문 블록(P / TABLE / ATOC="N" 제목)을 삭제하고 **섹션 제목은 남긴다**(§4.9: `ATOC="Y"` 제목 삭제 금지).
    - 통합문서에 있는 시트라도 `section_map` 에 없는 시트는 무시한다(사용자가 추가한 메모 시트 등).
    """
    plan = RowPlan()
    #: 섹션 제목 요소 ID — 삭제 금지·삽입 앵커 금지의 기준 (`--split-notes` 의 주석 제목 포함).
    title_ids = {str(e.get("title_id") or "") for e in state.section_map if e.get("title_id")}
    # 원본 매핑 행: 시트별 → [row dict], eid → row dict
    base_sheets = {s["name"]: s for s in baseline["sheets"]}
    base_rows_by_eid: dict = {}
    for s in baseline["sheets"]:
        for r in s["rows"]:
            if r["kind"].removesuffix(RO_SUFFIX) in MAPPED_KINDS:
                base_rows_by_eid[r["id"]] = r
    # 표별 원본 행 목록 (TABLE-START 앵커용)
    model_rows_by_table: dict = {}
    for s in baseline["sheets"]:
        for r in s["rows"]:
            if r["kind"] in ("TR", "THEAD-TR"):
                model_rows_by_table.setdefault(r["table"], []).append(
                    {"ordinal": eid_ordinal(r["id"]), "kind": r["kind"]}
                )

    kept_eids: set = set()
    kept_rows = []
    insert_candidates = []
    for entry in state.section_map:
        sheet = _section_sheet(state, entry)
        if sheet is None:
            continue
        name = sheet.name
        for i, r in enumerate(sheet.rows):
            if r.is_header or r.is_guide:
                continue
            if r.seq is not None:
                if r.base_kind not in MAPPED_KINDS or not r.id:
                    continue        # 게이트가 이미 걸렀거나 무관한 행
                kept_eids.add(r.id)
                kept_rows.append(KeptRow(name, r, eid_ordinal(r.id), r.base_kind))
            else:
                if not r.id and not r.kind and not r.table and not r.cells and not r.values:
                    continue        # 완전히 빈 행은 삽입이 아니다 (Excel 이 남기는 빈 줄)
                insert_candidates.append((name, sheet, i, r))
    plan.kept = kept_rows

    # 삭제 — 원본 매핑 행 중 어느 시트에도 없는 것 (section_map 순회)
    for entry in state.section_map:
        base = base_sheets.get(entry["sheet"])
        if base is None:
            continue
        sheet = _section_sheet(state, entry)
        name = sheet.name if sheet is not None else entry["sheet"]
        mapped = [r for r in base["rows"] if r["kind"].removesuffix(RO_SUFFIX) in MAPPED_KINDS]
        missing = [r for r in mapped if r["id"] not in kept_eids]
        if not missing:
            continue
        if sheet is None:
            if not allow_bulk_delete:
                # 파이프라인상 행 분류(6)가 AC-23 게이트(9)보다 먼저 돌므로 여기서 같은 판정을 낸다
                raise SheetIntegrityError(
                    f"오류: 시트 '{name}' — 시트가 없습니다(50% 초과). "
                    "의도한 삭제라면 --allow-bulk-delete 를 사용하세요."
                )
            # 시트 통삭제: 제목은 남기고 본문 블록을 통째로 지운다
            title_id = entry["title_id"]
            tables_done = set()
            for r in mapped:
                if r["id"] == title_id:
                    continue
                if r["kind"] in ("TR", "THEAD-TR"):
                    if r["table"] in tables_done:
                        continue
                    tables_done.add(r["table"])
                    plan.deleted.append(DeletedRow(name, eid_ordinal(r["table"]), "TABLE", whole_table=True))
                else:
                    plan.deleted.append(DeletedRow(name, eid_ordinal(r["id"]), r["kind"].removesuffix(RO_SUFFIX), r["seq"]))
            plan.warnings.append(
                f"시트 '{name}' 이(가) 없어 그 섹션의 본문 블록을 삭제했습니다(표 {len(tables_done)}개 포함). "
                "섹션 제목은 유지됩니다."
            )
            continue
        for r in missing:
            kind = r["kind"].removesuffix(RO_SUFFIX)
            # 섹션 제목은 종류가 아니라 **ID** 로 가린다 — `--split-notes` 의 주석 제목은 `P` 다.
            if r["id"] == entry["title_id"]:
                raise UnsupportedEdit(
                    f"오류: 시트 '{name}' — 섹션 제목 행(요소 {r['id']})은 삭제할 수 없습니다. "
                    "섹션 자체가 사라지므로 이 도구는 지원하지 않습니다. Ctrl+Z 로 되돌리세요."
                )
            plan.deleted.append(DeletedRow(name, eid_ordinal(r["id"]), kind, r["seq"]))

    # 삽입 — 앵커 결정 (삭제로 분류된 행은 앵커가 될 수 없다 → 위로 계속 올라간다)
    deleted_eids = {model.index.eid_of(d.ordinal) for d in plan.deleted if not d.whole_table}
    for name, sheet, i, r in insert_candidates:
        anchor, kind, table_id, before = _anchor_for_insert(
            name, sheet.rows, i, deleted_eids, model_rows_by_table, title_ids,
        )
        if r.id:
            # 복제 삽입 (AC-24) — 복제 원본은 _ID 의 요소
            try:
                clone = eid_ordinal(r.id)
            except ValueError as exc:
                raise UnsupportedEdit(
                    f"오류: {_row_ref(name, r)} — 삽입 행의 요소 ID 형식이 올바르지 않습니다({r.id!r}).",
                    exit_code=6,
                ) from exc
            base_row = base_rows_by_eid.get(r.id)
            if base_row is None:
                raise UnsupportedEdit(
                    f"오류: {_row_ref(name, r)} — 복제 원본 요소 {r.id} 를 원본에서 찾을 수 없습니다.",
                    exit_code=4,
                )
            kind = base_row["kind"].removesuffix(RO_SUFFIX)
            if base_row["table"] and table_id and base_row["table"] != table_id:
                raise UnsupportedEdit(
                    f"오류: {_row_ref(name, r)} — 다른 표의 행(요소 {r.id})을 이 표에 복제할 수 없습니다."
                )
            plan.inserted.append(InsertedRow(name, r, kind, anchor, clone, before,
                                             base_row["table"] or table_id, duplicate=True))
        else:
            if kind == "TITLE":
                raise UnsupportedEdit(
                    f"오류: {_row_ref(name, r)} — 제목 아래에 새 블록을 만드는 것은 지원하지 않습니다."
                )
            plan.inserted.append(InsertedRow(name, r, kind, anchor, anchor, before, table_id))
    return plan


# ─────────────────────────────────────────────────────────────────────────────
# 패치 생성
# ─────────────────────────────────────────────────────────────────────────────


def delete_span(index: SourceIndex, ordinal: int) -> Patch:
    """행/블록 하나를 지운다 — `[선행공백_시작, end_gt)` (§4.8). 앞의 빈 줄 하나까지 함께 지워
    삭제 후 빈 줄이 중복되지 않는다."""
    sp = index[ordinal]
    return Patch(index.leading_ws_start(ordinal), sp.end_gt, b"", label=f"{sp.tag}#{ordinal} 삭제")


def _open_tag_for_text(index: SourceIndex, ordinal: int) -> bytes:
    """자기닫힘 `<X/>` 면 `<X>` 로 바꾼 시작 태그 (AC-22), 아니면 원본 시작 태그."""
    sp = index[ordinal]
    tag = index.raw[sp.start_lt : sp.start_gt]
    if sp.self_closing:
        return tag[:-2].rstrip() + b">"
    return tag


def clone_tr_bytes(index: SourceIndex, raw: bytes, anchor: int, new_texts: list,
                   keep_verbatim: set | None = None, attr_values: dict | None = None) -> tuple:
    """앵커 `TR` 의 `[선행공백_시작, end_gt)` 를 복사하고 셀 내부 텍스트만 갈아끼운다 (§4.8, AC-7).

    - `new_texts[i]` 는 i 번째 셀의 **원문**(escape 전, `&cr;` 유지). `None` 이면 빈 문자열.
      앵커보다 짧으면 나머지 셀은 빈 문자열, 길면 초과분은 버린다(호출 측이 경고).
    - `keep_verbatim` 에 든 셀 순번은 내부 바이트를 **그대로** 복사한다 (혼합 셀의 값이 바뀌지 않은
      경우 — 자식 요소가 보존된다).
    - 혼합 셀에 새 텍스트를 넣으면 자식이 사라진다. 사라지는 자식 순번을 함께 돌려준다.

    반환: `(bytes, flattened_child_ordinals)`. 속성(`ACOPY`/`ADELETE`/`WIDTH`/`HEIGHT`/`ACODE`/`ADELIM`/
    `ALEVEL`/`USERMARK`/`ACOPYCOL`/`ADELETECOL`/`AMOVECOL` …)은 바이트 단위로 복제된다.
    `attr_values` = `{셀 순번: {속성: 값}}` 이면 그 셀 시작 태그의 속성값만 바꿔 넣는다 (0.3.0 날짜 칸).
    """
    attr_values = attr_values or {}

    def fix(c: int, tag: bytes) -> bytes:
        for name, value in attr_values.get(c, {}).items():
            tag = set_attr_value(tag, name, value)
        return tag

    keep_verbatim = keep_verbatim or set()
    sp = index[anchor]
    cells = [c for c in sp.child_ordinals if index[c].tag in CELL_TAGS]
    base = index.leading_ws_start(anchor)
    out = []
    pos = base
    flattened = set()
    for i, c in enumerate(cells):
        cs = index[c]
        out.append(raw[pos : cs.start_lt])
        if c in keep_verbatim:
            out.append(fix(c, raw[cs.start_lt : cs.start_gt]) + raw[cs.start_gt : cs.end_gt])
        else:
            text = new_texts[i] if i < len(new_texts) and new_texts[i] is not None else ""
            out.append(fix(c, _open_tag_for_text(index, c)))
            out.append(escape_text(text).encode("utf-8"))
            out.append(b"</" + cs.tag.encode("utf-8") + b">")
            flattened.update(cs.child_ordinals)
        pos = cs.end_gt
    out.append(raw[pos : sp.end_gt])
    return b"".join(out), flattened


def clone_p_bytes(index: SourceIndex, raw: bytes, anchor: int, new_text: str | None) -> bytes:
    """앵커 `P` 를 복제하고 내부 텍스트만 교체한다 (§4.9). `new_text=None` 이면 바이트 그대로 복제
    (혼합 `P:RO` 의 순수 복제)."""
    sp = index[anchor]
    base = index.leading_ws_start(anchor)
    if new_text is None:
        return raw[base : sp.end_gt]
    return (raw[base : sp.start_lt] + _open_tag_for_text(index, anchor)
            + escape_text(new_text).encode("utf-8") + b"</" + sp.tag.encode("utf-8") + b">")


def insert_patch(index: SourceIndex, ins: InsertedRow, payload: bytes) -> Patch:
    """§4.8 경계 규칙 1 — 삽입은 앵커 `end_gt` 의 길이 0 구간. `before` 면 앵커의 선행 공백 시작 위치."""
    sp = index[ins.anchor_ordinal]
    if ins.before:
        at = index.leading_ws_start(ins.anchor_ordinal)
    else:
        at = sp.end_gt
    return Patch(at, at, payload, label=f"{sp.tag}#{ins.anchor_ordinal} 뒤 삽입" if not ins.before
                 else f"{sp.tag}#{ins.anchor_ordinal} 앞 삽입")


__all__ = [
    "KeptRow",
    "InsertedRow",
    "DeletedRow",
    "RowPlan",
    "eid_ordinal",
    "plan_rows",
    "delete_span",
    "clone_tr_bytes",
    "clone_p_bytes",
    "insert_patch",
]
