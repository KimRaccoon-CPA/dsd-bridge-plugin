"""`dsd-bridge clone` — 표·주석 섹션을 복제해 늘리기 (0.7.0).

DSD브릿지는 표·섹션을 새로 **만들지는** 못한다. 대신 이미 있는 표나 주석 하나를 **바이트 그대로 복제**해
바로 뒤에 붙인다. 서식 속성(열 너비·계정코드 칸·표 그룹)은 검증된 원본에서 물려받으므로 안전하고, 복제한 뒤
export 해 내용을 바꾸면 된다.

- `clone 원본.dsd -o 새.dsd --table <표 ID> [--count N] [--with-prev K]` — 표를 N번 복제해 바로 뒤에.
  `--with-prev K` 는 표 앞의 형제 요소 K개(제목 문단·단위 표 등)를 함께 복제한다.
  표 ID 는 `dsd-bridge describe` 의 `### 표 n — \\`E…\\`` 에 나온다. `--sheet 이름 --index n` 으로도 고른다.
- `clone 원본.dsd -o 새.dsd --note N [--as M]` — 주석 N(번호 제목 문단부터 다음 번호 제목 직전까지)을 복제해
  바로 뒤에 붙이고, 복제본 제목의 번호를 M(기본 N+1)으로 바꾼다. 뒤 주석 번호와 재무제표의 `(주석N)` 참조는
  사람이 고친다(`check` 의 확인 필요가 알려 준다).

결과는 well-formed·XSD 검사를 거쳐 새 이름으로만 쓴다(원본 불변). 원본이 XSD 를 통과하는데 결과가 실패하면
쓰지 않는다(종료 코드 6).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from .container import DsdContainer, meta_file_info_size, update_meta_size
from .errors import InputError, OutputError, PatchIntegrityError
from .model import DocModel, note_heading
from .text import escape_text
from .validate import check_wellformed, xsd_validate
from .xmlsplice import SourceIndex


@dataclass
class CloneReport:
    source: Path
    output: Path
    what: str
    copies: int
    elements: int
    xsd: str = ""

    def lines_ko(self) -> list:
        return [f"복제 완료: {self.what} → {self.copies}개 복제 (요소 {self.elements}개씩) → {self.output}",
                f"XSD 검증: {self.xsd}",
                "다음: 새 파일을 export 해 복제한 표·주석의 내용을 바꾸세요. 주석을 늘렸으면 뒤 주석 번호와 "
                "재무제표의 (주석N) 참조도 고치고 `check` 의 확인 필요로 확인하세요."]


def _read(path):
    p = Path(path)
    if not p.is_file():
        raise InputError(f"오류: 파일을 찾을 수 없습니다: {p}")
    c = DsdContainer.read(p)
    raw = c.contents_bytes
    return c, raw, SourceIndex.build(raw)


def _write(container, raw, result, out) -> str:
    check_wellformed(result)
    meta = container.meta_bytes
    if len(result) != len(raw) and (meta_file_info_size(meta) or "") != "":
        meta = update_meta_size(meta, len(result))
    schema = container.schema()
    v = xsd_validate(result, schema)
    if v.status == "fail" and xsd_validate(raw, schema).status == "pass":
        raise PatchIntegrityError(f"오류: 복제한 결과가 DART XSD 검증에 실패해(원본은 통과) 파일을 만들지 않았습니다: {v.reason}. "
                                  "복제할 범위(--with-prev)를 줄이거나 다른 표·주석을 고르세요.")
    try:
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        container.write(out, contents=result, meta=meta)
    except OSError as exc:
        raise OutputError(f"오류: 출력 파일을 쓸 수 없습니다: {out}\n  사유: {exc}") from exc
    return v.status_ko + (f" ({v.reason})" if v.status not in ("pass", "skipped") and v.reason else "")


def _subtree_size(index, o: int) -> int:
    n = 1
    for c in index[o].child_ordinals:
        n += _subtree_size(index, c)
    return n


def _find_table(index, model, table: str, sheet: str, nth: int) -> int:
    if table:
        for sp in index.spans:
            if sp.tag == "TABLE" and index.eid_of(sp.ordinal) == table:
                return sp.ordinal
        raise InputError(f"오류: 표 ID {table} 를 찾을 수 없습니다(dsd-bridge describe 의 표 ID 를 쓰세요).")
    for section in model.sections:
        if section.sheet_name == sheet:
            tables = [b for b in section.blocks if b.table is not None]
            if not 1 <= nth <= len(tables):
                raise InputError(f"오류: 시트 '{sheet}' 의 표는 {len(tables)}개입니다(--index 1~{len(tables)}).")
            return tables[nth - 1].ordinal
    raise InputError(f"오류: 시트 '{sheet}' 를 찾을 수 없습니다.")


def clone_table(path, out, *, table: str = "", sheet: str = "", index_no: int = 0, count: int = 1,
                with_prev: int = 0) -> CloneReport:
    src, dst = Path(path), Path(out)
    if dst.resolve() == src.resolve():
        raise OutputError("오류: 출력 경로가 원본과 같습니다.")
    if count < 1:
        raise InputError("오류: --count 는 1 이상이어야 합니다.")
    container, raw, index = _read(src)
    model = DocModel.build(index, raw)
    o = _find_table(index, model, table, sheet, index_no)
    sp = index[o]
    first = o
    if with_prev:
        sibs = index[sp.parent_ordinal].child_ordinals
        k = sibs.index(o)
        if k < with_prev:
            raise InputError(f"오류: 표 앞에 형제 요소가 {k}개뿐입니다(--with-prev {with_prev}).")
        first = sibs[k - with_prev]
        odd = [index[x].tag for x in sibs[k - with_prev:k] if index[x].tag not in ("P", "TABLE")]
        if odd:
            raise InputError(f"오류: 표 앞 요소에 문단·표가 아닌 것({', '.join(odd)})이 있어 함께 복제할 수 없습니다. "
                             "--with-prev 를 줄이세요.")
    start = index.leading_ws_start(first)
    chunk = raw[start:sp.end_gt]
    result = raw[:sp.end_gt] + chunk * count + raw[sp.end_gt:]
    n = sum(_subtree_size(index, x) for x in index[sp.parent_ordinal].child_ordinals
            if index[first].start_lt <= index[x].start_lt <= sp.start_lt)
    what = f"표 {index.eid_of(o)}" + (f" (앞 요소 {with_prev}개 포함)" if with_prev else "")
    report = CloneReport(src, dst, what, count, n)
    report.xsd = _write(container, raw, result, dst)
    return report


def clone_note(path, out, note: int, new_no: int | None = None) -> CloneReport:
    src, dst = Path(path), Path(out)
    if dst.resolve() == src.resolve():
        raise OutputError("오류: 출력 경로가 원본과 같습니다.")
    container, raw, index = _read(src)
    model = DocModel.build(index, raw)
    for section in model.sections:
        blocks = section.blocks
        heads = [(i, note_heading(b.text)) for i, b in enumerate(blocks) if b.table is None]
        heads = [(i, h) for i, h in heads if h is not None]
        hit = [k for k, (i, h) in enumerate(heads) if h[0] == note]
        if not hit:
            continue
        k = hit[0]
        i0 = heads[k][0]
        i1 = heads[k + 1][0] if k + 1 < len(heads) else len(blocks)
        group = blocks[i0:i1]
        parents = {index[b.ordinal].parent_ordinal for b in group}
        if len(parents) != 1:
            raise InputError(f"오류: 주석 {note} 의 문단·표가 한 묶음 안에 있지 않아 복제할 수 없습니다.")
        first, last = index[group[0].ordinal], index[group[-1].ordinal]
        start = index.leading_ws_start(first.ordinal)
        chunk = raw[start:last.end_gt]
        target = new_no if new_no is not None else note + 1
        head = first
        inner = raw[head.start_gt:head.end_lt]
        m = re.search(rb"(^|[\s>])(" + str(note).encode() + rb")(\.\s)", inner)
        if not m:
            raise InputError(f"오류: 주석 {note} 제목에서 번호를 찾지 못했습니다.")
        rel = (head.start_gt - start) + m.start(2)
        chunk = chunk[:rel] + str(target).encode() + chunk[rel + len(m.group(2)):]
        result = raw[:last.end_gt] + chunk + raw[last.end_gt:]
        n = sum(_subtree_size(index, b.ordinal) for b in group)
        title = " ".join(group[0].text.split("&cr;")[0].split())[:30]
        report = CloneReport(src, dst, f"주석 {note} '{escape_text(title)}' → 주석 {target}", 1, n)
        report.xsd = _write(container, raw, result, dst)
        return report
    raise InputError(f"오류: 주석 {note} 의 번호 제목(`{note}. 제목`)을 찾을 수 없습니다.")


__all__ = ["CloneReport", "clone_table", "clone_note"]
