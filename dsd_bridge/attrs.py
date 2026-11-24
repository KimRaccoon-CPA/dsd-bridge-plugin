"""숨김 메타데이터 시트 `_dsdx_attrs` (계획 1.2 WP2).

표 행(`TR`)마다 원문의 DART 속성(`ALEVEL`/`ACODE`/`ADELIM`/`AUTOSUM`/`AUNIT`/`ACLASS`)을
한 줄로 적어 두는 **읽기 전용 참고 시트**다. 용도는 `dsd-bridge check`(WP4)의 합계검증이며,

- **import 는 이 시트를 읽지 않는다.** 지워도 되돌리기에 아무 영향이 없다.
- 게이트도 보지 않는다 (`sheetstate` 가 데이터 시트에서 제외한다).
- 기존 A~E 숨김 열과 F 열 데이터 시작은 그대로다 (게이트 표면 불변).

열 순서는 `dsd_bridge.model.ATTR_KEYS` 와 같고, 앞에 `_ID`(행 요소 ID) / `sheet` / `table`(표 요소 ID)
가 붙는다. 값은 전부 문자열이다.
"""

from __future__ import annotations

from pathlib import Path

from .errors import InputError
from .model import ATTR_KEYS

#: 숨김 속성 시트 이름.
ATTRS_SHEET = "_dsdx_attrs"
#: 1행(머리글) — `_dsdx_attrs` 의 열 순서.
ATTRS_HEADER = ("_ID", "sheet", "table") + ATTR_KEYS


def attrs_rows(proj: dict) -> list:
    """`project()` 결과에서 `_dsdx_attrs` 의 본문 행을 만든다 (시트 순서 → 행 순서)."""
    out = []
    for sheet in proj["sheets"]:
        for row in sheet["rows"]:
            attrs = row.get("attrs")
            if not attrs:
                continue
            out.append([row["id"], sheet["name"], row["table"]]
                       + [str(attrs.get(k, "")) for k in ATTR_KEYS])
    return out


def read_attrs_sheet(workbook_path) -> list:
    """통합문서의 `_dsdx_attrs` 를 읽어 `ATTRS_HEADER` 키의 dict 목록으로 준다.

    시트가 없으면 **빈 목록**이다 — 이 시트는 없어도 되는 참고 자료이므로 오류가 아니다.
    통합문서 자체를 열 수 없을 때만 `InputError`(코드 2).
    """
    from openpyxl import load_workbook

    p = Path(workbook_path)
    if not p.is_file():
        raise InputError(f"오류: 통합문서를 찾을 수 없습니다: {p}")
    try:
        wb = load_workbook(p, data_only=True, read_only=True)
    except Exception as exc:  # openpyxl 은 BadZipFile/KeyError/InvalidFileException 등을 던진다
        raise InputError(
            f"오류: Excel 통합문서(.xlsx)로 열 수 없습니다: {p}\n  사유: {type(exc).__name__}: {exc}"
        ) from exc
    try:
        if ATTRS_SHEET not in wb.sheetnames:
            return []
        out = []
        for i, tup in enumerate(wb[ATTRS_SHEET].iter_rows(values_only=True)):
            if i == 0:                     # 머리글 행
                continue
            vals = list(tup) + [None] * (len(ATTRS_HEADER) - len(tup))
            out.append({k: ("" if v is None else str(v))
                        for k, v in zip(ATTRS_HEADER, vals)})
        return out
    finally:
        wb.close()


__all__ = ["ATTRS_SHEET", "ATTRS_HEADER", "attrs_rows", "read_attrs_sheet"]
