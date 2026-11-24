"""Excel 시트 이름 규칙 (계획 AC-2′, 7단계).

1. 원문 그대로
2. Excel 금지문자 ``: \\ / ? * [ ]`` → ``_``
3. 앞뒤 작은따옴표 제거
4. 31자 초과 시 앞 31자 절단
5. ``History`` (대소문자 무시) → ``History_`` (Excel 예약 이름)
6. 빈 이름 → ``섹션{n}`` (n = 이미 만든 시트 수 + 1)
7. 충돌 시 앞 28자 + ``~{2,3,…}``

두 샘플의 제목 9개는 모두 규칙 (1)만 적용되어 원문 그대로다(최장 17자).
"""

from __future__ import annotations

from .text import cr_to_lf

#: Excel 이 시트 이름에 허용하지 않는 문자.
FORBIDDEN_CHARS = frozenset(":\\/?*[]")
#: Excel 시트 이름 최대 길이.
MAX_SHEET_NAME = 31
#: 충돌 접미사(`~N`)를 붙일 때 남기는 앞부분 길이.
_COLLISION_STEM = 28


def make_sheet_name(title: str, used) -> str:
    """제목 `title` 을 Excel 시트 이름으로 바꾼다. `used` 는 이미 쓴 이름들의 모음.

    Excel 은 시트 이름의 대소문자를 구분하지 않으므로 충돌 판정은 `casefold` 로 한다.
    돌려준 이름을 `used` 에 넣는 것은 호출 측의 몫이다.
    """
    taken = {u.casefold() for u in used}

    name = title if title is not None else ""                          # (1)
    # 안전장치 — `&cr;`·개행·제어문자는 Excel 시트 이름에 들어갈 수 없다(Excel 이 "복구" 하며 이름을 바꾸면
    # section_map 조회가 깨진다). 공백으로 바꾸고 양끝을 다듬는다. 두 샘플의 제목 9개는 영향 없음.
    name = cr_to_lf(name)
    name = "".join(" " if ord(ch) < 32 else ch for ch in name).strip()
    name = "".join("_" if ch in FORBIDDEN_CHARS else ch for ch in name)  # (2)
    name = name.strip("'")                                             # (3)
    if len(name) > MAX_SHEET_NAME:                                     # (4)
        name = name[:MAX_SHEET_NAME].rstrip("'")  # 절단 뒤 끝에 남은 따옴표는 Excel 이 거부한다
    if name.casefold() == "history":                                   # (5)
        name = "History_"
    if not name:                                                       # (6)
        name = f"섹션{len(used) + 1}"
    if name.casefold() in taken:                                       # (7)
        stem = name[:_COLLISION_STEM]
        n = 2
        while f"{stem}~{n}".casefold() in taken:
            n += 1
        name = f"{stem}~{n}"
    return name


__all__ = ["FORBIDDEN_CHARS", "MAX_SHEET_NAME", "make_sheet_name"]
