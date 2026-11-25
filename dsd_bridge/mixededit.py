"""회색 칸·문단(혼합 콘텐츠)의 글자만 바꾸기 (0.3.0).

서식 조각(`SPAN` 등 자식 요소)이 섞인 문단·셀은 1.x 까지 엑셀에서 고치면 거부됐다(코드 5). 고치려면
`--allow-lossy-mixed` 로 서식을 통째로 버리거나, 되돌린 뒤 XML 을 손으로 치환해야 했다.

이 모듈은 **바뀐 글자만** 원본 바이트에 반영한다.

1. 요소의 내부를 문서 순서대로 "글자 단위"(`_Unit`)로 쪼갠다. 단위마다 엑셀에 보이는 글자(`&amp;cr;` 는
   줄바꿈 한 글자), 원본 바이트 구간, 그 글자를 직접 가진 요소, 깊이를 기록한다. 단위를 이으면 엑셀에
   보이는 원래 글(`cr_to_lf(merge_mixed_text(...))`)과 같아야 한다(다르면 포기).
2. 원래 글과 새 글을 `difflib` 로 비교해 바뀐 구간을 얻는다.
3. 구간마다
   - 한 텍스트 조각(같은 요소의 연속 글자) 안이면 그 바이트만 갈아끼운다. 태그·속성·다른 글자는
     바이트 그대로다.
   - 글자를 끼우기만 하는 경우(삭제 없음)는 경계에서 더 바깥(얕은) 쪽 조각에 붙인다 — 굵은 제목
     `SPAN` 바로 뒤에 쓴 글자가 굵게 되지 않도록.
   - 여러 조각에 걸치면, 글자가 전부 지워지는 서식 조각(`SPAN` 만으로 된 하위 트리)은 요소째 지우고,
     나머지 조각은 첫 조각에 새 글자를 넣고 다른 조각에서는 지운다.
4. 원본 XML 의 줄바꿈(들여쓰기 개행, `&cr;` 아님)이 낀 구간에 새 줄바꿈을 넣는 편집, 글자 없는 요소
   (그림 등)를 품은 조각을 지우는 편집처럼 뜻이 모호하면 `None` 을 돌려준다 — 호출 측이 예전처럼
   거부(코드 5)하거나 `--allow-lossy-mixed` 로 평문화한다.

결과는 패치 목록과 통째로 지운 자식 요소 순번(구조 검사용 `planned_deletions`)이다.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field

from .text import CR_TOKEN, cr_to_lf, escape_text, lf_to_cr, merge_mixed_text, unescape
from .xmlsplice import Patch, SourceIndex

#: 글자가 모두 지워지면 요소째 지워도 되는 서식 조각 태그.
INLINE_TAGS = frozenset({"SPAN"})

_TOKEN = re.compile(r"&amp;cr;|&(?:#[0-9]+|#[xX][0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]*);|\r\n|[\s\S]")


@dataclass
class _Unit:
    disp: str          # 엑셀에 보이는 글자 (보통 한 글자)
    start: int         # 원본 바이트 시작 (contents.xml 기준)
    end: int           # 원본 바이트 끝
    owner: int         # 이 글자를 직접 가진 요소 순번
    depth: int         # 편집 대상 요소 기준 깊이 (0 = 그 요소 자신의 글자)
    raw_nl: bool       # 원본 XML 줄바꿈 (`&cr;` 가 아님)
    seg: int           # 조각 번호 — 같은 요소의 끊기지 않은 연속 글자


@dataclass
class MixedEdit:
    """혼합 요소 하나에 대한 글자 편집 결과."""

    patches: list = field(default_factory=list)
    deleted: set = field(default_factory=set)      # 통째로 지운 자식 요소 순번
    regions: int = 0                               # 바뀐 구간 수
    after_raw_newline: bool = False                # 원본 XML 줄바꿈 바로 뒤에 글을 넣었다 (DART 에선 줄이 안 바뀜)


def _units(index: SourceIndex, raw: bytes, ordinal: int) -> list:
    spans = index.spans
    units: list = []
    seg = [0]

    def text_run(a: int, b: int, owner: int, depth: int) -> None:
        if a >= b:
            return
        s = raw[a:b].decode("utf-8")
        pos = a
        for m in _TOKEN.finditer(s):
            tok = m.group(0)
            n = len(tok.encode("utf-8"))
            if tok == "&amp;cr;":
                disp, nl = "\n", False
            elif tok.startswith("&"):
                disp, nl = unescape(tok), False
            elif tok in ("\r\n", "\r", "\n"):
                disp, nl = "\n", True
            else:
                disp, nl = tok, False
            units.append(_Unit(disp, pos, pos + n, owner, depth, nl, seg[0]))
            pos += n
        seg[0] += 1

    def walk(o: int, depth: int) -> None:
        sp = spans[o]
        if sp.self_closing:
            return
        pos = sp.start_gt
        for c in sp.child_ordinals:
            ch = spans[c]
            text_run(pos, ch.start_lt, o, depth)
            walk(c, depth + 1)
            pos = ch.end_gt
        text_run(pos, sp.end_lt, o, depth)

    walk(ordinal, 0)
    return units


def _subtree(index: SourceIndex, o: int) -> list:
    out = [o]
    for c in index.spans[o].child_ordinals:
        out.extend(_subtree(index, c))
    return out


def plan_mixed_edit(index: SourceIndex, raw: bytes, ordinal: int, new_display: str) -> MixedEdit | None:
    """혼합 요소 `ordinal` 의 엑셀 표시 글을 `new_display` 로 바꾸는 패치를 만든다. 못 하면 `None`."""
    units = _units(index, raw, ordinal)
    old_display = "".join(u.disp for u in units)
    if old_display != cr_to_lf(merge_mixed_text(index, raw, ordinal)):
        return None
    new_display = new_display.replace("\r\n", "\n").replace("\r", "\n")
    if old_display == new_display:
        return MixedEdit()

    # 표시 글자 위치 → 단위 번호
    char_unit = []
    unit_first = []
    for k, u in enumerate(units):
        unit_first.append(len(char_unit))
        char_unit.extend([k] * len(u.disp))
    n_old = len(char_unit)

    # 1) 바뀐 구간 (단위 경계로 넓히고, 겹치면 합친다)
    regions = []
    sm = difflib.SequenceMatcher(None, old_display, new_display, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if i1 < n_old:
            left = i1 - unit_first[char_unit[i1]]
        else:
            left = 0
        if i2 > 0 and i2 < n_old and unit_first[char_unit[i2]] != i2:
            k = char_unit[i2 - 1]
            right = unit_first[k] + len(units[k].disp) - i2
        else:
            right = 0
        a, b, c, d = i1 - left, i2 + right, j1 - left, j2 + right
        if regions and a < regions[-1][1]:
            pa, pb, pc, pd = regions[-1]
            regions[-1] = (pa, max(pb, b), pc, max(pd, d))
        else:
            regions.append((a, b, c, d))

    edit = MixedEdit(regions=len(regions))
    spans = index.spans
    for a, b, c, d in regions:
        new_text = new_display[c:d]
        payload = escape_text(lf_to_cr(new_text)).encode("utf-8")
        if a == b:
            # 끼우기 — 자리 고르기:
            #  - 맨 앞/맨 끝이면 요소 자신의 처음/끝(서식 조각 밖)
            #  - 한쪽만 실제 글자(다른 쪽은 원본 XML 줄바꿈)면 실제 글자 쪽 조각에 붙인다
            #    (`<TD>\r\n<SPAN>…현재</SPAN>\r\n</TD>` 끝에 덧붙인 글이 SPAN 서식을 잇도록)
            #  - 양쪽 다 실제 글자면 얕은(바깥) 쪽, 같으면 왼쪽 — 굵은 제목 뒤 본문이 굵어지지 않도록
            ku = char_unit[a] if a < n_old else len(units)
            left_u = units[ku - 1] if ku > 0 else None
            right_u = units[ku] if ku < len(units) else None
            if left_u is None and right_u is None:
                return None
            if left_u is None:
                at = spans[ordinal].start_gt
            elif right_u is None:
                at = spans[ordinal].end_lt
            elif not left_u.raw_nl and right_u.raw_nl:
                at = left_u.end
            elif left_u.raw_nl and not right_u.raw_nl:
                at = right_u.start
            elif right_u.depth < left_u.depth:
                at = right_u.start
            else:
                at = left_u.end
            if left_u is not None and left_u.raw_nl and at == left_u.end and new_text.strip():
                edit.after_raw_newline = True
            edit.patches.append(Patch(at, at, payload, label=f"혼합#{ordinal} 끼우기"))
            continue

        ua = char_unit[a]
        ub = char_unit[b - 1] + 1
        sel = units[ua:ub]
        if "\n" in new_text and any(u.raw_nl for u in sel):
            return None                     # 원본 들여쓰기 개행 자리에 새 줄바꿈 — 뜻이 모호하다
        if len({u.seg for u in sel}) == 1:
            owner = sel[0].owner
            if not payload and owner != ordinal:
                # 서식 조각의 글자를 다 지웠으면 빈 조각을 남기지 않고 요소째 지운다(여러 조각 경로와 같게).
                own = [u for x in _subtree(index, owner) for u in units if u.owner == x]
                if (len(own) == len(sel) and all(spans[x].tag in INLINE_TAGS for x in _subtree(index, owner))):
                    sp_o = spans[owner]
                    edit.patches.append(Patch(sp_o.start_lt, sp_o.end_gt, b"",
                                              label=f"혼합#{ordinal} {sp_o.tag}#{owner} 삭제"))
                    edit.deleted.add(owner)
                    continue
            edit.patches.append(Patch(sel[0].start, sel[-1].end, payload, label=f"혼합#{ordinal} 글자"))
            continue

        # 여러 조각에 걸친 편집 — 글자가 다 지워지는 서식 조각은 요소째 지운다.
        sel_ids = {id(u) for u in sel}
        units_of: dict = {}
        for u in units:
            units_of.setdefault(u.owner, []).append(u)

        def subtree_units(o: int) -> list:
            out = []
            for x in _subtree(index, o):
                out.extend(units_of.get(x, ()))
            return out

        covered = []
        stack = list(spans[ordinal].child_ordinals)
        while stack:
            o = stack.pop(0)
            su = subtree_units(o)
            if su and all(id(u) in sel_ids for u in su):
                if any(spans[x].tag not in INLINE_TAGS for x in _subtree(index, o)):
                    return None             # 그림 등 글자 없는 요소를 함께 지우게 된다
                covered.append(o)
            else:
                stack[0:0] = list(spans[o].child_ordinals)
        covered_units = set()
        for o in covered:
            covered_units.update(id(u) for u in subtree_units(o))
        rest = [u for u in sel if id(u) not in covered_units]

        # 남은 글자를 조각별 연속 구간으로 묶는다.
        runs = []
        for u in rest:
            if runs and runs[-1][-1].seg == u.seg and runs[-1][-1].end == u.start:
                runs[-1].append(u)
            else:
                runs.append([u])
        placed = False
        for run in runs:
            body = payload if not placed else b""
            placed = True
            edit.patches.append(Patch(run[0].start, run[-1].end, body, label=f"혼합#{ordinal} 글자"))
        for o in covered:
            sp = spans[o]
            edit.patches.append(Patch(sp.start_lt, sp.end_gt, b"", label=f"혼합#{ordinal} {sp.tag}#{o} 삭제"))
            edit.deleted.add(o)
        if not placed and payload:
            first = min(covered, key=lambda o: spans[o].start_lt)
            at = spans[first].start_lt
            edit.patches.append(Patch(at, at, payload, label=f"혼합#{ordinal} 끼우기"))
    return edit


__all__ = ["MixedEdit", "plan_mixed_edit", "INLINE_TAGS", "CR_TOKEN"]
