"""오프셋 인덱스와 바이트 스플라이스 패치 엔진 (계획 §4.2, §4.8, §4.10).

**이 모듈은 XML 을 직렬화하지 않는다.** `contents.xml` 의 원본 바이트를 그대로 두고
바이트 구간만 갈아끼운다(아키텍처 b1). 그래야 CRLF 3,993/3,210쌍, `<X></X>`
601/451개, 3중 개행 선언부, 속성 순서·인용부호, 빈 줄 배치가 **자동으로** 보존된다.

구성:

- `ElementSpan` — 요소 하나의 바이트 오프셋·트리 정보
- `SourceIndex` — `expat` 로 문서 전체를 훑어 만든 `ElementSpan` 목록 + eid/다이제스트
- `Patch` / `PatchSet` — 겹침 검사와 적용 (§4.8 경계 규칙)
- `structural_signature` / `check_rewrite_identity` / `escape_roundtrip_check` — §4.10 사후 검사
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from xml.parsers import expat

from .errors import PatchIntegrityError
from .text import cr_to_lf, escape_text, lf_to_cr, unescape

# ─────────────────────────────────────────────────────────────────────────────
# ElementSpan
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class ElementSpan:
    """요소 하나의 바이트 오프셋과 트리 위치 (계획 §4.2).

    - 내부 텍스트 구간 = `[start_gt, end_lt)`
    - 요소 전체 구간   = `[start_lt, end_gt)`
    """

    ordinal: int
    tag: str
    attrs: dict
    start_lt: int          # 시작 태그 '<' 오프셋
    start_gt: int          # 시작 태그 '>' 다음 오프셋
    end_lt: int = -1       # 종료 태그 '<' 오프셋 (자기닫힘이면 start_gt)
    end_gt: int = -1       # 종료 태그 '>' 다음 오프셋 (자기닫힘이면 start_gt)
    self_closing: bool = False
    parent_ordinal: int = -1
    depth: int = 0
    child_ordinals: list = field(default_factory=list)
    line: int = 0          # 시작 태그의 1-기반 행 번호
    end_line: int = 0      # 종료 태그의 1-기반 행 번호
    text: str = ""         # 자기 자신의 문자 데이터(자손 제외, expat 정규화 상태)

    @property
    def inner_span(self) -> tuple:
        """내부 텍스트 구간 `[start_gt, end_lt)`."""
        return (self.start_gt, self.end_lt)

    @property
    def outer_span(self) -> tuple:
        """요소 전체 구간 `[start_lt, end_gt)`."""
        return (self.start_lt, self.end_gt)

    @property
    def has_child_elements(self) -> bool:
        """자식 **요소**를 가지는가 (= 혼합 콘텐츠 후보). 텍스트만 있으면 False."""
        return bool(self.child_ordinals)


def _scan_start_gt(raw: bytes, lt: int) -> int:
    """시작 태그의 '>' 다음 오프셋. 속성값 인용부호 안의 '>' 는 건너뛴다."""
    i = lt + 1
    quote = 0
    n = len(raw)
    while i < n:
        c = raw[i]
        if quote:
            if c == quote:
                quote = 0
        elif c in (0x22, 0x27):   # " '
            quote = c
        elif c == 0x3E:           # >
            return i + 1
        i += 1
    raise ValueError(f"시작 태그가 닫히지 않았습니다 (오프셋 {lt})")


def build_spans(raw: bytes) -> list:
    """`contents.xml` 바이트를 문서 순서로 훑어 `ElementSpan` 목록을 만든다.

    `parser.buffer_text = False` 는 **필수**다. 기본값(True)에서는 문자 데이터가
    버퍼링되어 `CurrentByteIndex` 가 어긋난다.
    """
    spans: list = []
    stack: list = []

    p = expat.ParserCreate()
    p.buffer_text = False

    def start(tag, attrs):
        lt = p.CurrentByteIndex
        sp = ElementSpan(
            ordinal=len(spans),
            tag=tag,
            attrs=dict(attrs),
            start_lt=lt,
            start_gt=_scan_start_gt(raw, lt),
            depth=len(stack),
            line=p.CurrentLineNumber,
            parent_ordinal=stack[-1].ordinal if stack else -1,
        )
        if stack:
            stack[-1].child_ordinals.append(sp.ordinal)
        spans.append(sp)
        stack.append(sp)

    def end(tag):
        sp = stack.pop()
        sp.end_line = p.CurrentLineNumber
        # 실측(CPython 3.13.15): expat 은 자기닫힘 <X/> 의 EndElementHandler 에서
        # CurrentByteIndex 로 **시작 태그의 다음 오프셋**(= start_gt)을 준다. 이 값은
        # <X></X> 의 종료 태그 '<' 오프셋과 정확히 같아서 바이트 인덱스만으로는 둘을
        # 구분할 수 없다. 계획 §4.2 의 `end_lt == start_lt` 규칙은 성립하지 않는다.
        # 유일하게 확실한 판별은 시작 태그가 '/>' 로 끝나는지 원본 바이트로 보는 것이다.
        sp.self_closing = raw[sp.start_gt - 2 : sp.start_gt] == b"/>"
        if sp.self_closing:
            sp.end_lt = sp.start_gt
            sp.end_gt = sp.start_gt
        else:
            sp.end_lt = p.CurrentByteIndex
            sp.end_gt = raw.index(b">", sp.end_lt) + 1

    def chars(data):
        if stack:
            stack[-1].text += data

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    p.Parse(raw, True)
    return spans


# ─────────────────────────────────────────────────────────────────────────────
# SourceIndex
# ─────────────────────────────────────────────────────────────────────────────

#: eid 접미사의 **마지막 글자**를 `[a-f]` 로 옮기는 고정 대응표.
#:
#: Excel 채우기 핸들은 끝이 숫자인 문자열을 증분하므로(계획 §4.2 / AC-29), 마지막
#: 글자가 절대 숫자가 되지 않도록 못 박는다. 대응은 `'abcdef'[int(c) % 6]` 이다::
#:
#:     0→a  1→b  2→c  3→d  4→e  5→f  6→a  7→b  8→c  9→d
#:
#: (`a`–`f` 는 그대로 둔다.) 계획 §4.2 는 "`[a-f]` 가 나올 때까지 다이제스트를 한 칸씩
#: 민다"고 썼으나, 그 규칙은 다이제스트에 따라 정의되지 않을 수 있어 고정 대응으로 바꿨다.
_LAST_HEX_TO_AF = {
    "0": "a", "1": "b", "2": "c", "3": "d", "4": "e",
    "5": "f", "6": "a", "7": "b", "8": "c", "9": "d",
}


class SourceIndex:
    """`contents.xml` 원본 바이트에 대한 요소 인덱스 (계획 §4.2).

    오프셋은 **이 `raw` 에 절대적으로 종속**된다. 다른 바이트열에 그대로 쓰면 안 된다.
    """

    __slots__ = ("raw", "spans", "_tag_digest")

    def __init__(self, raw: bytes, spans: list) -> None:
        self.raw = raw
        self.spans = spans
        self._tag_digest = None

    # 구축 ────────────────────────────────────────────────────────────────
    @classmethod
    def build(cls, raw: bytes) -> "SourceIndex":
        return cls(raw, build_spans(raw))

    # 목록 프로토콜 ────────────────────────────────────────────────────────
    def __len__(self) -> int:
        return len(self.spans)

    def __iter__(self):
        return iter(self.spans)

    def __getitem__(self, ordinal: int) -> ElementSpan:
        return self.spans[ordinal]

    # 구간 ────────────────────────────────────────────────────────────────
    def inner_span(self, ordinal: int) -> tuple:
        """내부 텍스트 구간 `[start_gt, end_lt)`."""
        return self.spans[ordinal].inner_span

    def outer_span(self, ordinal: int) -> tuple:
        """요소 전체 구간 `[start_lt, end_gt)`."""
        return self.spans[ordinal].outer_span

    def inner_bytes(self, ordinal: int) -> bytes:
        start, end = self.inner_span(ordinal)
        return self.raw[start:end]

    def outer_bytes(self, ordinal: int) -> bytes:
        start, end = self.outer_span(ordinal)
        return self.raw[start:end]

    def start_tag_bytes(self, ordinal: int) -> bytes:
        sp = self.spans[ordinal]
        return self.raw[sp.start_lt : sp.start_gt]

    def leading_ws_start(self, ordinal: int) -> int:
        """`start_lt` 바로 앞의 선행 공백 구간 시작 오프셋 (계획 §4.2).

        규칙: 같은 줄의 스페이스·탭을 거슬러 올라간 뒤, **줄바꿈을 최대 2개까지**
        (= 빈 줄 1개까지) 먹는다. 줄바꿈 사이의 스페이스·탭도 함께 먹는다.
        비공백을 만나거나 줄바꿈 2개를 채우면 멈춘다.

        왜 2개인가 — 실측상 363개 `TR` 전부가 앞에 빈 줄 하나를 두고 있다
        (`</TR>\\r\\n` `\\r\\n` `<TR>`; 예 샘플A:499 → 샘플A:500). 줄바꿈을 **하나만** 먹으면
        `[leading_ws_start, end_gt)` 삭제 후 `<TBODY>\\r\\n\\r\\n\\r\\n<TR>` 처럼 빈 줄이
        하나 늘고, 복제 삽입도 빈 줄 없이 붙어 원본 배치가 깨진다. 2개를 먹으면 삭제 후
        `<TBODY>\\r\\n\\r\\n<TR>` 로 원본 간격이 그대로 남는다 (§4.8 / US-006 시나리오 2).
        """
        raw = self.raw
        i = self.spans[ordinal].start_lt
        breaks = 0
        while True:
            j = i
            while j > 0 and raw[j - 1] in (0x20, 0x09):      # space, tab
                j -= 1
            if j > 0 and raw[j - 1] == 0x0A:                 # \n
                j -= 1
                if j > 0 and raw[j - 1] == 0x0D:             # \r
                    j -= 1
                i = j
                breaks += 1
                if breaks == 2:                              # 빈 줄 1개까지
                    break
            else:
                if breaks == 0:      # 같은 줄의 공백은 언제나 포함한다
                    i = j
                break
        return i

    # 식별자 ──────────────────────────────────────────────────────────────
    def eid_of(self, ordinal: int) -> str:
        """요소 ID `E{순번:05d}.{체크섬 4자리}` (계획 §4.2, AC-29).

        체크섬 = `sha256(태그명 UTF-8 + 시작 태그 원본 바이트)` 의 앞 4자리 hex 이며,
        **마지막 글자만** `_LAST_HEX_TO_AF` 로 `[a-f]` 에 옮긴다. 결과는 항상
        정규식 ``E\\d{5,7}\\.[0-9a-f]{4}`` 를 만족하고 마지막 글자가 `[a-f]` 다 —
        `:05d` 는 5자리 미만을 0 으로 채울 뿐이라 요소가 10만 개를 넘으면 자릿수가 자연히
        늘어난다 (`dsd_bridge.gates.EID_RE` 가 그 범위를 받는다).
        """
        sp = self.spans[ordinal]
        digest = hashlib.sha256(
            sp.tag.encode("utf-8") + self.raw[sp.start_lt : sp.start_gt]
        ).hexdigest()
        suffix = digest[:4]
        return f"E{ordinal:05d}.{suffix[:3]}{_LAST_HEX_TO_AF.get(suffix[3], suffix[3])}"

    def tag_digest(self) -> str:
        """전 요소의 태그명을 `\\n` 으로 이은 문자열의 SHA-256 (AC-18 3조건 중 하나).

        구분자를 두는 이유: 구분자 없이 이으면 `TR`+`TE` 와 `TRT`+`E` 가 같은 문자열이
        되어 서로 다른 문서가 같은 다이제스트를 가질 수 있다.
        """
        if self._tag_digest is None:
            joined = "\n".join(sp.tag for sp in self.spans)
            self._tag_digest = hashlib.sha256(joined.encode("utf-8")).hexdigest()
        return self._tag_digest

    # 편의 ────────────────────────────────────────────────────────────────
    def parent_path(self, ordinal: int) -> tuple:
        """루트부터 부모까지의 태그 시퀀스."""
        path = []
        po = self.spans[ordinal].parent_ordinal
        while po >= 0:
            path.append(self.spans[po].tag)
            po = self.spans[po].parent_ordinal
        path.reverse()
        return tuple(path)

    def attr_bytes(self, ordinal: int) -> bytes:
        """시작 태그에서 태그명과 `>` 사이의 **속성 원본 바이트** (양끝 공백·`/` 제거).

        `<X A="1"/>` 와 `<X A="1">` 가 같은 값을 주므로, AC-22 의 자기닫힘 전환이
        구조 서명을 흔들지 않는다.
        """
        sp = self.spans[ordinal]
        head = self.raw[sp.start_lt + 1 + len(sp.tag.encode("utf-8")) : sp.start_gt - 1]
        if head.endswith(b"/"):
            head = head[:-1]
        return head.strip()


# ─────────────────────────────────────────────────────────────────────────────
# 패치 엔진 (계획 §4.8)
# ─────────────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Patch:
    """원본 바이트의 `[start, end)` 를 `replacement` 로 갈아끼우는 하나의 편집.

    - `start == end` 이면 **길이 0 삽입**(§4.8 경계 규칙 1)
    - `replacement == b""` 이면 삭제
    """

    start: int
    end: int
    replacement: bytes
    label: str = ""

    def __post_init__(self) -> None:
        if self.start < 0 or self.end < self.start:
            raise PatchIntegrityError(
                f"패치 구간이 올바르지 않습니다: [{self.start}, {self.end}) {self.label}"
            )

    @property
    def is_insert(self) -> bool:
        return self.start == self.end

    @property
    def is_delete(self) -> bool:
        return self.start < self.end and not self.replacement

    @property
    def delta(self) -> int:
        """이 패치가 문서 길이에 주는 변화량."""
        return len(self.replacement) - (self.end - self.start)


class PatchSet:
    """패치 모음. 적용 전 반드시 겹침을 검사한다 (§4.8, 실패 시 종료 코드 6)."""

    __slots__ = ("patches",)

    def __init__(self, patches=None) -> None:
        self.patches: list = list(patches) if patches else []

    def __len__(self) -> int:
        return len(self.patches)

    def __iter__(self):
        return iter(self.patches)

    def add(self, patch: Patch) -> Patch:
        """패치 하나를 넣는다. 겹침 검사는 `check_overlaps()`/`apply()` 에서 한다."""
        if not isinstance(patch, Patch):
            raise TypeError("PatchSet.add 에는 Patch 를 넘겨야 합니다.")
        self.patches.append(patch)
        return patch

    def _ordered(self) -> list:
        """적용 순서 — `(start, end)` 오름차순.

        `(start, end)` 로 정렬하면 같은 오프셋에서 **길이 0 삽입이 삭제 구간보다 먼저**
        나온다. 그래서 삽입 내용이 삭제 구간에 삼켜지지 않는다(§4.8 경계 규칙 3의 의도).
        같은 구간의 패치들 사이에서는 넣은 순서를 유지한다(파이썬 정렬은 안정적이다).
        """
        return sorted(self.patches, key=lambda p: (p.start, p.end))

    def check_overlaps(self) -> None:
        """실제로 겹치는 두 구간이 있으면 `PatchIntegrityError`(코드 6).

        §4.8 경계 규칙 2: **끝점에서만 맞닿는 두 구간은 겹침이 아니다** —
        `a.end == b.start` 는 통과시킨다. 길이 0 삽입도 마찬가지다.
        """
        ordered = self._ordered()
        for prev, cur in zip(ordered, ordered[1:]):
            if cur.start < prev.end:
                raise PatchIntegrityError(
                    "패치 구간이 겹칩니다: "
                    f"[{prev.start}, {prev.end}){' ' + prev.label if prev.label else ''} 와 "
                    f"[{cur.start}, {cur.end}){' ' + cur.label if cur.label else ''}. "
                    "같은 요소에 두 번 이상 편집이 걸렸을 수 있습니다."
                )

    def apply(self, raw: bytes) -> bytes:
        """`raw` 에 모든 패치를 적용한 새 바이트열을 돌려준다.

        빈 `PatchSet` 은 원본을 그대로 돌려준다(항등 패치).
        """
        for p in self.patches:
            if p.end > len(raw):
                raise PatchIntegrityError(
                    f"패치 구간 [{p.start}, {p.end}) 가 원본 길이 {len(raw)} 를 벗어납니다."
                )
        self.check_overlaps()
        out = []
        pos = 0
        for p in self._ordered():
            out.append(raw[pos : p.start])
            out.append(p.replacement)
            pos = p.end
        out.append(raw[pos:])
        return b"".join(out)

    def stats(self) -> dict:
        """패치 요약 — 로그·`-v` 출력용."""
        inserts = sum(1 for p in self.patches if p.is_insert)
        deletes = sum(1 for p in self.patches if p.is_delete)
        return {
            "count": len(self.patches),
            "inserts": inserts,
            "deletes": deletes,
            "replacements": len(self.patches) - inserts - deletes,
            "bytes_removed": sum(p.end - p.start for p in self.patches),
            "bytes_added": sum(len(p.replacement) for p in self.patches),
            "delta": sum(p.delta for p in self.patches),
        }


def set_text_patch(index: SourceIndex, ordinal: int, new_inner: bytes) -> Patch:
    """요소의 내부 텍스트 구간을 `new_inner` 로 바꾸는 패치 하나를 만든다.

    **AC-22 (자기닫힘 예외)**: 원본이 `<X/>` 면 텍스트를 담을 자리가 없으므로 요소 전체
    구간을 `<X>new</X>` 로 갈아끼운다. 속성은 원본 바이트 그대로 옮겨진다.
    """
    sp = index[ordinal]
    if not sp.self_closing:
        return Patch(sp.start_gt, sp.end_lt, new_inner, label=f"{sp.tag}#{ordinal}")
    tag = sp.tag.encode("utf-8")
    # b'<X ... />' 에서 '/>' 를 떼고 그 앞 공백을 없앤 뒤 '>' 를 붙인다.
    open_tag = index.raw[sp.start_lt : sp.start_gt][:-2].rstrip() + b">"
    return Patch(
        sp.start_lt,
        sp.end_gt,
        open_tag + new_inner + b"</" + tag + b">",
        label=f"{sp.tag}#{ordinal} 자기닫힘→쌍태그",
    )


# ─────────────────────────────────────────────────────────────────────────────
# 사후 검사 (계획 §4.10)
# ─────────────────────────────────────────────────────────────────────────────


def structural_signature(raw: bytes) -> list:
    """구조 서명 — 요소마다 `(태그, 시작 태그 속성 원본 바이트, 부모 경로)` (§4.10 검사 3).

    문서 순서 그대로의 목록이다. 원본과 결과의 이 수열이 정확히 같으면 "구조 동등"이다.
    속성은 **원본 바이트**로 비교하므로 속성 순서·인용부호·공백까지 지킨다.
    """
    index = SourceIndex.build(raw)
    return [
        (sp.tag, index.attr_bytes(sp.ordinal), index.parent_path(sp.ordinal))
        for sp in index.spans
    ]


def check_rewrite_identity(raw: bytes, index: SourceIndex | None = None) -> None:
    """불변식 `REWRITE_IDENTITY` — 전량 재기입이 원본과 바이트 동일해야 한다 (§4.10).

    모든 요소의 내부 구간을 **`raw[start_gt:end_lt]` 바이트 슬라이스 그대로** 되돌려
    넣고 결과가 원본과 같은지 본다. 파서가 준 텍스트를 재이스케이프하는 것이 **아니다**
    — expat 은 문자 데이터의 CRLF 를 LF 로 정규화하므로 그렇게 하면 실제 CR/LF 를 가진
    요소(샘플A 16 / 샘플B 112)에서 반드시 어긋난다.

    중첩 요소의 내부 구간은 서로 포함 관계이므로 한 `PatchSet` 에 다 넣을 수 없다.
    **깊이별로** 묶으면 형제끼리는 구간이 서로소이므로 겹침 없이 전량을 덮을 수 있다.

    이 불변식이 증명하는 것은 **오프셋 산술뿐**이다 — `start_gt`/`end_lt` 가 정확한
    경계를 가리키고 패치 적용이 구간을 어긋나게 만들지 않음. 텍스트 왕복 충실도는
    `escape_roundtrip_check` 와 AC-4′ 가 담당한다.

    구간을 **자기 자신의 슬라이스로** 되돌려 넣는 재기입은 구간이 서로소이고 범위 안에만
    있으면 언제나 원본과 같다. 그래서 경계가 실제로 태그를 가리키는지를 함께 못 박는다
    — 시작 오프셋에 `<태그`, 종료 구간에 `</태그>` 가 놓여 있어야 한다. 이 두 가지가
    같이 성립해야 "오프셋 산술이 이 문서에서 성립한다"고 말할 수 있다.

    실패하면 `PatchIntegrityError`(코드 6). 성공하면 `None`.
    """
    if index is None:
        index = SourceIndex.build(raw)

    for sp in index.spans:
        tag = sp.tag.encode("utf-8")
        if raw[sp.start_lt : sp.start_lt + 1 + len(tag)] != b"<" + tag:
            raise PatchIntegrityError(
                f"시작 태그 오프셋이 태그를 가리키지 않습니다: {sp.tag}#{sp.ordinal} "
                f"(행 {sp.line}, 오프셋 {sp.start_lt}, 실제 {raw[sp.start_lt:sp.start_lt + 16]!r})"
            )
        if sp.self_closing:
            if not (sp.end_lt == sp.end_gt == sp.start_gt
                    and raw[sp.start_gt - 2 : sp.start_gt] == b"/>"):
                raise PatchIntegrityError(
                    f"자기닫힘 요소의 경계가 어긋났습니다: {sp.tag}#{sp.ordinal} (행 {sp.line})"
                )
            continue
        closing = raw[sp.end_lt : sp.end_gt]
        if not (
            closing.startswith(b"</" + tag)
            and closing.endswith(b">")
            and not closing[2 + len(tag) : -1].strip()
        ):
            raise PatchIntegrityError(
                f"종료 태그 오프셋이 `</{sp.tag}>` 를 가리키지 않습니다: "
                f"{sp.tag}#{sp.ordinal} (행 {sp.line}, 오프셋 {sp.end_lt}, 실제 {closing[:16]!r})"
            )

    by_depth: dict = {}
    for sp in index.spans:
        by_depth.setdefault(sp.depth, []).append(sp)

    for depth in sorted(by_depth):
        ps = PatchSet()
        for sp in by_depth[depth]:
            ps.add(
                Patch(
                    sp.start_gt,
                    sp.end_lt,
                    raw[sp.start_gt : sp.end_lt],
                    label=f"{sp.tag}#{sp.ordinal}",
                )
            )
        out = ps.apply(raw)
        if out != raw:
            where = next(
                (i for i in range(min(len(out), len(raw))) if out[i] != raw[i]),
                min(len(out), len(raw)),
            )
            raise PatchIntegrityError(
                f"전량 재기입이 원본과 다릅니다 (깊이 {depth}, 첫 불일치 오프셋 {where}). "
                "오프셋 계산이 이 문서에서 성립하지 않으므로 편집을 시도하지 않습니다."
            )


def escape_roundtrip_check(index: SourceIndex, raw: bytes) -> dict:
    """이스케이프 왕복 검증 — 순수 텍스트 요소에 한해 (§5 단계 2).

    각 요소의 내부 바이트 슬라이스에 대해
    ``escape_text(lf_to_cr(cr_to_lf(unescape(slice)))) == slice`` 를 단언한다.

    **혼합 콘텐츠 요소(자식 요소를 가진 것)는 제외한다** — 내부 구간에 자식 태그 원본이
    섞여 있어 텍스트 규칙이 적용되지 않으며, 애초에 AC-16 으로 편집이 막혀 있다.
    자기닫힘 요소도 내부 구간이 비어 있으므로 제외한다.

    돌려주는 값: `{"checked": 검사한 요소 수, "skipped": 건너뛴 요소 수}`.
    어긋나면 `PatchIntegrityError`(코드 6).
    """
    checked = 0
    skipped = 0
    for sp in index.spans:
        if sp.has_child_elements or sp.self_closing:
            skipped += 1
            continue
        slice_ = raw[sp.start_gt : sp.end_lt]
        try:
            s = slice_.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise PatchIntegrityError(
                f"요소 {sp.tag}#{sp.ordinal} (행 {sp.line}) 의 내부 구간이 UTF-8 이 아닙니다: {exc}"
            ) from exc
        again = escape_text(lf_to_cr(cr_to_lf(unescape(s)))).encode("utf-8")
        if again != slice_:
            raise PatchIntegrityError(
                f"이스케이프 왕복이 어긋났습니다: {sp.tag}#{sp.ordinal} (행 {sp.line})\n"
                f"  원본: {slice_[:120]!r}\n"
                f"  왕복: {again[:120]!r}"
            )
        checked += 1
    return {"checked": checked, "skipped": skipped}


__all__ = [
    "ElementSpan",
    "SourceIndex",
    "Patch",
    "PatchSet",
    "build_spans",
    "_scan_start_gt",
    "set_text_patch",
    "structural_signature",
    "check_rewrite_identity",
    "escape_roundtrip_check",
]
