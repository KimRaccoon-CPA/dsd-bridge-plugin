"""출력 전 사후 검사 (계획 §4.10) 와 XSD 게이트 (AC-25).

1. `check_wellformed` — `ET.fromstring` 이 성공해야 한다.
2. `count_elements` — 요소 총수 = 원본 + Σ삽입 서브트리 − Σ삭제 서브트리 − Σ평탄화 서브트리.
3. `structural_equivalence` — 원본 서명 수열에서 `RowPlan` 대로 서브트리를 빼고 넣은 **기대 수열**이
   결과의 서명 수열과 정확히 일치해야 한다. 서명 = `(태그, 시작 태그 속성 바이트, 부모 경로)`.
4. `xsd_validate` — `Dart{3,4}.xsd` 로 lxml XSD 검증. 로드 실패는 **조용히 넘기지 않고**
   `status="skipped"` + 사유를 돌려준다(P5).

XSD 위치 탐색 순서: 인자 `dart_root` → 환경변수 `DSD_BRIDGE_DART_XSD_DIR` → `C:\\DART\\Resource\\Dtd`
→ 프로젝트 안 `dart-xsd/`.
"""

from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

from .errors import PatchIntegrityError
from .xmlsplice import SourceIndex, structural_signature

WINDOWS_DART_DTD = Path(r"C:\DART\Resource\Dtd")
PROJECT_XSD_DIR = Path(__file__).resolve().parents[1] / "dart-xsd"


# ─────────────────────────────────────────────────────────────────────────────
# 서브트리 도우미
# ─────────────────────────────────────────────────────────────────────────────


def subtree_range(index: SourceIndex, ordinal: int) -> tuple:
    """`ordinal` 의 서브트리가 차지하는 순번 구간 `[ordinal, end)` — expat 은 전위 순서로 번호를
    매기므로 서브트리는 언제나 연속 구간이다(두 샘플 실측: 위반 0)."""
    spans = index.spans
    depth = spans[ordinal].depth
    end = ordinal + 1
    n = len(spans)
    while end < n and spans[end].depth > depth:
        end += 1
    return ordinal, end


def subtree_size(index: SourceIndex, ordinal: int) -> int:
    a, b = subtree_range(index, ordinal)
    return b - a


# ─────────────────────────────────────────────────────────────────────────────
# 검사 1 · 2
# ─────────────────────────────────────────────────────────────────────────────


def check_wellformed(raw: bytes) -> None:
    """검사 1. 실패하면 `PatchIntegrityError`(코드 6)."""
    try:
        ET.fromstring(raw)
    except ET.ParseError as exc:
        raise PatchIntegrityError(
            f"오류: 되돌린 XML 이 올바른 형식이 아닙니다 ({exc}). 출력 파일을 만들지 않았습니다."
        ) from exc


def count_elements(raw: bytes) -> int:
    return len(SourceIndex.build(raw))


# ─────────────────────────────────────────────────────────────────────────────
# 검사 3 — 구조 동등성 (서브트리 전개)
# ─────────────────────────────────────────────────────────────────────────────


def expected_signatures(orig_index: SourceIndex, plan) -> list:
    """`RowPlan` 으로부터 기대 서명 수열을 만든다.

    - 유지 → 원본 서명 그대로
    - 삭제 행 / 평탄화된 자식 → 그 서브트리 전체 제외
    - 삽입(복제) 행 → 앵커(복제 원본) 서브트리의 서명을 삽입 위치 뒤에 한 번 더 (단계 6)
    """
    raw = orig_index.raw
    sigs = [
        (sp.tag, orig_index.attr_bytes(sp.ordinal), orig_index.parent_path(sp.ordinal))
        for sp in orig_index.spans
    ]
    removed = set()
    for o in list(plan.deleted_ordinals()) + list(plan.planned_deletions):
        a, b = subtree_range(orig_index, o)
        removed.update(range(a, b))
    # 삽입: 위치 앵커 순번 → [복제 원본 서브트리 서명들] (앵커 서브트리 끝 뒤에 붙는다)
    after: dict = {}
    before: dict = {}
    for ins in plan.inserted:
        a, b = subtree_range(orig_index, ins.clone_ordinal)
        gone = set()
        for f in getattr(ins, "flattened", ()):
            fa, fb = subtree_range(orig_index, f)
            gone.update(range(fa, fb))
        block = [sigs[i] for i in range(a, b) if i not in gone]
        target = before if ins.before else after
        target.setdefault(ins.anchor_ordinal, []).append(block)
    out = []
    n = len(sigs)
    i = 0
    while i < n:
        if i in before:
            for block in before[i]:
                out.extend(block)
        if i in removed:
            i += 1
            continue
        out.append(sigs[i])
        if i in after:
            _a, b = subtree_range(orig_index, i)
            # 서브트리 끝 뒤에 붙인다: 서브트리 안의 요소들을 먼저 내보낸 다음
            for j in range(i + 1, b):
                if j not in removed:
                    out.append(sigs[j])
            for block in after[i]:
                out.extend(block)
            i = b
            continue
        i += 1
    del raw
    return out


def structural_equivalence(orig_index: SourceIndex, new_raw: bytes, plan) -> int:
    """검사 3. 통과하면 결과의 요소 수를 돌려주고, 실패하면 `PatchIntegrityError`(코드 6)."""
    expected = expected_signatures(orig_index, plan)
    actual = structural_signature(new_raw)
    if expected == actual:
        return len(actual)
    new_index = SourceIndex.build(new_raw)
    k = next((i for i in range(min(len(expected), len(actual))) if expected[i] != actual[i]),
             min(len(expected), len(actual)))
    exp = expected[k] if k < len(expected) else None
    act = actual[k] if k < len(actual) else None
    line = new_index.spans[k].line if k < len(new_index) else "?"
    raise PatchIntegrityError(
        "오류: 되돌린 XML 의 구조가 기대와 다릅니다 (구조 동등성 실패). 출력 파일을 만들지 않았습니다.\n"
        f"  첫 불일치: 순번 {k} (결과 {line}행)\n"
        f"  기대: {exp[0] if exp else '(없음)'} {exp[1][:60]!r}\n"
        f"  실제: {act[0] if act else '(없음)'} {act[1][:60]!r}\n"
        f"  기대 요소 수 {len(expected):,} / 실제 {len(actual):,}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# 검사 4 — XSD (AC-25)
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class ValidationResult:
    status: str                     # "pass" | "fail" | "skipped"
    reason: str = ""                # skipped 사유 / 첫 오류
    errors: list = field(default_factory=list)
    xsd_path: str = ""

    @property
    def status_ko(self) -> str:
        return {"pass": "통과", "fail": "실패", "skipped": "건너뜀"}[self.status]


def find_xsd_dir(dart_root=None) -> Path | None:
    """XSD 폴더 탐색: 인자 → DSD_BRIDGE_DART_XSD_DIR → C:\\DART\\Resource\\Dtd → 프로젝트 dart-xsd/."""
    candidates = []
    if dart_root:
        candidates.append(Path(dart_root))
    # 옛 이름(dsdx 시절) 환경변수도 받는다
    env = os.environ.get("DSD_BRIDGE_DART_XSD_DIR") or os.environ.get("DSDX_DART_XSD_DIR")
    if env:
        candidates.append(Path(env))
    candidates.append(WINDOWS_DART_DTD)
    candidates.append(PROJECT_XSD_DIR)
    for c in candidates:
        if c.is_dir():
            return c
    return None


def xsd_path_for(schema: str | None, dart_root=None) -> tuple:
    """(xsd 경로 | None, 사유). `schema` 는 `dart3` / `dart4`."""
    if not schema:
        return None, "meta.xml 의 GENERATOR@schema 가 없어 어느 XSD 를 쓸지 알 수 없습니다"
    name = schema.lower()
    if not name.startswith("dart") or not name[4:].isdigit():
        return None, f"알 수 없는 스키마 이름: {schema}"
    filename = f"Dart{name[4:]}.xsd"
    d = find_xsd_dir(dart_root)
    if d is None:
        return None, ("XSD 폴더를 찾지 못했습니다 — DSD_BRIDGE_DART_XSD_DIR 환경변수를 설정하거나 "
                      f"{WINDOWS_DART_DTD} / {PROJECT_XSD_DIR} 에 Dart3.xsd·Dart4.xsd 를 두세요")
    p = d / filename
    if not p.is_file():
        return None, f"{p} 가 없습니다"
    return p, ""


def xsd_validate(raw: bytes, schema: str | None, dart_root=None) -> ValidationResult:
    """`raw`(contents.xml) 를 `schema` 의 XSD 로 검증한다. 로드 실패는 `skipped` + 사유."""
    path, reason = xsd_path_for(schema, dart_root)
    if path is None:
        return ValidationResult("skipped", reason)
    try:
        from lxml import etree
    except ImportError as exc:  # pragma: no cover - 핀 버전에 lxml 포함
        return ValidationResult("skipped", f"lxml 을 불러올 수 없습니다: {exc}", xsd_path=str(path))
    try:
        xmlschema = etree.XMLSchema(etree.parse(str(path)))
    except Exception as exc:  # XMLSchemaParseError, OSError …
        return ValidationResult("skipped", f"XSD 를 읽을 수 없습니다: {path} ({exc})", xsd_path=str(path))
    try:
        doc = etree.fromstring(raw)
    except etree.XMLSyntaxError as exc:
        return ValidationResult("fail", f"XML 구문 오류: {exc}", [str(exc)], xsd_path=str(path))
    if xmlschema.validate(doc):
        return ValidationResult("pass", xsd_path=str(path))
    errors = [f"{e.line}행: {e.message}" for e in xmlschema.error_log]
    return ValidationResult("fail", errors[0] if errors else "검증 실패", errors, xsd_path=str(path))


__all__ = [
    "ValidationResult",
    "subtree_range",
    "subtree_size",
    "check_wellformed",
    "count_elements",
    "expected_signatures",
    "structural_equivalence",
    "find_xsd_dir",
    "xsd_path_for",
    "xsd_validate",
]
