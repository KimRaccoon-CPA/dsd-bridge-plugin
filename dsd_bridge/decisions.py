"""판정 기록 — 사람이 확인한 경보를 기억해 다음 실행·다음 기수에 다시 띄우지 않는다 (0.6.0).

검사 결과(합계 불일치·부호주의, 정산표 대사 불일치·주의·`정산표에만 있음`, 확인 필요, 주석 참조,
전기와 같은 표·문단)마다 **안정된 번호**(`#1a2b3c4d`)를 붙인다. 번호는 시트·계정과목·검사 종류처럼
기수가 바뀌어도 그대로인 것으로만 만든다(금액·행 번호는 넣지 않는다). 그래서 같은 보고기업의
다음 기수에도 같은 경보에는 같은 번호가 붙는다.

사람이 확인한 번호는 판정 기록 파일(JSON)에 사유와 함께 적는다::

    dsd-bridge accept 판정.json 1a2b3c4d --reason "부호주의: 현금유출을 양수로 적는 관행"

`check`·`diff` 에 `--decisions 판정.json` 을 주면

- 기록된 경보는 본문 목록에서 빠지고 `── 확인 처리 n건 ──` 블록에 사유와 함께 따로 나온다
- 기록된 합계·대사 불일치는 종료 코드 8 에 세지 않는다
- 나머지 경보 줄 끝에 번호 ` [#…]` 를 붙인다 (`--decisions` 가 없으면 출력은 예전과 같다)

상태는 두 가지다: `확인함`(수정할 필요 없음) / `대상 아님`(이 검사를 적용하지 않는 칸).
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from .errors import InputError, OutputError

#: 상태.
CONFIRMED = "확인함"
NOT_APPLICABLE = "대상 아님"
STATUSES = (CONFIRMED, NOT_APPLICABLE)

_RE_WS = re.compile(r"[\s　]+")
_RE_ID = re.compile(r"^[0-9a-f]{8}$")


def item_id(*parts) -> str:
    """안정된 8자리 번호 — 각 부분의 공백을 지우고 이어 SHA-1 앞 8자리."""
    text = "|".join(_RE_WS.sub("", "" if p is None else str(p)) for p in parts)
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:8]


def tag(i: str) -> str:
    return f" [#{i}]"


@dataclass
class Decisions:
    """판정 기록 파일 하나."""

    path: Path
    items: dict = field(default_factory=dict)     # 번호 → {"status", "reason", "what", "date"}

    def get(self, i: str):
        return self.items.get(i)

    def accepted(self, i: str) -> bool:
        return i in self.items

    def reason_ko(self, i: str) -> str:
        d = self.items.get(i) or {}
        return f"{d.get('status', CONFIRMED)}: {d.get('reason', '')}".rstrip(": ")

    def add(self, i: str, reason: str, status: str = CONFIRMED, what: str = "") -> None:
        if not _RE_ID.match(i or ""):
            raise InputError(f"오류: 판정 번호는 8자리 영숫자여야 합니다(출력의 [#…] 안 값): {i}")
        if status not in STATUSES:
            raise InputError(f"오류: 상태는 {', '.join(STATUSES)} 중 하나여야 합니다: {status}")
        if not (reason or "").strip():
            raise InputError("오류: --reason 으로 확인한 사유를 적어 주세요.")
        self.items[i] = {"status": status, "reason": reason.strip(), "what": what.strip(),
                         "date": _dt.date.today().isoformat()}

    def remove(self, i: str) -> bool:
        return self.items.pop(i, None) is not None

    def save(self) -> None:
        obj = {"설명": "DSD브릿지 판정 기록 — dsd-bridge accept 로 추가. 번호는 검사 출력의 [#…] 값.",
               "items": [dict(id=k, **v) for k, v in sorted(self.items.items())]}
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            raise OutputError(f"오류: 판정 기록을 쓸 수 없습니다: {self.path}\n  사유: {exc}") from exc


def load(path) -> Decisions:
    """판정 기록을 읽는다. 파일이 없으면 빈 기록(저장하면 새로 만든다)."""
    p = Path(path)
    out = Decisions(p)
    if not p.exists():
        return out
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise InputError(f"오류: 판정 기록을 읽을 수 없습니다: {p}\n  사유: {exc}") from exc
    for it in (data.get("items") if isinstance(data, dict) else None) or []:
        if isinstance(it, dict) and _RE_ID.match(str(it.get("id", ""))):
            out.items[it["id"]] = {k: it.get(k, "") for k in ("status", "reason", "what", "date")}
    return out


def split(items, key, decisions):
    """`(남은 것, 확인 처리된 것)` — `decisions` 가 없으면 전부 남는다."""
    if decisions is None:
        return list(items), []
    keep, done = [], []
    for it in items:
        (done if decisions.accepted(key(it)) else keep).append(it)
    return keep, done


def accepted_lines(done, key, describe, decisions) -> list:
    """`── 확인 처리 n건 ──` 블록."""
    if not done:
        return []
    lines = [f"── 확인 처리 {len(done)}건 (판정 기록) ──"]
    lines.extend(f"[확인 처리] {describe(it)} — {decisions.reason_ko(key(it))}{tag(key(it))}" for it in done)
    return lines


__all__ = ["CONFIRMED", "NOT_APPLICABLE", "STATUSES", "Decisions", "item_id", "tag", "load", "split",
           "accepted_lines"]
