"""dsd-bridge 오류 계층과 종료 코드 (계획 §4.13).

모든 사용자 대면 오류는 `DsdBridgeError` 의 자손이며 두 가지를 반드시 갖는다.

- `exit_code` — 프로세스 종료 코드 (0 성공 / 2 입력 / 3 해시 / 4 시트 무결성 /
  5 미지원 편집 / 6 패치 정합성 / 7 출력 / 8 합계 불일치)
- `message_ko` — 사용자에게 그대로 보여줄 **한국어** 메시지

CLI 는 `except DsdBridgeError as e: print(e.message_ko); sys.exit(e.exit_code)` 만 하면 된다.

코드 8 만 예외가 아니다 — `dsd-bridge check` 는 오류가 아니라 **검사 결과**로 이 코드를 돌려주므로
(`FOOTING_MISMATCH_EXIT`) 메시지도 `오류:` 가 아니라 검사 보고서다.
"""

from __future__ import annotations

#: `dsd-bridge check` 가 합계 불일치를 하나라도 찾았을 때의 종료 코드 (계획 1.2 WP4).
#: 예외가 아니다 — 검사 자체는 성공했고 문서에 계산 오류가 있다는 뜻이다.
FOOTING_MISMATCH_EXIT = 8


class DsdBridgeError(Exception):
    """dsd-bridge 의 모든 사용자 대면 오류의 기반 클래스.

    `exit_code` 는 클래스 속성으로 고정되며, 필요하면 생성 시 덮어쓸 수 있다.
    """

    exit_code: int = 1

    def __init__(self, message_ko: str, exit_code: int | None = None) -> None:
        super().__init__(message_ko)
        self.message_ko = message_ko
        if exit_code is not None:
            self.exit_code = exit_code

    def __str__(self) -> str:  # pragma: no cover - 표현용
        return self.message_ko


class InputError(DsdBridgeError):
    """입력 파일 문제 — 없음 / ZIP 아님 / contents.xml 없음 / 통합문서 메타 문제."""

    exit_code = 2


class HashMismatch(DsdBridgeError):
    """원본 DSD 가 내보내기 당시와 다르다 (contents.xml 해시·요소 수 불일치)."""

    exit_code = 3


class SheetIntegrityError(DsdBridgeError):
    """시트 무결성 위반 — 대량 소실 / 재배열 / _ID 체크섬 / _SEQ·_ID 중복."""

    exit_code = 4


class UnsupportedEdit(DsdBridgeError):
    """이 도구가 지원하지 않는 편집 (혼합 콘텐츠 편집, 열 삽입, 표 신규 생성 등)."""

    exit_code = 5


class PatchIntegrityError(DsdBridgeError):
    """패치 내부 정합성 실패 — 구간 겹침 / well-formed 실패 / 구조 동등성 실패."""

    exit_code = 6


class OutputError(DsdBridgeError):
    """출력 경로 문제 — 쓸 수 없음 / 파일 잠김."""

    exit_code = 7


__all__ = [
    "FOOTING_MISMATCH_EXIT",
    "DsdBridgeError",
    "InputError",
    "HashMismatch",
    "SheetIntegrityError",
    "UnsupportedEdit",
    "PatchIntegrityError",
    "OutputError",
]
