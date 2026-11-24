"""`.dsd` ZIP 컨테이너 읽기/쓰기 (계획 §4.12, AC-19, AC-20).

`.dsd` 는 항목 두 개(`contents.xml`, `meta.xml`)짜리 ZIP 이다. 이 모듈은 그 두 항목의
**원본 바이트**와 ZIP 항목 메타(순서·압축 방식·`date_time` 등)를 그대로 들고 있다가,
쓸 때 같은 형태로 되돌린다.

`meta.xml` 은 **절대 재직렬화하지 않는다**. `update_meta_size` 는 `FILE-INFO` 의
`size` 속성값만 바이트 단위 정규식으로 갈아끼운다 (AC-19).
"""

from __future__ import annotations

import hashlib
import os
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from .errors import InputError, OutputError

CONTENTS_NAME = "contents.xml"
META_NAME = "meta.xml"

#: 출력 항목 순서 — 원본과 동일하게 고정한다 (§4.12, AC-20).
ENTRY_ORDER = (CONTENTS_NAME, META_NAME)

#: `<FILE-INFO ... size="184888" ...>` 의 **size 속성값만** 잡는다.
#: `date` 는 어떤 경우에도 건드리지 않는다 (AC-19).
_FILE_INFO_SIZE = re.compile(rb'(<FILE-INFO\b[^>]*?\bsize=")([^"]*)(")')

#: `<GENERATOR schema="dart4.xsd" editver="5.036"/>` 에서 값을 뽑는다.
_GENERATOR = re.compile(rb"<GENERATOR\b[^>]*>")
_ATTR = re.compile(rb'([A-Za-z_][-A-Za-z0-9_.]*)="([^"]*)"')

#: `<?xml version="1.0" encoding="UTF-8"?>` 의 encoding 값 (선언이 없거나 encoding 이 없으면 None).
_XML_DECL_ENCODING = re.compile(rb"""^<\?xml[^>]*?\bencoding\s*=\s*["']([^"']*)["']""")

#: dsd-bridge 가 다루는 유일한 인코딩. 오프셋 산술·`escape_text`·`_ID` 체크섬이 모두 UTF-8 바이트를
#: 전제하므로, 다른 인코딩의 문서는 조용히 깨뜨리는 대신 거부한다 (P5).
UTF8_NAMES = frozenset({"utf-8", "utf8"})


@dataclass(frozen=True)
class EntryInfo:
    """ZIP 항목 하나의 메타 (원본 `ZipInfo` 에서 뽑은, 재포장에 필요한 전부)."""

    name: str
    order: int
    compress_type: int
    create_system: int
    extract_version: int
    flag_bits: int
    external_attr: int
    extra: bytes
    comment: bytes
    date_time: tuple
    file_size: int
    compress_size: int
    crc: int


class DsdContainer:
    """`.dsd` 컨테이너 하나."""

    __slots__ = ("path", "contents_bytes", "meta_bytes", "entry_info", "zip_comment")

    def __init__(
        self,
        path,
        contents_bytes: bytes,
        meta_bytes: bytes,
        entry_info: dict,
        zip_comment: bytes = b"",
    ) -> None:
        self.path = path
        self.contents_bytes = contents_bytes
        self.meta_bytes = meta_bytes
        self.entry_info = entry_info      # 이름 → EntryInfo (원본 순서 유지)
        self.zip_comment = zip_comment

    # 읽기 ────────────────────────────────────────────────────────────────
    @classmethod
    def read(cls, path) -> "DsdContainer":
        """`.dsd` 를 읽는다. 입력 문제는 모두 `InputError`(종료 코드 2).

        - 경로가 없다 / 폴더다
        - ZIP 이 아니다 (또는 깨졌다)
        - ZIP 이지만 `contents.xml` 이 없다
        - `contents.xml` 의 XML 선언 인코딩이 UTF-8 이 아니다
        """
        p = Path(path)
        if not p.exists():
            raise InputError(f"입력 파일을 찾을 수 없습니다: {p}")
        if p.is_dir():
            raise InputError(f"입력 경로가 파일이 아니라 폴더입니다: {p}")
        try:
            zf = zipfile.ZipFile(p)
        except zipfile.BadZipFile as exc:
            raise InputError(
                f"입력 파일이 ZIP 형식이 아닙니다(.dsd 가 맞는지 확인하세요): {p} ({exc})"
            ) from exc
        except OSError as exc:
            raise InputError(f"입력 파일을 열 수 없습니다: {p} ({exc})") from exc

        with zf:
            names = zf.namelist()
            if CONTENTS_NAME not in names:
                raise InputError(
                    f"ZIP 안에 {CONTENTS_NAME} 이(가) 없습니다: {p}\n"
                    f"  들어 있는 항목: {', '.join(names) if names else '(없음)'}"
                )
            entry_info = {}
            for order, info in enumerate(zf.infolist()):
                entry_info[info.filename] = EntryInfo(
                    name=info.filename,
                    order=order,
                    compress_type=info.compress_type,
                    create_system=info.create_system,
                    extract_version=info.extract_version,
                    flag_bits=info.flag_bits,
                    external_attr=info.external_attr,
                    extra=info.extra,
                    comment=info.comment,
                    date_time=info.date_time,
                    file_size=info.file_size,
                    compress_size=info.compress_size,
                    crc=info.CRC,
                )
            try:
                contents = zf.read(CONTENTS_NAME)
                meta = zf.read(META_NAME) if META_NAME in names else b""
            except (zipfile.BadZipFile, OSError) as exc:
                raise InputError(f"ZIP 항목을 읽을 수 없습니다: {p} ({exc})") from exc
            zip_comment = zf.comment

        enc = xml_declaration_encoding(contents)
        if enc is not None and enc.lower() not in UTF8_NAMES:
            raise InputError(
                f"오류: 지원하지 않는 인코딩입니다 — {CONTENTS_NAME} 의 XML 선언이 "
                f'encoding="{enc}" 입니다: {p}\n'
                "  dsd-bridge 는 UTF-8 문서만 다룹니다. DART 편집기에서 다시 저장한 파일을 사용하세요."
            )
        return cls(p, contents, meta, entry_info, zip_comment)

    # 해시 ────────────────────────────────────────────────────────────────
    def contents_sha256(self) -> str:
        return hashlib.sha256(self.contents_bytes).hexdigest()

    def meta_sha256(self) -> str:
        return hashlib.sha256(self.meta_bytes).hexdigest()

    # meta.xml 속성 ───────────────────────────────────────────────────────
    def _generator_attrs(self) -> dict:
        m = _GENERATOR.search(self.meta_bytes)
        if not m:
            return {}
        return {
            k.decode("ascii", "replace"): v.decode("utf-8", "replace")
            for k, v in _ATTR.findall(m.group(0))
        }

    def schema(self) -> str | None:
        """`dart3` / `dart4` — `meta.xml` 의 `GENERATOR@schema` 에서 `.xsd` 를 뗀 값."""
        value = self._generator_attrs().get("schema")
        if not value:
            return None
        return value[:-4] if value.lower().endswith(".xsd") else value

    def editver(self) -> str | None:
        """DART 편집기 버전 — `GENERATOR@editver` (예: `5.036` / `3.051`)."""
        return self._generator_attrs().get("editver")

    # 쓰기 ────────────────────────────────────────────────────────────────
    def write(self, path, contents: bytes | None = None, meta: bytes | None = None) -> Path:
        """계획 §4.12 / AC-20 대로 재포장한다.

        - 항목 순서 `contents.xml` → `meta.xml`
        - `ZIP_DEFLATED(8)`, `create_system=0`, `extract_version=20`
        - `extra`/`comment` 는 빈 값 (원본이 빈 값)
        - `date_time` 은 원본 항목 값 그대로
        """
        out = Path(path)
        payload = {
            CONTENTS_NAME: self.contents_bytes if contents is None else contents,
            META_NAME: self.meta_bytes if meta is None else meta,
        }
        # 원자적 쓰기 — 같은 폴더의 임시 파일에 다 쓴 뒤 os.replace 로 바꿔치기한다.
        # 쓰다가 실패해도 잘린 .dsd 가 남지 않는다 (§4.10 "출력 파일을 만들지 않았습니다" 약속).
        tmp = out.with_name(out.name + ".dsd-bridge-tmp")
        try:
            with zipfile.ZipFile(tmp, "w") as z:
                for name in ENTRY_ORDER:
                    data = payload[name]
                    if not data and name not in self.entry_info:
                        continue          # 원본에 없던 빈 meta.xml 은 만들지 않는다
                    src = self.entry_info.get(name)
                    zi = zipfile.ZipInfo(
                        name,
                        date_time=src.date_time if src else (1980, 1, 1, 0, 0, 0),
                    )
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zi.create_system = 0
                    zi.external_attr = src.external_attr if src else 0
                    z.writestr(zi, data)
                    # zipfile 은 저장 후 extract_version 을 20 으로 둔다(원본과 일치).
                if self.zip_comment:
                    z.comment = self.zip_comment
            os.replace(tmp, out)
        except OSError as exc:
            try:
                tmp.unlink()
            except OSError:
                pass
            raise OutputError(
                f"출력 파일을 쓸 수 없습니다: {out} ({exc})\n"
                "  다른 프로그램이 파일을 열고 있는지 확인하세요."
            ) from exc
        return out


def update_meta_size(meta_bytes: bytes, n: int) -> bytes:
    """`meta.xml` 의 `FILE-INFO@size` 만 `n` 으로 바꾼다 (AC-19).

    규칙:

    - 원본 `size` 값이 **비어 있지 않을 때만** 치환한다.
      `size=""`(샘플B) 는 그대로 둔다.
    - `date` 는 어느 경우에도 건드리지 않는다.
    - 재직렬화하지 않는다 — 바이트 단위 정규식으로 값 부분만 갈아끼운다.
      값이 이미 같으면 결과는 원본과 **바이트 동일**이다.
    """
    if not meta_bytes:
        return meta_bytes
    new = str(int(n)).encode("ascii")

    def repl(m: "re.Match[bytes]") -> bytes:
        if not m.group(2):
            return m.group(0)          # size="" 는 유지
        return m.group(1) + new + m.group(3)

    return _FILE_INFO_SIZE.sub(repl, meta_bytes, count=1)


def xml_declaration_encoding(data: bytes) -> str | None:
    """XML 선언의 `encoding` 값 (선언이 없거나 `encoding` 속성이 없으면 `None`).

    XML 명세상 선언이 없으면 UTF-8 이므로 `None` 은 "UTF-8 로 본다" 는 뜻이다.
    """
    m = _XML_DECL_ENCODING.match(data)
    return None if m is None else m.group(1).decode("ascii", "replace")


def meta_file_info_size(meta_bytes: bytes) -> str | None:
    """`FILE-INFO@size` 의 현재 값 (없으면 `None`, 빈 값이면 빈 문자열)."""
    m = _FILE_INFO_SIZE.search(meta_bytes)
    return None if m is None else m.group(2).decode("ascii", "replace")


__all__ = [
    "DsdContainer",
    "EntryInfo",
    "UTF8_NAMES",
    "update_meta_size",
    "meta_file_info_size",
    "xml_declaration_encoding",
    "CONTENTS_NAME",
    "META_NAME",
    "ENTRY_ORDER",
]
