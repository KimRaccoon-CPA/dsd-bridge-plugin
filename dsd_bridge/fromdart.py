"""`dsd-bridge from-dart` — 공시된 감사보고서 원문(OpenDART)으로 전기 틀 .dsd 만들기 (0.7.0).

전기 .dsd 파일을 받지 못했어도 그 보고기업의 전기 감사보고서가 DART 에 공시돼 있으면, OpenDART
"공시서류원본파일"(접수번호로 내려받는 원문 XML)로 틀을 만들 수 있다. 원문 XML 은 DART 편집기 파일과 같은
서식(dart4.xsd, ACODE·ADELIM·AUNIT 포함)이지만 공시용으로 가공돼 그대로는 열리지 않는다. 실측한 차이:

- 머리말을 감싸는 `<DOCUMENT-HEADER>` 가 빠져 있다 → 다시 감싼다
- 본문의 `<당기말>` 처럼 글자로 쓴 꺾쇠가 이스케이프되지 않았다 → `&lt;` 로 바꾼다 (같은 이유로 `&`)
- 목차를 감싸는 `<TOC>` 가 빠져 있다 → 목차 제목부터 첫 쪽나눔까지 다시 감싼다 (0.7.1)
- 감사인의 감사보고서·재무제표 `LIBRARY` 를 감싸는 `<INSERTION>`(편집기 서식 연결 정보)과 LIBRARY 첫 자식
  `<FILENAME ATYPE="dsl">01_재무상태표</FILENAME>`(어느 단위서식인지)이 빠져 있다 → LIBRARY 마다 편집기 서식의 값
  (ABASISNUMBER·COMMENT·LIBRARYLIST·FILENAME)으로 되살린다 (0.7.1, 회사노트북 DART 편집기 5.1 시험):
  · 한 INSERTION 에 여럿을 넣거나 연결 정보를 비우면 파일형식 검사가 "삽입된 단위서식이 지정된 개수보다 많습니다"
  · 연결 정보는 채웠는데 FILENAME 이 없으면 파일형식 검사 때 편집기가 강제 종료된다(ucrtbase 0xc0000409)
- 쓰지 않은 단위서식 자리표시(정오표, 재무상태표-각주, 신탁계정용, 부속명세서 목차, 내부회계관리제도 의견, 감사인의
  중요성 금액 — `AFREQUENCY="0"` 인 빈 INSERTION)가 모두 지워져 있다 → 편집기 빈 서식과 같은 자리에 되살린다. 없으면
  파일형식 검사가 "보고서에 문제가 발생했습니다(주석 문단)" 을 낸다 (0.7.1)
- 금액 칸 끝에 정렬용 공백이 붙어 있다(`12,864,156 `) → 숫자로 읽히도록 공백을 뗀다 (0.7.1)
- 회사 이름 글자·요약값은 비어 있다(회사 고유번호는 있어 편집기가 회사를 알아본다)
- 문단 안 줄바꿈(`&cr;`)이 지워져 있다 → 대부분 **복원할 수 없다**. 주석 번호 제목 조각(`<SPAN>N. 제목</SPAN>`) 바로
  뒤에 본문이 붙은 곳만 줄바꿈을 넣는다 (0.7.1). 표·계정코드·주석 구조는 그대로다

입력은 접수번호(14자리, OpenDART 인증키 필요) 또는 내려받은 원문 `.xml`/`.zip` 이다.
`--search 회사명` 은 그 회사의 최근 외부감사 관련 공시(접수번호)를 보여 준다.

인증키: `--key` → 환경변수 `DART_API_KEY` → `~/.dsd-bridge/config.json` 의 `dart_api_key`.
"""

from __future__ import annotations

import datetime as _dt
import io
import json
import os
import re
import time
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from .errors import InputError, OutputError

API = "https://opendart.fss.or.kr/api"
UA = "Mozilla/5.0 (dsd-bridge)"
#: 감사보고서 문서코드 (DOCUMENT-NAME ACODE).
AUDIT_REPORT = "00760"
_HEADER_TAGS = ("DOCUMENT-NAME", "FORMULA-VERSION", "COMPANY-NAME", "SUMMARY")
_TAGLIKE = re.compile(r"</?[A-Z][A-Z0-9\-]*[\s/>]|<[?!]")
_ENTITY = re.compile(r"&(?:[A-Za-z][A-Za-z0-9]*|#\d+|#x[0-9A-Fa-f]+);")


def api_key(cli_key=None) -> str:
    if cli_key:
        return cli_key.strip()
    if os.environ.get("DART_API_KEY"):
        return os.environ["DART_API_KEY"].strip()
    cfg = Path.home() / ".dsd-bridge" / "config.json"
    if cfg.is_file():
        try:
            k = str(json.loads(cfg.read_text(encoding="utf-8")).get("dart_api_key", "")).strip()
        except (OSError, ValueError):
            k = ""
        if k:
            return k
    raise InputError("오류: OpenDART 인증키가 없습니다. --key, 환경변수 DART_API_KEY, 또는 "
                     "~/.dsd-bridge/config.json 의 \"dart_api_key\" 에 넣어 주세요(https://opendart.fss.or.kr 에서 무료 발급).")


def _get(path: str, params: dict, timeout: int = 60) -> bytes:
    url = f"{API}/{path}?{urllib.parse.urlencode(params)}"
    last = None
    for i in range(3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read()
        except Exception as exc:          # noqa: BLE001 - 네트워크 오류는 다시 시도
            last = exc
            time.sleep(1.5 * (i + 1))
    raise InputError(f"오류: OpenDART 요청이 실패했습니다({path}): {last}")


def repair(raw: bytes) -> tuple:
    """공시 원문 XML → DART 편집기 형식 XML. 반환 `(바이트, 고친 내용 목록)`."""
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp949")
        text = re.sub(r'encoding="[^"]+"', 'encoding="utf-8"', text, count=1)
    notes = []
    out, pos, n_lt, n_amp = [], 0, 0, 0
    for m in re.finditer(r"[<&]", text):
        j = m.start()
        out.append(text[pos:j])
        if text[j] == "<":
            if _TAGLIKE.match(text, j):
                out.append("<")
            else:
                out.append("&lt;")
                n_lt += 1
        else:
            if _ENTITY.match(text, j):
                out.append("&")
            else:
                out.append("&amp;")
                n_amp += 1
        pos = j + 1
    out.append(text[pos:])
    text = "".join(out)
    if n_lt or n_amp:
        notes.append(f"본문에 글자로 쓴 꺾쇠·앰퍼샌드 {n_lt + n_amp}곳을 이스케이프했습니다(<당기말> 등)")
    if "<DOCUMENT-HEADER" not in text:
        first = min((text.find(f"<{t}") for t in _HEADER_TAGS if f"<{t}" in text), default=-1)
        end = text.find("</SUMMARY>")
        if first >= 0 and end >= 0:
            end += len("</SUMMARY>")
            text = text[:first] + '<DOCUMENT-HEADER AEXT-CLASS="Y">\r\n' + text[first:end] + \
                "\r\n</DOCUMENT-HEADER>" + text[end:]
            notes.append("빠진 머리말 묶음(DOCUMENT-HEADER)을 다시 감쌌습니다")
    if not text.lstrip().startswith("<?xml"):
        text = '<?xml version="1.0" encoding="utf-8"?>\r\n' + text
    text, n_num = _RE_PADDED_AMOUNT.subn(r"\1\2\3", text)
    if n_num:
        notes.append(f"금액 칸 {n_num}곳 끝의 정렬용 공백을 떼어 숫자로 읽히게 했습니다")
    text, n_head = _RE_HEADING_GLUED.subn(r"\1&amp;cr;", text)
    if n_head:
        notes.append(f"주석 번호 제목 바로 뒤에 붙은 본문 {n_head}곳에 줄바꿈을 넣었습니다")
    data = text.encode("utf-8")
    data, toc = _wrap_toc(data)
    if toc:
        notes.append("빠진 목차 묶음(TOC)을 다시 감쌌습니다")
    code = (re.search(rb'<DOCUMENT-NAME\s+ACODE="(\d+)"', data) or [None, AUDIT_REPORT.encode()])[1].decode()
    data, wrapped, unknown = _wrap_libraries(data, code)
    if wrapped:
        notes.append(f"INSERTION 밖에 떨어진 LIBRARY {wrapped}개를 편집기 서식 연결 정보(단위서식 이름 FILENAME 포함)로 하나씩 감쌌습니다"
                     + (f"(그중 {unknown}개는 어떤 서식인지 몰라 연결 정보를 비워 둠)" if unknown else ""))
    if code == AUDIT_REPORT:
        data, n_ph = _add_placeholders(data, code)
        if n_ph:
            notes.append(f"지워져 있던 쓰지 않은 단위서식 자리표시 {n_ph}개(정오표·각주·신탁계정용 등)를 되살렸습니다")
    return data, notes


#: 정렬용 공백이 붙은 금액 칸 (`<TE …>12,864,156 </TE>`, `- `). 전각 공백(빈 칸 표시)은 건드리지 않는다.
_RE_PADDED_AMOUNT = re.compile(r"(<T[DE]\b[^>]*>) *(\(?-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?|-) +(</T[DE]>)")
#: 주석 번호 제목 조각 바로 뒤에 본문 글자가 붙은 곳 (`<P><SPAN …>4. 제목</SPAN>4-1 …`)
_RE_HEADING_GLUED = re.compile(r"(<P\b[^>]*>\s*<SPAN\b[^>]*>\s*\d{1,2}\.\s*[^<]{1,80}</SPAN>)(?=[^<\s&])")
#: 편집기 감사보고서 서식(00760, 편집기 5.1 서식 6.0)의 INSERTION·FILENAME 값.
#: (판별, ABASISNUMBER, AFREQUENCY, COMMENT, LIBRARYLIST 꼴, FILENAME 후보 [(조건, 이름)]).
#: 편집기로 만든 감사보고서 2개와 편집기 새 문서(빈 서식·재무상태표 삽입)에서 실측. `{code}` 는 문서코드.
_FORM_INSERTIONS = (
    ("title:독립된감사인의감사보고서", "E1", "1", "◆click◆ 『감사인의 감사보고서』 삽입",
     "{code}#*감사인의 감사보고서.dsl", (("", "61_독립된 감사인의 감사보고서"),)),
    ("title:외부감사인의감사보고서", "E1", "1", "◆click◆ 『감사인의 감사보고서』 삽입",
     "{code}#*감사인의 감사보고서.dsl", (("", "62_외부감사인의 감사보고서"),)),
    ("group:PBS", "E1", "1", "◆click◆『재무상태표』 삽입", "{code}_BS#*.dsl", (("", "01_재무상태표"),)),
    ("group:PIS", "E1", "1", "◆click◆『손익계산서』 삽입", "{code}_IS#*.dsl",
     (("title:포괄", "11_포괄손익계산서"), ("", "10_손익계산서"))),
    ("group:PEF", "E1", "1", "◆click◆『자본변동표』 삽입", "{code}_PEF#*.dsl", (("", "20_자본변동표"),)),
    ("group:PCF", "SE1", "1", "◆click◆『현금흐름표』 삽입", "{code}_CF#*.dsl",
     (("group:PCF-D", "17_현금흐름표-직접법"), ("", "16_현금흐름표-간접법"))),
    ("title:내부회계관리제도감사또는검토의견", "SE1", "1", "◆click◆ 『내부회계관리제도 감사 또는 검토의견』 삽입",
     "{code}#*내부회계관리제도감사또는검토의견.dsl", (("", "50_내부회계관리제도감사또는검토의견"),)),
    ("title:5.감사인의중요성금액", "SE1", "1", "◆click◆ 『감사인의 중요성 금액』 삽입",
     "{code}#*감사인의중요성금액.dsl", (("", "70_감사인의중요성금액"),)),
)

#: 쓰지 않은 단위서식 자리표시 (편집기 5.1 감사보고서 빈 서식에서 실측). (자리, 기준, [(ABASISNUMBER, COMMENT, LIBRARYLIST)])
#: 자리: before = 기준 요소 앞, after = 기준 INSERTION 뒤, after_section = 기준 INSERTION 이 든 SECTION-1 뒤,
#: section_end = 기준 제목(AASSOCNOTE)이 든 SECTION-1 의 끝.
_FORM_PLACEHOLDERS = (
    ("before", "tag:COVER", (("SE1", "◆click◆ 정정문서 작성시 『정오표』 삽입", "일반-정정*.LCommon"),)),
    ("after", "list:{code}_BS#*.dsl", (("SE1", "◆click◆『재무상태표-각주』 삽입", "{code}_NOTE#*.dsl"),
                                      ("SE1", "◆click◆『재무상태표(신탁계정용)』 삽입", "{code}_BS-1#*.dsl"))),
    ("after", "list:{code}_IS#*.dsl", (("SE1", "◆click◆『손익계산서(신탁계정용)』 삽입", "{code}_IS-1#*.dsl"),)),
    ("after", "list:{code}_CF#*.dsl", (("SE1", "◆click◆『현금흐름표(신탁계정용)』 삽입", "{code}_CF-1#*.dsl"),)),
    ("after_section", "list:{code}_BS#*.dsl",
     (("SE1", "◆click◆ 『부속명세서 목차』 삽입", "{code}#*부속명세서_목차.dsl"),
      ("SE1", "◆click◆ 『내부회계관리제도 감사 또는 검토의견』 삽입", "{code}#*내부회계관리제도감사또는검토의견.dsl"))),
    ("section_end", "note:D-0-2-0-0", (("SE1", "◆click◆ 『감사인의 중요성 금액』 삽입", "{code}#*감사인의중요성금액.dsl"),)),
)


def _form_insertion(inner: bytes, code: str):
    """LIBRARY 내용 → (ABASISNUMBER, AFREQUENCY, COMMENT, LIBRARYLIST, FILENAME) 또는 None(모르는 서식)."""
    text = inner.decode("utf-8", "replace")
    title = re.search(r"<TITLE[^>]*>(.*?)</TITLE>", text, re.S)
    title = re.sub(r"\s+", "", title.group(1)) if title else ""
    group = re.search(r'<TABLE-GROUP\s+ACLASS="([^"]+)"', text)
    group = group.group(1) if group else ""

    def hit(cond: str) -> bool:
        if not cond:
            return True
        kind, val = cond.split(":", 1)
        return (kind == "title" and val in title) or (kind == "group" and group.startswith(val))

    for key, basis, freq, comment, lst, names in _FORM_INSERTIONS:
        kind, val = key.split(":", 1)
        if (kind == "title" and title == val) or (kind == "group" and group.startswith(val)):
            filename = next(name for cond, name in names if hit(cond))
            return basis, freq, comment, lst.format(code=code), filename
    return None


def _wrap_libraries(xml: bytes, code: str = AUDIT_REPORT) -> tuple:
    """`INSERTION` 밖에 떨어진 `LIBRARY` 를 **하나씩** 편집기 서식 값의 `INSERTION` 으로 감싼다.

    반환 `(바이트, 감싼 수, 서식을 몰라 값을 비운 수)`.
    """
    from .text import escape_text
    from .xmlsplice import SourceIndex

    index = SourceIndex.build(xml)
    spans = index.spans
    libs = [sp for sp in spans
            if sp.tag == "LIBRARY" and sp.parent_ordinal >= 0 and spans[sp.parent_ordinal].tag != "INSERTION"]
    out, unknown = xml, 0
    for sp in reversed(libs):
        lib = xml[sp.start_lt:sp.end_gt]
        form = _form_insertion(lib, code)
        if form is None:
            unknown += 1
            form = ("", "1", "", "", "")
        basis, freq, comment, lst, filename = form
        if filename and b"<FILENAME" not in lib[:200]:
            cut = lib.index(b">") + 1
            lib = (lib[:cut] + f'\r\n<FILENAME ATYPE="dsl">{escape_text(filename)}</FILENAME>'.encode("utf-8")
                   + lib[cut:])
        head = (f'<INSERTION ABASISNUMBER="{basis}" AFREQUENCY="{freq}" ADUPLICATION="N">\r\n'
                f"<COMMENT>{escape_text(comment)}</COMMENT>\r\n<LIBRARYLIST>{escape_text(lst)}</LIBRARYLIST>\r\n")
        out = out[:sp.start_lt] + head.encode("utf-8") + lib + b"\r\n</INSERTION>" + out[sp.end_gt:]
    return out, len(libs), unknown


def _add_placeholders(xml: bytes, code: str) -> tuple:
    """편집기 빈 서식에 있는 쓰지 않은 단위서식 자리표시(AFREQUENCY="0")를 같은 자리에 되살린다. 반환 `(바이트, 넣은 수)`."""
    from .text import escape_text
    from .xmlsplice import SourceIndex

    ix = SourceIndex.build(xml)
    spans = ix.spans
    lists = {}
    for sp in spans:
        if sp.tag == "INSERTION":
            m = re.search(rb"<LIBRARYLIST>(.*?)</LIBRARYLIST>", xml[sp.start_lt:sp.end_gt], re.S)
            if m:
                lists.setdefault(m.group(1).decode("utf-8", "replace"), sp)

    def section_of(sp):
        o = sp.parent_ordinal
        while o >= 0 and spans[o].tag != "SECTION-1":
            o = spans[o].parent_ordinal
        return spans[o] if o >= 0 else None

    edits = []
    for where, key, items in _FORM_PLACEHOLDERS:
        items = [(b, c, lst.format(code=code)) for b, c, lst in items]
        items = [it for it in items if it[2] not in lists]
        if not items:
            continue
        kind, val = key.split(":", 1)
        val = val.format(code=code)
        offset = None
        if kind == "tag":
            hit = next((sp for sp in spans if sp.tag == val), None)
            offset = hit.start_lt if hit is not None and where == "before" else None
        elif kind == "list" and val in lists:
            anchor = lists[val]
            if where == "after":
                offset = anchor.end_gt
            elif where == "after_section":
                sec = section_of(anchor)
                offset = sec.end_gt if sec is not None else None
        elif kind == "note":
            for sp in spans:
                if sp.tag == "SECTION-1" and f'AASSOCNOTE="{val}"'.encode() in xml[sp.start_lt:sp.start_lt + 600]:
                    offset = sp.end_lt
                    break
        if offset is None:
            continue
        body = "".join(f'\r\n<INSERTION ABASISNUMBER="{b}" AFREQUENCY="0" ADUPLICATION="N">\r\n'
                       f"<COMMENT>{escape_text(c)}</COMMENT>\r\n<LIBRARYLIST>{escape_text(lst)}</LIBRARYLIST>\r\n"
                       f"</INSERTION>\r\n" for b, c, lst in items)
        edits.append((offset, body.encode("utf-8"), len(items)))
    for offset, body, _ in sorted(edits, key=lambda e: e[0], reverse=True):
        xml = xml[:offset] + body + xml[offset:]
    return xml, sum(n for _, _, n in edits)


def _wrap_toc(xml: bytes) -> tuple:
    """본문 바로 아래의 목차 제목(`TITLE ATOC="N"`, 글자 '목차')부터 첫 쪽나눔까지를 `<TOC>` 로 감싼다."""
    from .xmlsplice import SourceIndex

    if b"<TOC" in xml:
        return xml, False
    index = SourceIndex.build(xml)
    spans = index.spans
    for sp in spans:
        if sp.tag != "TITLE" or sp.parent_ordinal < 0 or spans[sp.parent_ordinal].tag != "BODY":
            continue
        seg = xml[sp.start_lt:sp.end_gt].decode("utf-8", "replace")
        if 'ATOC="N"' not in seg[:seg.index(">")] or re.sub(r"\s+|<[^>]+>", "", seg) != "목차":
            continue
        sibs = spans[sp.parent_ordinal].child_ordinals
        k = sibs.index(sp.ordinal)
        last = None
        for o in sibs[k + 1:]:
            tag = spans[o].tag
            if tag not in ("P", "TABLE", "TABLE-GROUP", "PGBRK"):
                break
            last = o
            if tag == "PGBRK":
                break
        if last is None or spans[last].tag != "PGBRK":
            return xml, False
        start, end = sp.start_lt, spans[last].end_gt
        return xml[:start] + b"<TOC>\r\n" + xml[start:end] + b"\r\n</TOC>" + xml[end:], True
    return xml, False


#: 편집기 파일의 meta.xml 에 적힌 값 (실측: dart4 = DART 편집기 5.1 이 다시 저장한 값 5.118·6.0, dart3 = 3.051·3.8).
_EDITOR_VERSIONS = {"dart4.xsd": ("5.118", "6.0"), "dart3.xsd": ("3.051", "3.8")}


def make_meta(xml: bytes) -> bytes:
    """편집기 파일과 같은 모양의 최소 meta.xml (회사 고유번호·이름·크기·요약값은 비움)."""
    text = xml.decode("utf-8", "replace")
    schema = (re.search(r'noNamespaceSchemaLocation="([^"]+)"', text) or [None, "dart4.xsd"])[1]
    code = (re.search(r'<DOCUMENT-NAME\s+ACODE="(\d+)"', text) or [None, AUDIT_REPORT])[1]
    editver, docver = _EDITOR_VERSIONS.get(schema.lower(), _EDITOR_VERSIONS["dart4.xsd"])
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<METAINFO>\r\n'
            f'\t<GENERATOR schema="{schema}" editver="{editver}"/>\r\n'
            f'\t<DOCUMENT-HEADER regcik="" regname=""/>\r\n'
            f'\t<DOCUMENT-INFO docver="{docver}" bsn-id="{code}" rpt-id="{code}" doc-id="{code}" iscorrection="N"/>\r\n'
            f'\t<FILE-INFO date="" size="" rncheck="" rnimgfile="" rnnumber=""/>\r\n'
            f'</METAINFO>').encode("utf-8")


def _pick(names_and_data: list, want: str = "") -> tuple:
    """원문 zip 의 여러 XML 중 하나 — 이름 지정 > 감사보고서(00760) > 가장 큰 것."""
    xmls = [(n, d) for n, d in names_and_data if n.lower().endswith(".xml")]
    if not xmls:
        raise InputError("오류: 원문에 XML 파일이 없습니다.")
    if want:
        hit = [(n, d) for n, d in xmls if want in n]
        if not hit:
            raise InputError(f"오류: 원문에 '{want}' 파일이 없습니다. 있는 파일: {', '.join(n for n, _ in xmls)}")
        return hit[0]
    audit = [(n, d) for n, d in xmls if re.search(rb'<DOCUMENT-NAME\s+ACODE="00760"', d[:4000])]
    return (audit or sorted(xmls, key=lambda t: -len(t[1])))[0]


@dataclass
class FromDartReport:
    output: Path
    source: str
    picked: str = ""
    files: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    xsd: str = ""

    def lines_ko(self) -> list:
        lines = [f"전기 틀 만들기: {self.source} → {self.output}",
                 f"원문 파일: {self.picked}" + (f" (원문에 든 파일 {len(self.files)}개: {', '.join(self.files)})"
                                             if len(self.files) > 1 else "")]
        lines.extend(f"고침: {n}" for n in self.notes)
        lines.append(f"XSD 검증: {self.xsd}")
        lines.append("안내: 공시 원문에는 문단 안 줄바꿈(&cr;)이 없어 주석 번호 제목 뒤 말고는 복원되지 않습니다. 표·계정코드·"
                     "주석 구조는 그대로이니 export 해 당기 내용으로 채우고, 붙어 있는 문단(예: '4-1 …' 소제목)은 줄을 나눠 "
                     "다듬으세요. 요약값은 비어 있어 DART 편집기에서 재추출이 필요합니다. 편집기에서 한 번 열어 저장한 파일을 "
                     "전기 .dsd 로 쓰고, 재무제표 검사(F9)의 '계정과목 오류'는 공시된 보고서에도 있던 비표준 계정이니 확인만 하세요.")
        return lines


def from_dart(source: str, out, *, key=None, doc: str = "") -> FromDartReport:
    """접수번호 또는 원문 파일 → .dsd."""
    dst = Path(out)
    src = str(source)
    if re.fullmatch(r"\d{14}", src):
        data = _get("document.xml", {"crtfc_key": api_key(key), "rcept_no": src})
        try:
            zf = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile:
            msg = re.search(rb"<message>(.*?)</message>", data)
            raise InputError("오류: OpenDART 원문을 받지 못했습니다: "
                             + (msg.group(1).decode("utf-8", "replace") if msg else data[:120].decode("utf-8", "replace")))
        items = [(n, zf.read(n)) for n in zf.namelist()]
    else:
        p = Path(src)
        if not p.is_file():
            raise InputError(f"오류: 접수번호(14자리)나 원문 파일 경로가 아닙니다: {src}")
        if zipfile.is_zipfile(p):
            with zipfile.ZipFile(p) as zf:
                items = [(n, zf.read(n)) for n in zf.namelist()]
        else:
            items = [(p.name, p.read_bytes())]
    name, raw = _pick(items, doc)
    xml, notes = repair(raw)
    from .validate import check_wellformed, xsd_validate

    check_wellformed(xml)
    report = FromDartReport(dst, src, name, [n for n, _ in items], notes)
    v = xsd_validate(xml, "dart4" if b"dart4.xsd" in xml[:800] else "dart3")
    report.xsd = v.status_ko + (f" ({v.reason})" if v.status not in ("pass", "skipped") and v.reason else "")
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("contents.xml", xml)
            z.writestr("meta.xml", make_meta(xml))
    except OSError as exc:
        raise OutputError(f"오류: 출력 파일을 쓸 수 없습니다: {dst}\n  사유: {exc}") from exc
    return report


# ─────────────────────────────────────────────────────────────────────────────
# 접수번호 찾기
# ─────────────────────────────────────────────────────────────────────────────


def _corp_list(key: str) -> list:
    cache = Path.home() / ".dsd-bridge" / "cache" / "corpCode.xml"
    if not cache.is_file() or (time.time() - cache.stat().st_mtime) > 7 * 86400:
        data = _get("corpCode.xml", {"crtfc_key": key}, timeout=180)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read(zf.namelist()[0])
        cache.parent.mkdir(parents=True, exist_ok=True)
        cache.write_bytes(xml)
    text = cache.read_bytes().decode("utf-8", "replace")
    return re.findall(r"<corp_code>(\d+)</corp_code>\s*<corp_name>([^<]*)</corp_name>", text)


def search(name: str, *, key=None, years: int = 3) -> list:
    """회사 이름(일부)으로 최근 외부감사 관련 공시 `[(회사, 접수번호, 보고서명, 접수일)]`."""
    k = api_key(key)
    needle = re.sub(r"\s+", "", name)
    corps = [(c, n) for c, n in _corp_list(k) if needle and needle in re.sub(r"\s+", "", n)]
    corps.sort(key=lambda t: (re.sub(r"\s+", "", t[1]) != needle, len(t[1])))       # 정확히 같은 이름 먼저
    if not corps:
        return []
    end = _dt.date.today()
    bgn = end.replace(year=end.year - years)
    out = []
    for code, corp in corps[:10]:
        js = json.loads(_get("list.json", {"crtfc_key": k, "corp_code": code, "bgn_de": bgn.strftime("%Y%m%d"),
                                           "end_de": end.strftime("%Y%m%d"), "pblntf_ty": "F",
                                           "page_count": 100}).decode("utf-8"))
        for it in js.get("list", []) or []:
            out.append((corp, it.get("rcept_no", ""), it.get("report_nm", ""), it.get("rcept_dt", "")))
    return out


__all__ = ["repair", "make_meta", "from_dart", "search", "api_key", "FromDartReport"]
