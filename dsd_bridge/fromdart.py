"""`dsd-bridge from-dart` — 공시된 감사보고서 원문(OpenDART)으로 전기 틀 .dsd 만들기 (0.7.0).

전기 .dsd 파일을 받지 못했어도 그 보고기업의 전기 감사보고서가 DART 에 공시돼 있으면, OpenDART
"공시서류원본파일"(접수번호로 내려받는 원문 XML)로 틀을 만들 수 있다. 원문 XML 은 DART 편집기 파일과 같은
서식(dart4.xsd, ACODE·ADELIM·AUNIT 포함)이지만 공시용으로 가공돼 그대로는 열리지 않는다. 실측한 차이:

- 머리말을 감싸는 `<DOCUMENT-HEADER>` 가 빠져 있다 → 다시 감싼다
- 본문의 `<당기말>` 처럼 글자로 쓴 꺾쇠가 이스케이프되지 않았다 → `&lt;` 로 바꾼다 (같은 이유로 `&`)
- 회사 이름·요약값은 비어 있다 (틀로는 오히려 알맞다)
- 문단 안 줄바꿈(`&cr;`)이 지워져 있다 → **복원할 수 없다**. 표·계정코드·주석 구조는 그대로다

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
    data, n_ins = _wrap_libraries(text.encode("utf-8"))
    if n_ins:
        notes.append(f"INSERTION 밖에 떨어진 LIBRARY 묶음 {n_ins}곳을 INSERTION 으로 감쌌습니다"
                     "(편집기 서식 연결 정보 LIBRARYLIST 는 비어 있음)")
    return data, notes


def _wrap_libraries(xml: bytes) -> tuple:
    """`INSERTION` 밖에 떨어진 `LIBRARY` 연속 묶음을 `INSERTION` 으로 감싼다 (편집기 파일에는 늘 감싸져 있다)."""
    from .xmlsplice import SourceIndex

    index = SourceIndex.build(xml)
    spans = index.spans
    groups = []
    for sp in spans:
        if sp.tag != "LIBRARY" or sp.parent_ordinal < 0 or spans[sp.parent_ordinal].tag == "INSERTION":
            continue
        parent = spans[sp.parent_ordinal]
        siblings = parent.child_ordinals
        k = siblings.index(sp.ordinal)
        if groups and groups[-1][-1] == siblings[k - 1] if k else False:
            groups[-1].append(sp.ordinal)
        else:
            groups.append([sp.ordinal])
    out = xml
    for g in reversed(groups):
        start, end = spans[g[0]].start_lt, spans[g[-1]].end_gt
        out = (out[:start] + b'<INSERTION ABASISNUMBER="" AFREQUENCY="1" ADUPLICATION="N">\r\n'
               b"<COMMENT></COMMENT>\r\n<LIBRARYLIST></LIBRARYLIST>\r\n" + out[start:end]
               + b"\r\n</INSERTION>" + out[end:])
    return out, len(groups)


#: 편집기 파일의 meta.xml 에 적힌 값 (서식 버전별 실측: dart4 = 편집기 5.036·문서 4.1, dart3 = 3.051·3.8).
_EDITOR_VERSIONS = {"dart4.xsd": ("5.036", "4.1"), "dart3.xsd": ("3.051", "3.8")}


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
        lines.append("안내: 공시 원문에는 문단 안 줄바꿈(&cr;)이 없어 복원되지 않습니다. 표·계정코드·주석 구조는 "
                     "그대로이니 export 해 당기 내용으로 채우고, 문단은 새로 다듬으세요. 회사 이름·요약값은 비어 있어 "
                     "DART 편집기에서 회사 지정·요약값 재추출이 필요합니다. 이렇게 만든 파일이 DART 편집기에서 열리는지는 "
                     "아직 실제로 확인하지 않았으니, 처음 쓸 때 편집기로 열어 보고 열리면 편집기에서 저장한 파일을 틀로 쓰세요.")
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
