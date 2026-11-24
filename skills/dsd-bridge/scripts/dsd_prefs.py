# -*- coding: utf-8 -*-
"""사전재무제표(감사전 재무제표 .dsd, 문서코드 00860)를 당기 감사보고서 .dsd 로 만든다(실제 보고서 형식으로 검증한 방법).
사전재무제표 = DART 편집기 '감사보고서 변환' 결과물. 변환 규칙(감사보고서·사전재무제표 쌍으로 검증):
  ① ENG(영문) 속성 삭제 ② ACODE 값 비움 ③ 표 그룹명 PBS-2:A·PIS-2:A·PCF-I:A 의 ':A' 제거
전기 사전재무제표의 머리말·REP_KND·'재 무 제 표' 제목·부속명세서(INSERTION)·meta.xml 은 그대로 두고,
그 사이(표지 COVER2 ~ 주석 SECTION-2 끝)를 당기 감사보고서 '(첨부)재 무 제 표' SECTION-1 안쪽(규칙 적용)으로 바꾼다.
  python dsd_prefs.py 전기_사전재무제표.dsd 당기_감사보고서.dsd -o 당기_사전재무제표.dsd
검증(자동): XSD(dart4), 남은 ENG·ACODE 값·':A' 그룹 0. 뒤이어 사람이 dsd-bridge diff --rollforward / check 로 기간 표기·합계 확인.
감사보고서가 바뀌면 사전재무제표도 이 스크립트로 다시 만든다.
"""
import os, re, sys, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dsd_xml import xsd  # noqa: E402


def fs_region(t):
    """(첨부)재무제표 / 재무제표 SECTION-1 의 TITLE 직후 ~ 그 SECTION-1 닫힘 직전"""
    m = re.search(r'<SECTION-1 ACLASS="MANDATORY" APARTSOURCE="SOURCE">\s*<TITLE ATOC="Y" AASSOCNOTE="D-0-0-0-0"[^>]*>[^<]*</TITLE>', t)
    assert m, "재무제표 SECTION-1 제목을 못 찾음(양식이 다름)"
    s = m.end()
    n = t.find("<SECTION-2", s)
    e = t.find("</SECTION-2>", n)
    e2 = t.find("</SECTION-1>", e)
    assert t[e + len("</SECTION-2>"):e2].strip() == "", repr(t[e:e2 + 20])
    return m.start(), m.end(), e2


def to_pre(b):
    b = re.sub(r'\sENG="[^"]*"', "", b)
    b = re.sub(r'ACODE="[^"]*"', 'ACODE=""', b)
    b = re.sub(r'(<TABLE-GROUP ACLASS="[^":]+):A"', r'\1"', b)
    return b


def main():
    a = sys.argv[1:]
    pre_path, aud_path = a[0], a[1]
    out = a[a.index("-o") + 1]
    assert os.path.abspath(out) not in (os.path.abspath(pre_path), os.path.abspath(aud_path)), "입력 덮어쓰기 금지"
    assert not os.path.exists(out), f"같은 이름 파일이 이미 있음: {out}"
    zp = zipfile.ZipFile(pre_path)
    p = zp.read("contents.xml").decode("utf-8")
    t = zipfile.ZipFile(aud_path).read("contents.xml").decode("utf-8")
    _, pt, pe = fs_region(p)
    _, at, ae = fs_region(t)
    body = to_pre(t[at:ae])
    # 줄바꿈은 원본 그대로(CRLF): zip 에서 바이트로 읽어 디코드하므로 유지된다(텍스트 모드 파일 읽기는 LF 로 바꿈)
    new = p[:pt] + body + p[pe:]
    print("남은 ENG:", len(re.findall(r'\sENG=', body)), "| 값 있는 ACODE:", len(re.findall(r'ACODE="[^"]+"', body)),
          "| ':A' 그룹:", re.findall(r'ACLASS="[^"]*:A"', body))
    print("표 그룹:", re.findall(r'<TABLE-GROUP ACLASS="([^"]+)"', new))
    raw = new.encode("utf-8")
    print("XSD:", xsd(raw))
    with zipfile.ZipFile(out, "w") as zo:
        for info in zp.infolist():
            zo.writestr(info, raw if info.filename == "contents.xml" else zp.read(info.filename), compress_type=info.compress_type)
    print("저장:", out)


if __name__ == "__main__":
    main()
