# -*- coding: utf-8 -*-
"""DSD브릿지(dsd-bridge) import 가 못 하는 .dsd(contents.xml) 후처리 + XSD(dart4) 검증. 입력 파일은 덮어쓰지 않는다(-o 필수).

명령:
  info     <in.dsd>                                   기간 속성(AUNITVALUE)·COMM_DATE·자리표시(○)·콤마 누락·같은 ACODE 가 붙은 행 보고
  period   <in.dsd> -o <out.dsd> FROM TO               PERIODFROM*/SUB_PERIODFROM -> FROM, PERIODTO*/SUB_PERIODTO -> TO (YYYYMMDD)
  acode    <in.dsd> -o <out.dsd> "행 라벨" CODE [N번째]  라벨이 든 TR 의 모든 칸 ACODE 를 CODE 로(복제 삽입 행은 위 행 코드를 물려받음)
  replace  <in.dsd> -o <out.dsd> "옛 글" "새 글" [개수]   회색(편집 불가) 문단 등 글자만 바꾸기. 개수를 주면 정확히 그만큼일 때만
  commas   <in.dsd> -o <out.dsd>                       4자리 이상 숫자에 천 단위 콤마 넣기(전기 값이 짧아 콤마 없이 들어간 칸)
여러 작업은 출력 파일을 다음 입력으로 이어서 실행한다. 끝나면 'info' 로 다시 확인.
표준계정코드 표: C:\\DART\\Resource\\Code\\A\\PBS2.cod(재무상태표)·PIS2.cod(손익)·PCFI.cod(간접법 현금흐름). 없는 계정은 99999999999999(신설).
"""
import os, re, sys, zipfile


def _bridge_home():
    """DSD브릿지 프로그램 폴더를 찾는다.

    1) 플러그인으로 설치된 경우: 이 스크립트 기준 세 단계 위(플러그인 루트, scripts → dsd-bridge → skills → 루트)
    2) 환경변수 DSD_BRIDGE_HOME (옛 이름 DSDX_HOME)
    3) 흔한 복제 위치 ~/Work/dsd-bridge
    돌려주는 값: (폴더, 패키지 이름) 또는 (None, None).
    """
    here = os.path.dirname(os.path.abspath(__file__))
    plugin_root = os.path.abspath(os.path.join(here, "..", "..", ".."))
    cands = (
        plugin_root,
        os.environ.get("DSD_BRIDGE_HOME"), os.environ.get("DSDX_HOME"),
        os.path.expanduser("~/Work/dsd-bridge"),
    )
    for p in cands:
        if not p:
            continue
        if os.path.isfile(os.path.join(p, "dsd_bridge.py")):
            return p, "dsd_bridge"
        if os.path.isfile(os.path.join(p, "dsdx.py")):   # 이름 변경(2026-11-24) 전 복제본
            return p, "dsdx"
    return None, None


def xsd(raw):
    home, pkg = _bridge_home()
    if not home:
        return "XSD 검증 생략(DSD브릿지 폴더 못 찾음, DSD_BRIDGE_HOME 지정)"
    sys.path.insert(0, home)
    if pkg == "dsd_bridge":
        from dsd_bridge.validate import xsd_validate
    else:
        from dsdx.validate import xsd_validate
    return xsd_validate(raw, "dart4")


def load(path):
    z = zipfile.ZipFile(path)
    return z, z.read("contents.xml").decode("utf-8")


def save(z, s, out, src):
    assert os.path.abspath(out) != os.path.abspath(src), "입력 파일 덮어쓰기 금지: 새 이름으로 -o 지정"
    raw = s.encode("utf-8")
    print("XSD:", xsd(raw))
    with zipfile.ZipFile(out, "w") as zo:
        for info in z.infolist():
            zo.writestr(info, raw if info.filename == "contents.xml" else z.read(info.filename), compress_type=info.compress_type)
    print("저장:", out)


TAG = re.compile(r"<[^>]+>")
# 표 칸(TE·TD·TU) 안의 숫자만. EXTRACTION(요약값: 회사코드·자산총계 등)·날짜 속성은 콤마를 넣으면 안 된다
NUM_NOCOMMA = re.compile(r"(<(?:TE|TD|TU)(?=[\s>])[^>]*>\s*)(\(?-?)(\d{4,})(\)?)(\s*</(?:TE|TD|TU)>)")


def _is_date(t):
    return len(t) == 8 and t.startswith(("19", "20"))


def info(path):
    _, s = load(path)
    print("[기간 속성]")
    for m in re.finditer(r'AUNIT="((?:SUB_)?PERIOD(?:FROM|TO)2?)" AUNITVALUE="(\d*)"', s):
        print(f"  {m.group(1)} = {m.group(2)}")
    print("[COMM_DATE 등 날짜 속성]", re.findall(r'AUNIT="(COMM_DATE)" AUNITVALUE="([^"]*)"', s)[:10])
    txt = TAG.sub("", s)
    print("[자리표시 ○]", len(re.findall("○", txt)), "곳")
    bad = [m.group(3) for m in NUM_NOCOMMA.finditer(s) if not _is_date(m.group(3))]
    print("[콤마 누락 의심]", bad[:30])
    print("[같은 ACODE 가 연속 행에 반복]")
    rows = re.findall(r"<TR\b[^>]*>(.*?)</TR>", s, re.S)
    prev = None
    for r in rows:
        codes = set(re.findall(r'ACODE="(\d+)"', r))
        label = TAG.sub("", (re.findall(r"<TE\b[^>]*>(.*?)</TE>", r, re.S) or [""])[0]).strip()
        if len(codes) == 1:
            c = codes.pop()
            if prev and prev[0] == c and c != "99999999999999":  # 신설 코드는 여러 행이 같이 씀
                print(f"  {c}: '{prev[1]}' / '{label}'  <- 복제 삽입 후 코드 미변경 의심")
            prev = (c, label)
        else:
            prev = None


def period(path, out, frm, to):
    z, s = load(path)
    n = 0
    def rep(m):
        nonlocal n
        v = frm if "FROM" in m.group(1) else to
        if m.group(2) != v:
            print(f"  {m.group(1)}: {m.group(2)} -> {v}"); n += 1
        return f'AUNIT="{m.group(1)}" AUNITVALUE="{v}"'
    s = re.sub(r'AUNIT="((?:SUB_)?PERIOD(?:FROM|TO)2?)" AUNITVALUE="(\d*)"', rep, s)
    print("바꾼 곳:", n)
    save(z, s, out, path)


def acode(path, out, label, code, nth=1):
    z, s = load(path)
    hits = [m for m in re.finditer(r"<TR\b[^>]*>.*?</TR>", s, re.S)
            if label in TAG.sub("", (re.findall(r"<TE\b[^>]*>(.*?)</TE>", m.group(0), re.S) or [""])[0])]
    print(f"'{label}' 행 {len(hits)}개")
    assert len(hits) >= nth, "행을 못 찾음"
    m = hits[nth - 1]
    new = re.sub(r'ACODE="\d*"', f'ACODE="{code}"', m.group(0))
    print("  이전 코드:", sorted(set(re.findall(r'ACODE="(\d*)"', m.group(0)))), "->", code)
    s = s[:m.start()] + new + s[m.end():]
    save(z, s, out, path)


def replace(path, out, old, new, count=None):
    z, s = load(path)
    n = s.count(old)
    print(f"'{old[:40]}' {n}곳")
    assert n > 0 and (count is None or n == count), "개수 불일치: 글자가 태그로 끊겨 있으면 더 짧게 지정"
    save(z, s.replace(old, new), out, path)


def commas(path, out):
    z, s = load(path)
    n = 0
    def rep(m):
        nonlocal n
        n += 1
        return f"{m.group(1)}{m.group(2)}{int(m.group(3)):,}{m.group(4)}{m.group(5)}"
    s2 = NUM_NOCOMMA.sub(lambda m: m.group(0) if _is_date(m.group(3)) else rep(m), s)
    print("콤마 넣은 칸:", n)
    save(z, s2, out, path)


if __name__ == "__main__":
    a = sys.argv[1:]
    cmd, src = a[0], a[1]
    if cmd == "info":
        info(src); sys.exit()
    assert a[2] == "-o", "사용법은 파일 머리말 참고"
    out, rest = a[3], a[4:]
    if cmd == "period":
        period(src, out, rest[0], rest[1])
    elif cmd == "acode":
        acode(src, out, rest[0], rest[1], int(rest[2]) if len(rest) > 2 else 1)
    elif cmd == "replace":
        replace(src, out, rest[0], rest[1], int(rest[2]) if len(rest) > 2 else None)
    elif cmd == "commas":
        commas(src, out)
