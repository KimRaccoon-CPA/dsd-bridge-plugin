# -*- coding: utf-8 -*-
"""DSD브릿지(dsd-bridge) export 작업 엑셀의 기간 표기(제N기, 시작·종료일)를 한 번에 갱신한다(2단계 이월용, openpyxl).
옛 당기 표기 -> 새 당기, 옛 전기 표기 -> 새 전기를 '한 번의 치환'으로 바꿔서 연쇄 치환(옛 당기말 = 새 전기말 날짜가 두 번 바뀌는 사고)을 막는다.

기본은 미리보기(바뀔 칸 목록만 출력). 목록을 보고 주석 본문 속 거래일처럼 바꾸면 안 되는 칸이 없을 때 --apply.
  python dsd_period.py 작업.xlsx --n 9 10 --old 20240101 20241231 20230101 20231231 --new 20250101 20251231 20240101 20241231
         [--apply] [--only "시트!F5,시트!G36"] [--skip "주석!F120"]
  --n 옛당기기수 새당기기수, --old 옛당기시작 옛당기종료 옛전기시작 옛전기종료, --new 새당기시작 새당기종료 새전기시작 새전기종료
표기 형식: 'YYYY년 MM월 DD일', 'YYYY년 M월 D일', 'YYYY.MM.DD', 'YYYY. MM. DD', 'YYYY-MM-DD', 'YYYY/MM/DD',
          '제 N (당) 기', '제 N(당) 기', '제 N(당)기', '제N(당)기', '제 N 기'(당기만), 전기도 같은 형식.
회색 칸(서식 조각이 섞인 칸)도 바꿔도 된다. 0.3.0 부터 되돌릴 때 서식을 두고 글자만 반영하고, 표지·외부감사 실시내용의
날짜 칸 글이 바뀌면 숨은 날짜 속성값(AUNITVALUE)도 같이 맞춘다.
남는 것(사람 확인): 감사의견 문단의 'A와 B' 순서, 자본변동표 '당기초/전기말' 라벨, 감사보고서일(○ 자리표시), 외부감사 실시내용의 감사대상 사업연도.
"""
import argparse, datetime, re, sys

import openpyxl


def fmts(ymd):
    d = datetime.datetime.strptime(ymd, "%Y%m%d").date()
    y, m, dd = d.year, d.month, d.day
    return [f"{y}년 {m:02d}월 {dd:02d}일", f"{y}년 {m}월 {dd}일", f"{y}.{m:02d}.{dd:02d}", f"{y}. {m:02d}. {dd:02d}",
            f"{y}-{m:02d}-{dd:02d}", f"{y}/{m:02d}/{dd:02d}"]


def build_map(n_old, n_new, old, new):
    rep = {}
    for o, nw in zip(old, new):
        for a, b in zip(fmts(o), fmts(nw)):
            rep[a] = b
    po, pn = n_old - 1, n_new - 1
    for tag, a, b in (("당", n_old, n_new), ("전", po, pn)):
        for pat in ("제 {} ({}) 기", "제 {}({}) 기", "제 {}({})기", "제{}({})기"):
            rep[pat.format(a, tag)] = pat.format(b, tag)
    rep[f"제 {n_old} 기"] = f"제 {n_new} 기"
    return rep


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--n", nargs=2, type=int, required=True)
    ap.add_argument("--old", nargs=4, required=True)
    ap.add_argument("--new", nargs=4, required=True)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--only", default="")
    ap.add_argument("--skip", default="")
    a = ap.parse_args()
    rep = build_map(a.n[0], a.n[1], a.old, a.new)
    pat = re.compile("|".join(re.escape(k) for k in sorted(rep, key=len, reverse=True)))
    only = {x.strip() for x in a.only.split(",") if x.strip()}
    skip = {x.strip() for x in a.skip.split(",") if x.strip()}
    wb = openpyxl.load_workbook(a.xlsx)
    hits = []
    for ws in wb.worksheets:
        if ws.title.startswith("_dsdx"):
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.column <= 5 or not isinstance(c.value, str):   # A~E 열은 DSD브릿지 숨김 열
                    continue
                key = f"{ws.title}!{c.coordinate}"
                if (only and key not in only) or key in skip:
                    continue
                nv = pat.sub(lambda m: rep[m.group(0)], c.value)
                if nv != c.value:
                    hits.append((key, c.value, nv, c))
    for key, ov, nv, _ in hits:
        print(f"{key}\n   전: {ov[:150]}\n   후: {nv[:150]}")
    print(f"\n바뀔 칸 {len(hits)}개" + ("" if a.apply else " (미리보기, 반영하려면 --apply)"))
    if a.apply:
        for _, _, nv, c in hits:
            c.value = nv
        wb.save(a.xlsx)
        print("저장:", a.xlsx)


if __name__ == "__main__":
    main()
