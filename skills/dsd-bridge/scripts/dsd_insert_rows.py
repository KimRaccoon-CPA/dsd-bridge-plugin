# -*- coding: utf-8 -*-
"""DSD브릿지 작업 엑셀의 표 안에 빈 행 끼우기(Excel COM). 병합 칸이 있는 표는 openpyxl insert_rows 로 하면 병합·DSD브릿지 행 정보가 깨진다.
새 행은 바로 위 행의 서식을 따른다. 끼운 뒤 라벨·금액을 채우고, import 후 위 행 ACODE 를 물려받았으면 dsd_xml.py acode 로 고친다.
  python dsd_insert_rows.py 작업.xlsx "시트이름" 행번호 [개수] [--expect 위행F열글자]
예) 자본변동표 전기 구간 '당기순이익'(F107) 아래에 한 줄: dsd_insert_rows.py 제10기_작업.xlsx "(첨부)재 무 제 표" 108 1 --expect 당기순이익
"""
import os, sys

import win32com.client as win32

a = sys.argv[1:]
path, sheet, row = os.path.normpath(os.path.abspath(a[0])), a[1], int(a[2])
count = int(a[3]) if len(a) > 3 and not a[3].startswith("--") else 1
expect = a[a.index("--expect") + 1] if "--expect" in a else None
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False; xl.AskToUpdateLinks = False
try:
    wb = xl.Workbooks.Open(path, 0)
    try:
        wb.AutoSaveOn = False
    except Exception:
        pass
    ws = wb.Worksheets(sheet)
    above = ws.Cells(row - 1, 6).Value
    if expect is not None:
        assert expect in str(above), f"위 행(F{row - 1}) 글자가 예상과 다름: {above!r}"
    for _ in range(count):
        ws.Rows(row).Insert(-4121, 0)   # 아래로 밀기, 위 행 서식
    print(f"{sheet} {row}행에 {count}줄 삽입. 위 행 F{row - 1}: {above!r} / 아래 행 F{row + count}: {ws.Cells(row + count, 6).Value!r}")
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
