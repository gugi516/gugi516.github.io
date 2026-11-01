# -*- coding: utf-8 -*-
"""application_info.xlsx -> _csv/<시트명>.csv 미러 생성.

xlsx 는 바이너리라 git diff 로 무엇이 바뀌었는지 보이지 않는다.
이 스크립트가 시트별 CSV 를 떨어뜨려서 변경 내역을 사람이 읽을 수 있게 한다.
CSV 는 읽기 전용 미러다. 편집은 항상 xlsx 에서 한다.

  실행:  python _export_csv.py        (xlsx 를 고칠 때마다)
"""
import csv
import datetime as dt
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "application_info.xlsx")
OUT = os.path.join(HERE, "_csv")


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour or v.minute) == 0 else v.isoformat(" ")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    return str(v).replace("\r\n", "\n").strip()


def main():
    os.makedirs(OUT, exist_ok=True)
    wb = openpyxl.load_workbook(SRC, data_only=True)

    written = set()
    for ws in wb.worksheets:
        name = ws.title.replace("/", "_").replace("&", "and").strip()
        path = os.path.join(OUT, f"{name}.csv")
        rows = [[fmt(c) for c in r] for r in ws.iter_rows(values_only=True)]
        while rows and not any(rows[-1]):          # 꼬리 빈 행 제거
            rows.pop()
        # newline="" + LF 고정: 플랫폼 무관하게 동일한 diff
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, lineterminator="\n").writerows(rows)
        written.add(f"{name}.csv")
        print(f"  {ws.title:24s} -> {name}.csv  ({len(rows) - 1}건)")

    # xlsx 에서 사라진 시트의 CSV 는 남겨두지 않는다
    for stale in sorted(set(os.listdir(OUT)) - written):
        if stale.endswith(".csv"):
            os.remove(os.path.join(OUT, stale))
            print(f"  삭제(원본에 없는 시트): {stale}")

    print(f"\n{len(written)}개 시트 -> {OUT}")


if __name__ == "__main__":
    main()
