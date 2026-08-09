import csv
import sys

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

CSV_PATH = "data/basis/토지특성정보경기csv/AL_D195_41_20260402.csv"
ENCODING = "cp949"
READ_COUNT = 3

with open(CSV_PATH, encoding=ENCODING, newline="") as f:
    reader = csv.reader(f)
    fields = next(reader)
    print("필드 목록:", fields)
    print("-" * 60)
    for i, row in enumerate(reader):
        if i >= READ_COUNT:
            break
        record = dict(zip(fields, row))
        print(f"[{i + 1}]")
        for k, v in record.items():
            print(f"    {k}: {v}")
