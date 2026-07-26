import sys
from pathlib import Path

import openpyxl
import shapefile

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

SHP_PATH = "shp/AL_D152_11_20260520.shp"
COLUMN_DEF_PATH = "shp/국가중점데이터_컬럼정의서(26.01.02)_배포용.xlsx"
READ_COUNT = 3

# 파일명 앞부분(예: AL_D152)을 컬럼정의서 조회 키로 사용
FILE_CODE = Path(SHP_PATH).stem.rsplit("_", 2)[0]


def load_column_names(xlsx_path: str, file_code: str) -> dict:
    """컬럼정의서(전체) 시트에서 file_code에 해당하는 'A0'~'A28' -> 한글 항목명 매핑을 만든다."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["테이블정의서(전체)"]

    names = {}
    for row in ws.iter_rows(min_row=7, values_only=True):
        file_name, field_code, field_name_kr = row[3], row[6], row[7]
        if file_name == file_code and field_code:
            names[field_code] = field_name_kr
    return names


field_names_kr = load_column_names(COLUMN_DEF_PATH, FILE_CODE)

sf = shapefile.Reader(SHP_PATH, encoding="cp949")

fields = [f[0] for f in sf.fields[1:]]  # 첫 필드(DeletionFlag) 제외
print("필드 목록:")
for f in fields:
    print(f"  {f}: {field_names_kr.get(f, '(정의 없음)')}")
print(f"전체 레코드 수: {len(sf)}")
print("-" * 60)

for i, sr in enumerate(sf.iterShapeRecords()):
    if i >= READ_COUNT:
        break
    record = sr.record.as_dict()
    shape = sr.shape
    print(f"[{i + 1}] shapeType={shape.shapeType}, points={len(shape.points)}")
    for code, value in record.items():
        label = field_names_kr.get(code, code)
        print(f"    {code}({label}): {value}")

sf.close()
