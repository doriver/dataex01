# geopandas를 활용해 SHP 데이터를 좌표와 함께 추출하는 예시 (Python)
import sys

import geopandas as gpd

if sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

SHP_PATH = "shp/AL_D152_11_20260520.shp"
READ_COUNT = 3

# SHP 파일 읽기
gdf = gpd.read_file(SHP_PATH, encoding="cp949")

print(f"좌표계(CRS): {gdf.crs}")
print(f"전체 레코드 수: {len(gdf)}")
print("-" * 60)

for i, row in gdf.head(READ_COUNT).iterrows():
    geom = row.geometry
    centroid = geom.centroid
    print(f"[{i + 1}] geom_type={geom.geom_type}")
    print(f"    중심좌표: ({centroid.x:.3f}, {centroid.y:.3f})")
    print(f"    좌표(coords): {list(geom.exterior.coords)}")
    print(f"    속성: {row.drop('geometry').to_dict()}")
    print()
