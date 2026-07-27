"""
Xac dinh Vung cua 1 toa do GPS bang cach so voi danh sach polygon nguoi dung
tu ve tren ban do va luu o Google Sheets (xem sheets_client.py, zone_editor/app.py).
Thay the cho cot "Vung" co san cua Adsun de tranh nhuoc diem: Adsun chi nhan
dien vung tu thoi diem vung do duoc tao tren he thong, khong ap dung hoi to
duoc cho du lieu qua khu. O day tinh lai tu toa do tho moi lan chay bao cao,
nen vung moi tao van ap dung dung cho ca ngay hom truoc.
"""

from shapely.geometry import Point, Polygon

import sheets_client


def load_zone_polygons():
    """Tra ve list (ten_vung, shapely_polygon) theo dung thu tu trong Sheets —
    vung dung truoc trong danh sach duoc uu tien khi toa do nam trong nhieu
    vung chong lan nhau."""
    zones = sheets_client.list_zones()
    polygons = []
    for z in zones:
        try:
            # points luu dang [lat, lon]; shapely can thu tu (x, y) = (lon, lat)
            coords = [(p[1], p[0]) for p in z["points"]]
            polygons.append((z["name"], Polygon(coords)))
        except Exception:
            continue
    return polygons


def find_zone(lat, lon, polygons):
    point = Point(lon, lat)
    for name, poly in polygons:
        if poly.covers(point):
            return name
    return None


def parse_coordinate(text):
    """'10.8159733, 106.755318' -> (10.8159733, 106.755318)"""
    if text is None:
        return None
    try:
        lat_str, lon_str = str(text).split(",")
        return float(lat_str.strip()), float(lon_str.strip())
    except Exception:
        return None


def compute_zone_column(df, polygons, coord_column="Tọa độ", zone_column="Vùng"):
    """Tra ve DataFrame moi voi cot Vung duoc GHI DE bang ket qua tinh tu toa do."""
    def _lookup(coord_text):
        parsed = parse_coordinate(coord_text)
        if parsed is None:
            return None
        lat, lon = parsed
        return find_zone(lat, lon, polygons)

    df = df.copy()
    df[zone_column] = df[coord_column].apply(_lookup)
    return df
