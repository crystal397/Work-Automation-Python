import requests, os
from dotenv import load_dotenv
load_dotenv()
KEY = os.getenv("VWORLD_API_KEY")
BASE_URL = "http://api.vworld.kr/ned/data/getIndvdLandPriceAttr"

# 1. 팔달로3가 133 → 팔달로2가 코드로 최신 데이터 확인
pnu = "4111512100101330000"
resp = requests.get(BASE_URL, params={
    "key": KEY, "pnu": pnu, "format": "json", "numOfRows": "100", "pageNo": "1"
}, timeout=10).json()
if "indvdLandPrices" in resp:
    fields = resp["indvdLandPrices"].get("field", [])
    if isinstance(fields, dict): fields = [fields]
    fields.sort(key=lambda x: x.get("stdrYear",""), reverse=True)
    print(f"팔달로3가 133 최신: {fields[0].get('stdrYear')}년 {fields[0].get('pblntfPclnd')}원/㎡")

# 2. 남창동 138-1 좌표로 PNU 찾기
print("\n남창동 138-1 주변 필지 검색:")
resp2 = requests.get("https://api.vworld.kr/req/address", params={
    "service": "address", "request": "getcoord", "crs": "epsg:4326",
    "address": "경기도 수원시 팔달구 남창동 138-1",
    "format": "json", "type": "PARCEL", "key": KEY,
}, timeout=10).json()
x = resp2["response"]["result"]["point"]["x"]
y = resp2["response"]["result"]["point"]["y"]
print(f"좌표: {x}, {y}")

delta = 0.0002
bbox = f"{float(x)-delta},{float(y)-delta},{float(x)+delta},{float(y)+delta}"
resp3 = requests.get("https://api.vworld.kr/req/data", params={
    "service": "data", "request": "GetFeature", "data": "LP_PA_CBND_BUBUN",
    "key": KEY, "format": "json", "geometry": "false",
    "geomFilter": f"BOX({bbox})", "size": "10",
}, timeout=10).json()
features = resp3["response"]["result"]["featureCollection"]["features"]
for f in features:
    props = f["properties"]
    pnu2 = props.get("pnu")
    addr = props.get("addr")
    # 공시지가 조회
    r = requests.get(BASE_URL, params={
        "key": KEY, "pnu": pnu2, "format": "json", "numOfRows": "1", "pageNo": "1"
    }, timeout=10).json()
    if "indvdLandPrices" in r:
        flds = r["indvdLandPrices"].get("field", [])
        if isinstance(flds, dict): flds = [flds]
        flds.sort(key=lambda x: x.get("stdrYear",""), reverse=True)
        price_info = f"{flds[0].get('stdrYear')}년 {flds[0].get('pblntfPclnd')}원" if flds else "없음"
    else:
        price_info = "없음"
    print(f"  PNU:{pnu2}  주소:{addr}  공시지가:{price_info}")