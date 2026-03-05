import os
import re
import pandas as pd
import requests
import folium
from folium.plugins import HeatMap, MarkerCluster
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
from dotenv import load_dotenv

# .env 파일 로드 (KAKAO_REST_API_KEY, DEFAULT_CENTER_LAT, DEFAULT_CENTER_LNG 포함)
load_dotenv()

# ─────────────────────────────────────────────
# ① 설정 및 보안
# ─────────────────────────────────────────────
KAKAO_API_KEY = os.getenv("KAKAO_REST_API_KEY")
CENTER_LAT = float(os.getenv("DEFAULT_CENTER_LAT", 37.5665))
CENTER_LNG = float(os.getenv("DEFAULT_CENTER_LNG", 126.9780))

# 지오코딩 결과 캐시 (동일 주소 중복 호출 방지)
address_cache = {}

# ─────────────────────────────────────────────
# ② 데이터 전처리 로직
# ─────────────────────────────────────────────

def clean_address(addr):
    """지오코딩 성공률을 높이기 위한 주소 정제"""
    if not addr or pd.isna(addr): return ""
    addr = str(addr)
    # 괄호 안 내용 삭제 (예: 아파트명, 빌라명 등 상세정보)
    addr = re.sub(r'\(.*\)', '', addr)
    addr = re.sub(r'\[.*\]', '', addr)
    # 불필요한 공백 및 상세 층/호수 패턴 제거
    patterns = [r'\s+\d+동\s+\d+호', r'\s+\d+층', r'\s+\d+호']
    for p in patterns:
        addr = re.sub(p, '', addr)
    return addr.strip()

def get_coords(address):
    """카카오 API 호출 (캐싱 적용)"""
    cleaned = clean_address(address)
    if not cleaned:
        return None, None, "EMPTY"
    
    # 1. 이미 처리한 주소인지 확인 (중복 제거로 API 쿼터 절약)
    if cleaned in address_cache:
        return address_cache[cleaned]

    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    
    try:
        res = requests.get(url, headers=headers, params={"query": cleaned}, timeout=5)
        if res.status_code == 200:
            docs = res.json().get("documents")
            if docs:
                lat, lng = float(docs[0]["y"]), float(docs[0]["x"])
                address_cache[cleaned] = (lat, lng, "SUCCESS")
                return lat, lng, "SUCCESS"
    except Exception as e:
        pass
    
    return None, None, "FAIL"

# ─────────────────────────────────────────────
# ③ 시각화 엔진 (2만 건 최적화 버전)
# ─────────────────────────────────────────────

def create_optimized_map(df, output_name="analysis_result.html"):
    print("🎨 지도 생성 중... (데이터가 많으면 다소 시간이 걸릴 수 있습니다)")
    
    # 지도 초기화 (배경은 깔끔한 CartoDB Positron)
    m = folium.Map(location=[CENTER_LAT, CENTER_LNG], zoom_start=13, tiles="CartoDB positron")
    
    # 좌표가 있는 데이터만 필터링
    valid_df = df[df['status'] == "SUCCESS"].copy()
    
    # 1. 히트맵 레이어 (전체적인 밀집도 파악)
    heat_data = valid_df[['lat', 'lng']].values.tolist()
    HeatMap(heat_data, name="🔥 밀집도 히트맵", radius=12, blur=8).add_to(m)
    
    # 2. 마커 클러스터 레이어 (2만 건의 점을 성능 저하 없이 표시)
    # 개별 Marker가 아닌 MarkerCluster를 써야 브라우저가 안 멈춥니다.
    marker_cluster = MarkerCluster(name="📍 개별 위치 (클러스터)").add_to(m)
    for _, row in valid_df.iterrows():
        folium.CircleMarker(
            location=[row['lat'], row['lng']],
            radius=4,
            color='#3186cc',
            fill=True,
            fill_color='#3186cc',
            popup=row['주소']
        ).add_to(marker_cluster)
    
    # 3. 분석 반경 표시 (250m, 500m, 1km)
    zones = folium.FeatureGroup(name="📏 거리 반경", show=True)
    radii = [(250, "#E63946", "250m"), (500, "#F4A261", "500m"), (1000, "#2A9D8F", "1km")]
    for r, color, label in radii:
        folium.Circle(
            location=[CENTER_LAT, CENTER_LNG],
            radius=r,
            color=color,
            fill=True,
            fill_opacity=0.07,
            popup=label,
            dash_array='5, 5'
        ).add_to(zones)
    zones.add_to(m)

    # 중심점 마커
    folium.Marker(
        [CENTER_LAT, CENTER_LNG], 
        popup="분석 중심지", 
        icon=folium.Icon(color='red', icon='star')
    ).add_to(m)

    folium.LayerControl().add_to(m)
    m.save(output_name)
    print(f"✅ 시각화 완료: {output_name}")

# ─────────────────────────────────────────────
# ④ 실행 메인 프로세스
# ─────────────────────────────────────────────

def main():
    input_file = "raw_addresses.csv"  # 입력 파일명
    if not os.path.exists(input_file):
        print(f"❌ '{input_file}' 파일이 없습니다.")
        return

    # 데이터 로드 (주소 컬럼명 확인 필요)
    df = pd.read_csv(input_file)
    if '주소' not in df.columns:
        print("❌ CSV 파일에 '주소' 컬럼이 필요합니다.")
        return

    print(f"🚀 총 {len(df):,}건 지오코딩 시작 (병렬 처리)...")

    # 병렬 처리 (Thread 15개 정도가 적당함)
    with ThreadPoolExecutor(max_workers=15) as executor:
        # tqdm으로 진행률 표시
        results = list(tqdm(executor.map(get_coords, df['주소']), total=len(df)))
    
    # 결과 병합
    df['lat'], df['lng'], df['status'] = zip(*results)

    # 결과물 중간 저장 (지오코딩된 데이터 보존)
    output_csv = "geocoded_results.csv"
    df.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"💾 좌표 변환 완료 및 저장: {output_csv}")

    # 실패 데이터 통계
    fail_count = len(df[df['status'] == "FAIL"])
    print(f"📊 통계: 성공 {len(df)-fail_count}건 / 실패 {fail_count}건")

    # 지도 생성
    create_optimized_map(df)

if __name__ == "__main__":
    main()
