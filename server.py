from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Dict
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import json
import uvicorn
import pandas as pd
import math

app = FastAPI(title="스마트 버스 서버")

# -----------------------------------------------------
# 1. CORS 설정 (클라이언트 앱과 통신 허용)
# -----------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------------------------------------
# 2. 엑셀/CSV 데이터 로드 (서버 구동 시 1회 로드)
# -----------------------------------------------------
FILE_NAME = "서울시버스노선별정류소정보.xlsx" 
try:
    if FILE_NAME.endswith(".csv"):
        df_stops = pd.read_csv(FILE_NAME)
    else:
        df_stops = pd.read_excel(FILE_NAME)
    print("✅ 버스 노선 데이터 로드 성공!")
except Exception as e:
    print("❌ 엑셀/CSV 데이터 로드 실패:", e)
    df_stops = pd.DataFrame()

# -----------------------------------------------------
# 3. 하버사인(Haversine) 거리 계산 함수
# -----------------------------------------------------
def haversine(lat1, lon1, lat2, lon2):
    """두 위경도 좌표 사이의 거리를 km 단위로 반환합니다."""
    R = 6371.0 # 지구 반지름 (km)
    dLat = math.radians(lat2 - lat1)
    dLon = math.radians(lon2 - lon1)
    a = math.sin(dLat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dLon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

# -----------------------------------------------------
# 4. 버스 대기열 데이터 저장소 및 구조
# -----------------------------------------------------
waiting_data: Dict[str, Dict[str, int]] = {}

class RideRequest(BaseModel):
    route: str
    stop: str

# -----------------------------------------------------
# 5. API 라우터 (엔드포인트) 설정
# -----------------------------------------------------
@app.post("/request_ride")
def request_ride(req: RideRequest):
    """승객 승차 요청 처리"""
    if req.route not in waiting_data:
        waiting_data[req.route] = {}
    if req.stop not in waiting_data[req.route]:
        waiting_data[req.route][req.stop] = 0
        
    waiting_data[req.route][req.stop] += 1
    return {"status": "success", "message": "예약 완료"}

@app.get("/check_stop/{route}/{stop}")
def check_stop(route: str, stop: str):
    """기사용 대기 승객 수 조회"""
    count = waiting_data.get(route, {}).get(stop, 0)
    return {"count": count}

@app.post("/clear_stop")
def clear_stop(req: RideRequest):
    """버스 정차 후 대기열 초기화"""
    if req.route in waiting_data and req.stop in waiting_data[req.route]:
        waiting_data[req.route][req.stop] = 0
    return {"status": "success", "message": "초기화 완료"}

# 🆕 특정 노선의 가까운 정류장 정렬 API (엑셀 구조 반영)
@app.get("/nearest_stops/{route_name}")
def get_nearest_stops(route_name: str, my_lat: float, my_lng: float):
    if df_stops.empty:
        return {"status": "error", "message": "서버에 노선 데이터가 없습니다."}
        
    # 캡처해주신 엑셀 구조에 맞춘 인덱스(열 위치) 매핑
    col_route_id   = df_stops.columns[0] # A열: 노선ID
    col_route_name = df_stops.columns[1] # B열: 노선명
    col_ord        = df_stops.columns[2] # C열: 순번
    col_node_id    = df_stops.columns[3] # D열: 노드ID (표준 정류장 ID)
    col_stop_name  = df_stops.columns[5] # F열: 정류소명
    col_x          = df_stops.columns[6] # G열: X좌표 (경도)
    col_y          = df_stops.columns[7] # H열: Y좌표 (위도)

    # 검색할 노선명 전처리 (예: "740번" -> "740")
    search_route = route_name.replace("번", "").strip()
    
    # B열(노선명)을 기준으로 해당 노선의 정류장만 필터링
    route_df = df_stops[df_stops[col_route_name].astype(str) == search_route].copy()
    
    if route_df.empty:
        return {"status": "error", "message": "해당 노선 정보를 찾을 수 없습니다."}
        
    # 각 정류장과 내 위치 사이의 거리 계산 (H열 Y좌표, G열 X좌표 사용)
    route_df['거리(km)'] = route_df.apply(
        lambda row: haversine(my_lat, my_lng, float(row[col_y]), float(row[col_x])), axis=1
    )
    
    # 거리가 가까운 순으로 정렬
    route_df = route_df.sort_values('거리(km)')
    
    # 노선 고유 ID 추출 (A열)
    route_id = str(route_df.iloc[0][col_route_id])
    
    # 앱으로 보낼 정류장 리스트 생성
    stops = []
    for _, row in route_df.iterrows():
        stops.append({
            "stop_name": str(row[col_stop_name]),
            "stop_id": str(row[col_node_id]), # 노드ID를 정류장 고유 ID로 사용
            "ord": str(row[col_ord]),
            "distance_km": round(row['거리(km)'], 2)
        })
        
    return {"status": "success", "route_id": route_id, "stops": stops}

# 공공데이터포털 일반 인증키(Decoding)
DATA_GO_KR_API_KEY = "473669718a0ff6107ecc58dd12eacee5e3a9ac9f3c6e28f34ecb9b96d5023db5"

@app.get("/real_bus_location/{route_id}/{stop_id}/{ord}")
def get_real_bus_location(route_id: str, stop_id: str, ord: str):
    """
    서울시 버스위치정보조회 API(getBusPosByRtidList)를 호출하여
    노선의 모든 버스 위치를 가져온 뒤, 승객과 가장 가까운 버스를 찾습니다.
    """
    url = "http://ws.bus.go.kr/api/rest/buspos/getBusPosByRtid"
    
    # urllib.parse가 키를 건드리지 못하게 ServiceKey는 직접 문자열로 결합
    params = urllib.parse.urlencode({
        'busRouteId': route_id,
        'resultType': 'json'
    })
    full_url = f"{url}?ServiceKey={DATA_GO_KR_API_KEY}&{params}"

    try:
        request = urllib.request.Request(full_url)
        request.get_method = lambda: 'GET'
        response = urllib.request.urlopen(request)
        res_code = response.getcode()
        
        if res_code == 200:
            response_body = response.read()
            data = json.loads(response_body.decode('utf-8'))
            
            if data['msgHeader']['headerCd'] == '0':
                item_list = data['msgBody'].get('itemList', [])
                
                if not item_list:
                    return {"status": "empty", "message": "현재 운행 중인 버스가 없습니다."}

                target_ord = int(ord) 
                approaching_buses = []
                
                for bus in item_list:
                    bus_ord = int(bus['sectOrd'])
                    if bus_ord <= target_ord:
                        approaching_buses.append(bus)
                
                if approaching_buses:
                    closest_bus = min(approaching_buses, key=lambda x: target_ord - int(x['sectOrd']))
                    stops_left = target_ord - int(closest_bus['sectOrd'])
                    plain_no = closest_bus['plainNo'] 
                    
                    if stops_left == 0:
                        arrmsg = "현재 정류장 구간에 진입했습니다."
                    else:
                        arrmsg = f"{stops_left}번째 전 정류장 통과 중"
                        
                    return {"status": "success", "arrmsg1": arrmsg, "vehId1": plain_no}
                else:
                    return {"status": "empty", "message": "이 정류장으로 접근 중인 버스가 없습니다."}
            else:
                return {"status": "error", "message": data['msgHeader']['headerMsg']}
        else:
            return {"status": "error", "message": f"HTTP Error {res_code}"}
            
    except Exception as e:
        return {"status": "error", "message": f"서버 통신 에러: {str(e)}"}
    
# -----------------------------------------------------
# 6. 서버 실행 (python3 server.py 로 실행 시)
# -----------------------------------------------------
if __name__ == "__main__":
    uvicorn.run("server:app", host="0.0.0.0", port=8000, reload=True)