import tkinter as tk
from tkinter import ttk, messagebox
import requests
import pandas as pd
import os
import urllib.parse
import json
import tkintermapview 
import math 
from PIL import Image, ImageTk 

SERVER_URL = "http://127.0.0.1:8000"
DATA_GO_KR_API_KEY = "473669718a0ff6107ecc58dd12eacee5e3a9ac9f3c6e28f34ecb9b96d5023db5"

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dLat = math.radians(lat2 - lat1)
    dLon = math.radians(lon2 - lon1)
    a = math.sin(dLat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dLon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

FILE_NAME_STOPS = "서울시버스노선별정류소정보(20260506).xlsx"
FILE_NAME_ROUTES = "서울시버스노선ID정보.xlsx"

DF_STOPS = pd.DataFrame()
DF_ROUTES = pd.DataFrame()

try:
    if os.path.exists(FILE_NAME_STOPS):
        if FILE_NAME_STOPS.endswith(".csv"):
            try:
                DF_STOPS = pd.read_csv(FILE_NAME_STOPS, encoding='cp949')
            except:
                DF_STOPS = pd.read_csv(FILE_NAME_STOPS, encoding='utf-8')
        else:
            DF_STOPS = pd.read_excel(FILE_NAME_STOPS)
except Exception as e:
    pass

try:
    if os.path.exists(FILE_NAME_ROUTES):
        if FILE_NAME_ROUTES.endswith(".csv"):
            try:
                DF_ROUTES = pd.read_csv(FILE_NAME_ROUTES, encoding='cp949')
            except:
                DF_ROUTES = pd.read_csv(FILE_NAME_ROUTES, encoding='utf-8')
        else:
            DF_ROUTES = pd.read_excel(FILE_NAME_ROUTES)
        print("✅ 노선 ID 매핑 데이터 로드 성공!")
    else:
        print(f"⚠️ '{FILE_NAME_ROUTES}' 파일이 없습니다. 기존 방식으로 ID를 추출합니다.")
except Exception as e:
    print("❌ 노선 ID 매핑 데이터 로드 실패:", e)

class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(list(completion_list))
        self['values'] = tuple(self._completion_list)
        self.bind('<KeyRelease>', self.handle_keyrelease)
        
    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down', 'Return', 'Shift_L', 'Shift_R', 'Tab'):
            return
        typed_text = self.get()
        if typed_text == '':
            self['values'] = tuple(self._completion_list)
        else:
            filtered = [item for item in self._completion_list if typed_text.lower() in item.lower()]
            self['values'] = tuple(filtered)

class PassengerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("스마트 버스 - 승객용 앱")
        self.root.geometry("480x800") 
        self.root.configure(bg="#F4F6F9") 
        
        self.my_lat, self.my_lng = 37.4979, 126.9108 
        self.current_route_id = None
        self.stop_data_map = {}
        self.all_route_list = []

        self.setup_ui()
        self.load_routes_from_data()

    def create_flat_button(self, parent, text, bg_color, hover_color, command, font_style, pady_val):
        btn = tk.Label(parent, text=text, bg=bg_color, fg="white", font=font_style, pady=pady_val, cursor="hand2")
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_color))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg_color))
        btn.bind("<Button-1>", lambda e: command())
        return btn

    def setup_ui(self):
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        style.configure('TCombobox', padding=6, arrowsize=14)

        header_frame = tk.Frame(self.root, bg="#2563EB", bd=0)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="스마트 버스 승차 (승객용)", bg="#2563EB", fg="white", font=("맑은 고딕", 16, "bold"), pady=15).pack()

        loc_frame = tk.Frame(self.root, bg="#DBEAFE", bd=0)
        loc_frame.pack(fill="x")
        tk.Label(loc_frame, text=f"📍 내 좌표: {self.my_lat:.4f}, {self.my_lng:.4f}", bg="#DBEAFE", fg="#1E40AF", font=("맑은 고딕", 10, "bold"), pady=8).pack()

        card_frame = tk.Frame(self.root, bg="white", bd=0, highlightthickness=1, highlightbackground="#E5E7EB")
        card_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        form = tk.Frame(card_frame, bg="white", padx=15, pady=10)
        form.pack(fill="both", expand=True)

        tk.Label(form, text="🚌 탑승할 노선 (타이핑하여 검색)", bg="white", fg="#374151", font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.route_combo = AutocompleteCombobox(form, font=("맑은 고딕", 13))
        self.route_combo.pack(fill="x", pady=(0, 15))
        self.route_combo.bind("<<ComboboxSelected>>", self.on_route_selected)
        self.route_combo.bind("<Return>", self.on_route_selected)

        tk.Label(form, text="📍 주변 정류장 (노선 진행 순서)", bg="white", fg="#374151", font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.stop_combo = ttk.Combobox(form, values=["노선을 먼저 선택하세요"], font=("맑은 고딕", 13), state="readonly")
        self.stop_combo.current(0)
        self.stop_combo.pack(fill="x", pady=(0, 15))

        display_frame = tk.Frame(form, bg="#F3F4F6", bd=0)
        display_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        self.realtime_info = tk.Text(display_frame, height=2, font=("맑은 고딕", 11, "bold"), 
                                     bg="#F3F4F6", fg="#1F2937", bd=0, highlightthickness=0, padx=10, pady=5)
        self.realtime_info.pack(fill="x")
        self.realtime_info.insert("1.0", "환영합니다! 노선을 검색해주세요.")
        self.realtime_info.config(state="disabled")

        self.map_widget = tkintermapview.TkinterMapView(display_frame, width=400, height=200, corner_radius=10)
        self.map_widget.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.map_widget.set_position(self.my_lat, self.my_lng)
        self.map_widget.set_zoom(15)
        self.my_marker = self.map_widget.set_marker(self.my_lat, self.my_lng, text="내 위치", marker_color_circle="#3B82F6")
        
        self.bus_markers = [] 
        
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(current_dir, "bus.png")
            
            raw_image = Image.open(icon_path)
            # 💡 [핵심 수정 1] resize 대신 thumbnail을 쓰면, 원본 비율을 철저히 유지하면서 최대 40 크기에 맞춰줍니다! (찌그러짐 완벽 해결)
            raw_image.thumbnail((40, 40)) 
            self.bus_icon = ImageTk.PhotoImage(raw_image, master=self.root) 
            print(f"✅ 버스 아이콘 로드 성공! ({icon_path})")
        except Exception as e:
            print(f"❌ 버스 아이콘 로드 실패: {e}")
            self.bus_icon = None

        search_btn = self.create_flat_button(form, "현재 버스 위치 조회 📡", "#10B981", "#059669", self.check_real_bus, ("맑은 고딕", 13, "bold"), 12)
        search_btn.pack(fill="x", pady=(0, 10))

        req_btn = self.create_flat_button(form, "승차 요청 (알림) 🔔", "#3B82F6", "#2563EB", self.on_ride_request, ("맑은 고딕", 15, "bold"), 15)
        req_btn.pack(fill="x")

    def load_routes_from_data(self):
        if not DF_STOPS.empty:
            raw_routes = DF_STOPS.iloc[:, 1].dropna().astype(str).unique()
            self.all_route_list = sorted([str(r).strip() for r in raw_routes])
            self.route_combo.set_completion_list(self.all_route_list)

    def on_route_selected(self, event):
        route = self.route_combo.get()
        if not route or DF_STOPS.empty: return
        self.display_realtime_text(f"{route}번 정류장 업데이트 완료!")
        
        try:
            col_route_id = DF_STOPS.columns[0]
            col_route_name = DF_STOPS.columns[1]
            col_ord = DF_STOPS.columns[2]
            col_node_id = DF_STOPS.columns[3]
            col_stop_name = DF_STOPS.columns[5]
            col_x = DF_STOPS.columns[6] 
            col_y = DF_STOPS.columns[7] 
            
            search_route = route.replace("번", "").strip()

            if not DF_ROUTES.empty:
                try:
                    matched_route = DF_ROUTES[DF_ROUTES['ROUTE_NM'].astype(str) == search_route]
                    if not matched_route.empty:
                        self.current_route_id = str(matched_route.iloc[0]['ROUTE_ID'])
                    else:
                        raise Exception("매핑 파일에 노선이 없습니다.")
                except Exception as e:
                    fallback_df = DF_STOPS[DF_STOPS[col_route_name].astype(str) == search_route]
                    if not fallback_df.empty:
                        self.current_route_id = str(fallback_df.iloc[0][col_route_id])
            else:
                fallback_df = DF_STOPS[DF_STOPS[col_route_name].astype(str) == search_route]
                if not fallback_df.empty:
                    self.current_route_id = str(fallback_df.iloc[0][col_route_id])
            
            route_df = DF_STOPS[DF_STOPS[col_route_name].astype(str) == search_route].copy()
            
            if route_df.empty: return
            
            route_df[col_ord] = pd.to_numeric(route_df[col_ord])
            route_df = route_df.sort_values(by=col_ord)
            
            combo_values = []
            self.stop_data_map.clear()
            for _, row in route_df.iterrows():
                raw_stop_name = str(row[col_stop_name])
                stop_lat = float(row[col_y])
                stop_lng = float(row[col_x])
                
                dist_km = haversine(self.my_lat, self.my_lng, stop_lat, stop_lng)
                display_name = f"{raw_stop_name} ({dist_km:.2f}km)"
                combo_values.append(display_name)
                
                self.stop_data_map[display_name] = {
                    "stop_id": str(row[col_node_id]), 
                    "ord": str(row[col_ord]),
                    "raw_name": raw_stop_name 
                }
            
            self.stop_combo['values'] = tuple(combo_values)
            if combo_values:
                self.stop_combo.current(0)
        except Exception as e:
            print("데이터 파싱 오류:", e)

    def check_real_bus(self):
        selected_stop_text = self.stop_combo.get()
        if not self.current_route_id or selected_stop_text not in self.stop_data_map:
            return

        stop_info = self.stop_data_map[selected_stop_text]
        route_id = self.current_route_id
        target_ord = int(stop_info["ord"])
        
        url = "http://ws.bus.go.kr/api/rest/buspos/getBusPosByRtid"
        params = {'ServiceKey': DATA_GO_KR_API_KEY, 'busRouteId': route_id, 'resultType': 'json'}
        full_url = f"{url}?{urllib.parse.urlencode(params)}"

        try:
            res = requests.get(full_url, timeout=3)
            data = res.json()
            
            if data['msgHeader']['headerCd'] == '0':
                item_list = data['msgBody'].get('itemList', [])
                approaching_buses = [b for b in item_list if (target_ord - 3) <= int(b['sectOrd']) <= target_ord]
                
                for marker in self.bus_markers:
                    marker.delete()
                self.bus_markers.clear()
                
                if approaching_buses:
                    approaching_buses.sort(key=lambda x: int(x['sectOrd']), reverse=True)
                    
                    info_texts = []
                    for bus in approaching_buses:
                        stops_left = target_ord - int(bus['sectOrd'])
                        bus_lat = float(bus['gpsY'])
                        bus_lng = float(bus['gpsX'])
                        
                        # 💡 [핵심 수정 2] 문자열 슬라이싱 [-4:]를 사용해 무조건 맨 뒤 4글자만 싹둑 잘라냅니다!
                        plain_no = bus['plainNo'][-4:]
                        
                        arrmsg = "🚨 구간 진입!" if stops_left == 0 else f"{stops_left}번째 전"
                        info_texts.append(f"[{plain_no}: {arrmsg}]")
                        
                        if getattr(self, 'bus_icon', None) is not None:
                            marker = self.map_widget.set_marker(bus_lat, bus_lng, text=f" {plain_no}", icon=self.bus_icon)
                        else:
                            marker = self.map_widget.set_marker(bus_lat, bus_lng, text=f"🚌 {plain_no}", marker_color_circle="#EF4444")
                        
                        self.bus_markers.append(marker)
                    
                    self.display_realtime_text(" | ".join(info_texts))
                    closest_bus = approaching_buses[0]
                    self.map_widget.set_position(float(closest_bus['gpsY']), float(closest_bus['gpsX']))
                    self.map_widget.set_zoom(14)
                else:
                    self.display_realtime_text("현재 이전 3정거장 이내에 접근 중인 버스가 없습니다.")
            else:
                self.display_realtime_text(f"API 오류: {data['msgHeader']['headerMsg']}")
        except Exception as e:
            self.display_realtime_text(f"통신 에러 발생 (인증키 동기화 대기중일 수 있습니다)")

    def display_realtime_text(self, text):
        self.realtime_info.config(state="normal")
        self.realtime_info.delete("1.0", tk.END)
        self.realtime_info.insert(tk.END, text)
        self.realtime_info.config(state="disabled")

    def on_ride_request(self):
        raw_route = self.route_combo.get()
        display_stop = self.stop_combo.get()
        
        if not raw_route or not display_stop: return
        route_name = raw_route.replace("번", "").strip()
        
        stop_info = self.stop_data_map.get(display_stop)
        if not stop_info: return
        real_stop_name = stop_info["raw_name"]
        
        try:
            res = requests.post(f"{SERVER_URL}/request_ride", json={"route": route_name, "stop": real_stop_name}, timeout=3)
            if res.status_code == 200:
                messagebox.showinfo("승차 예약 완료", f"✅ {route_name}번 ({real_stop_name})\n기사님께 승차 알림이 전송되었습니다!")
        except:
            messagebox.showerror("오류", "서버와 연결할 수 없습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PassengerApp(root)
    root.mainloop()