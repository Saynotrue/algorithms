import tkinter as tk
from tkinter import ttk, messagebox
import requests
import pandas as pd
import os

SERVER_URL = "http://127.0.0.1:8000"

# -------------------------------------------------------------------
# 1. 엑셀(CSV) 데이터 로드 (승객용/기사용 공통)
# -------------------------------------------------------------------
FILE_NAME = "서울시버스노선별정류소정보(20260506).xlsx"
DF_STOPS = pd.DataFrame()

try:
    if os.path.exists(FILE_NAME):
        if FILE_NAME.endswith(".csv"):
            DF_STOPS = pd.read_csv(FILE_NAME)
        else:
            DF_STOPS = pd.read_excel(FILE_NAME)
        print("✅ 버스 노선 데이터 로드 성공!")
    else:
        print(f"❌ 오류: '{FILE_NAME}' 파일을 찾을 수 없습니다.")
except Exception as e:
    print(f"❌ 데이터 로드 실패: {e}")

# -------------------------------------------------------------------
# 2. 자동완성 콤보박스 위젯 (Mac OS 포커스 버그 완벽 해결판)
# -------------------------------------------------------------------
class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list)
        self['values'] = self._completion_list
        self.bind('<KeyRelease>', self.handle_keyrelease)
        
    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down', 'Return', 'Shift_L', 'Shift_R', 'Tab'):
            return
            
        typed_text = self.get()
        if typed_text == '':
            self['values'] = self._completion_list
        else:
            filtered = [item for item in self._completion_list if typed_text.lower() in item.lower()]
            self['values'] = filtered

# -------------------------------------------------------------------
# 3. 기사용 앱 메인 클래스 (승객용 UI 패밀리룩 적용)
# -------------------------------------------------------------------
class DriverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("스마트 버스 - 기사용 단말기")
        self.root.geometry("450x700") # 승객용과 동일한 사이즈
        
        # 🎨 승객용 앱과 동일한 밝은 회색 배경
        self.root.configure(bg="#F4F6F9") 

        # 🎨 기본 콤보박스 테마 모던하게 설정 (전역 적용)
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        style.configure('TCombobox', padding=6, arrowsize=14)

        self.current_route = ""
        self.next_stop = ""
        self.is_running = False 
        
        self.all_route_list = []
        self.current_route_stops = [] 
        self.current_stop_index = 0

        self.load_routes_from_data()
        self.show_login_screen()

    def load_routes_from_data(self):
        if not DF_STOPS.empty:
            col_route_name = DF_STOPS.columns[1] 
            unique_routes = DF_STOPS[col_route_name].dropna().astype(str).unique().tolist()
            self.all_route_list = sorted(unique_routes)

    def create_flat_button(self, parent, text, bg_color, hover_color, command, font_style, pady_val):
        btn = tk.Label(parent, text=text, bg=bg_color, fg="white", font=font_style, pady=pady_val, cursor="hand2")
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_color))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg_color))
        btn.bind("<Button-1>", lambda e: command())
        return btn

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    # -------------------------------------------------------------------
    # 화면 1: 운행 시작 전 설정 화면
    # -------------------------------------------------------------------
    def show_login_screen(self):
        self.is_running = False 
        self.clear_window()

        # 🎨 파란색 헤더 영역
        header_frame = tk.Frame(self.root, bg="#2563EB", bd=0)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="스마트 버스 운행 (기사용)", bg="#2563EB", fg="white", 
                 font=("맑은 고딕", 16, "bold"), pady=15).pack()

        # 🎨 하얀색 메인 카드 영역
        card_frame = tk.Frame(self.root, bg="white", bd=0, highlightthickness=1, highlightbackground="#E5E7EB")
        card_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        form = tk.Frame(card_frame, bg="white", padx=20, pady=15)
        form.pack(fill="both", expand=True)

        # 안내 문구
        tk.Label(form, text="운행 시작 설정", bg="white", fg="#1F2937", 
                 font=("맑은 고딕", 20, "bold")).pack(pady=(10, 40))

        # 노선 검색
        tk.Label(form, text="🚌 운행 노선 (타이핑하여 검색)", bg="white", fg="#374151", 
                 font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.route_entry = AutocompleteCombobox(form, font=("맑은 고딕", 14))
        self.route_entry.pack(fill="x", pady=(0, 25))
        if self.all_route_list:
            self.route_entry.set_completion_list(self.all_route_list)
        
        self.route_entry.bind("<<ComboboxSelected>>", self.update_stop_list)
        self.route_entry.bind("<Return>", self.update_stop_list)

        # 정류장 선택
        tk.Label(form, text="📍 현재/다음 정류장", bg="white", fg="#374151", 
                 font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.stop_entry = ttk.Combobox(form, values=["노선을 먼저 선택하세요"], font=("맑은 고딕", 14), state="readonly")
        self.stop_entry.pack(fill="x", pady=(0, 50))

        # 🎨 승객용 앱과 동일한 파란색 버튼
        start_btn = self.create_flat_button(form, "운행 시작 🚀", "#3B82F6", "#2563EB", 
                                            self.start_driving, ("맑은 고딕", 16, "bold"), 15)
        start_btn.pack(fill="x")

    def update_stop_list(self, event=None):
        route = self.route_entry.get().strip()
        if not route or DF_STOPS.empty: return
        
        col_route_name = DF_STOPS.columns[1]
        col_ord        = DF_STOPS.columns[2]
        col_stop_name  = DF_STOPS.columns[5]
        
        route_df = DF_STOPS[DF_STOPS[col_route_name].astype(str) == route]
        if not route_df.empty:
            route_df = route_df.sort_values(by=col_ord)
            stops = route_df[col_stop_name].astype(str).tolist()
            stops = list(dict.fromkeys(stops))
            
            self.current_route_stops = stops 
            
            self.stop_entry['values'] = stops
            if stops:
                self.stop_entry.current(0) 

    def start_driving(self):
        route = self.route_entry.get().strip()
        stop = self.stop_entry.get().strip()

        if not route or not stop or stop == "노선을 먼저 선택하세요":
            messagebox.showwarning("경고", "노선과 정류장을 올바르게 선택해주세요.")
            return

        self.current_route = route
        self.next_stop = stop
        self.is_running = True

        if stop in self.current_route_stops:
            self.current_stop_index = self.current_route_stops.index(stop)
        else:
            self.current_stop_index = 0

        self.show_dashboard_screen()
        self.update_dashboard()

    # -------------------------------------------------------------------
    # 화면 2: 운행 중 대시보드 화면
    # -------------------------------------------------------------------
    def show_dashboard_screen(self):
        self.clear_window()

        # 🎨 파란색 헤더 영역 (뒤로가기 버튼 포함)
        header_frame = tk.Frame(self.root, bg="#2563EB", bd=0)
        header_frame.pack(fill="x")
        
        back_btn = tk.Label(header_frame, text="◀", bg="#2563EB", fg="white", font=("맑은 고딕", 14, "bold"), cursor="hand2")
        back_btn.pack(side="left", padx=15)
        back_btn.bind("<Button-1>", lambda e: self.show_login_screen())

        tk.Label(header_frame, text=f"{self.current_route}번 운행 대시보드", bg="#2563EB", fg="white", 
                 font=("맑은 고딕", 16, "bold"), pady=15).pack(side="left", fill="x", expand=True)

        # 🎨 하얀색 메인 카드 영역
        card_frame = tk.Frame(self.root, bg="white", bd=0, highlightthickness=1, highlightbackground="#E5E7EB")
        card_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        d_info = tk.Frame(card_frame, bg="white", padx=20, pady=15)
        d_info.pack(fill="both", expand=True)

        tk.Label(d_info, text="이번 정류장", bg="white", fg="#6B7280", font=("맑은 고딕", 14, "bold")).pack(pady=(15, 0))
        
        # 🎨 정류장 이름 (강조된 진한 회색)
        self.stop_name_label = tk.Label(d_info, text=f"{self.next_stop}", bg="white", fg="#1F2937", font=("맑은 고딕", 24, "bold"))
        self.stop_name_label.pack(pady=(0, 25))

        # 🎨 대기 승객 수 표시 (승객용 앱의 텍스트 박스와 유사한 연한 회색 박스)
        count_frame = tk.Frame(d_info, bg="#F3F4F6", relief="flat", bd=0)
        count_frame.pack(fill="x", pady=10, ipady=30)
        
        tk.Label(count_frame, text="대기 승객", bg="#F3F4F6", fg="#374151", font=("맑은 고딕", 14, "bold")).pack(pady=(10, 0))
        
        self.waiting_count_label = tk.Label(count_frame, text="0", bg="#F3F4F6", fg="#10B981", font=("맑은 고딕", 64, "bold"))
        self.waiting_count_label.pack()

        # 🎨 승객용 앱과 동일한 초록색 버튼
        arrive_btn = self.create_flat_button(d_info, "정차 / 다음 정류장 🔄", "#10B981", "#059669", 
                                             self.on_arrive, ("맑은 고딕", 15, "bold"), 15)
        arrive_btn.pack(fill="x", side="bottom", pady=(20, 0))

    # -------------------------------------------------------------------
    # 서버 통신 및 정류장 이동 로직
    # -------------------------------------------------------------------
    def update_dashboard(self):
        if not self.is_running: return

        try:
            res = requests.get(f"{SERVER_URL}/check_stop/{self.current_route}/{self.next_stop}", timeout=2)
            if res.status_code == 200:
                count = res.json().get("count", 0)
                # 🎨 0명이면 초록색(#10B981), 1명 이상이면 붉은색(#EF4444)으로 알림!
                color = "#EF4444" if count > 0 else "#10B981"
                self.waiting_count_label.config(text=str(count), fg=color)
        except Exception as e:
            self.waiting_count_label.config(text="오류", font=("맑은 고딕", 20), fg="#9CA3AF")

        self.root.after(100, self.update_dashboard)

    def on_arrive(self):
        """버튼을 누르면 현재 정류장 인원을 0으로 만들고, 화면을 다음 정류장으로 바꿉니다."""
        try:
            res = requests.post(f"{SERVER_URL}/clear_stop", json={"route": self.current_route, "stop": self.next_stop})
            
            if res.status_code == 200:
                self.current_stop_index += 1
                
                if self.current_stop_index < len(self.current_route_stops):
                    self.next_stop = self.current_route_stops[self.current_stop_index]
                    
                    self.stop_name_label.config(text=self.next_stop)
                    self.waiting_count_label.config(text="0", fg="#10B981")
                else:
                    messagebox.showinfo("운행 종료", "해당 노선의 종점에 도착했습니다. 수고하셨습니다!")
                    self.show_login_screen()
        except:
            messagebox.showerror("오류", "서버 통신 실패. Uvicorn 서버가 켜져 있는지 확인하세요.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DriverApp(root)
    root.mainloop()