import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import tkinter.font as tkfont
import os
import sys
from datetime import datetime
from tkcalendar import DateEntry
import warnings
import urllib.request # 인터넷 접속용 (업데이트 확인)

# 경고 무시
warnings.simplefilter(action='ignore', category=UserWarning)

# ===========================================================
# 🔄 [자동 업데이트 시스템] - 형님의 깃허브와 연동됨
# ===========================================================
GITHUB_USER = "DaonMaru"
REPO_NAME = "BoraSystem"
BRANCH = "main"
BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{REPO_NAME}/{BRANCH}"

def check_and_update():
    """
    서버(깃허브)의 version.txt를 확인하고,
    내 컴퓨터보다 최신 버전이면 bora_merge.py를 다운로드합니다.
    """
    try:
        # 1. 경로 설정
        user_profile = os.environ['USERPROFILE']
        # 원드라이브 우선, 없으면 바탕화면
        desktop_path = os.path.join(user_profile, 'OneDrive', '바탕 화면')
        if not os.path.exists(desktop_path):
            desktop_path = os.path.join(user_profile, 'Desktop')
        
        local_ver_file = os.path.join(desktop_path, 'version.txt')
        target_code_file = os.path.join(desktop_path, 'bora_merge.py')

        # 2. 내 컴퓨터 버전 확인 (없으면 0.0으로 간주)
        current_ver = 0.0
        if os.path.exists(local_ver_file):
            try:
                with open(local_ver_file, 'r') as f:
                    current_ver = float(f.read().strip())
            except:
                pass # 파일이 깨져있으면 0.0

        # 3. 서버(깃허브) 버전 확인
        ver_url = f"{BASE_URL}/version.txt"
        with urllib.request.urlopen(ver_url) as response:
            server_ver_str = response.read().decode('utf-8').strip()
            server_ver = float(server_ver_str)

        print(f"📡 버전 확인 - 내PC: {current_ver} / 서버: {server_ver}")

        # 4. 업데이트 진행 (서버 버전이 더 높으면)
        if server_ver > current_ver:
            print("🚀 업데이트 발견! 다운로드를 시작합니다...")
            
            # (1) bora_merge.py 다운로드
            code_url = f"{BASE_URL}/bora_merge.py"
            with urllib.request.urlopen(code_url) as response:
                code_data = response.read().decode('utf-8')
                with open(target_code_file, 'w', encoding='utf-8') as f:
                    f.write(code_data)
            
            # (2) 로컬 version.txt 업데이트
            with open(local_ver_file, 'w') as f:
                f.write(str(server_ver))
                
            return True, server_ver # 업데이트 성공했다는 신호

    except Exception as e:
        print(f"⚠️ 업데이트 확인 중 오류: {e}")
        return False, 0.0
    
    return False, 0.0 # 업데이트 없음

# ===========================================================
# [메인 프로그램 시작]
# ===========================================================

# 1. 시작하자마자 업데이트 체크
is_updated, new_ver = check_and_update()

# -----------------------------------------------------------
# [파일 경로 설정]
# -----------------------------------------------------------
def get_db_path():
    # 엑셀 파일 찾기 (바탕화면 등)
    user_profile = os.environ['USERPROFILE']
    paths_to_check = [
        os.path.join(user_profile, 'OneDrive', '바탕 화면', '보라물류_최종정밀단가표.xlsx'),
        os.path.join(user_profile, 'Desktop', '보라물류_최종정밀단가표.xlsx'),
        '보라물류_최종정밀단가표.xlsx' # 현재 폴더
    ]
    
    for p in paths_to_check:
        if os.path.exists(p):
            return p
            
    # 파일이 없으면 그냥 현재 폴더 경로 리턴 (나중에 생성됨)
    return '보라물류_최종정밀단가표.xlsx'

db_file = get_db_path()

class BoraUltimateApp:
    def __init__(self, root):
        self.root = root
        
        # [화면 크기]
        w, h = 1150, 850 
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2) - 50 
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))

        # [폰트]
        self.font_header = tkfont.Font(family="Malgun Gothic", size=20, weight="bold")
        self.font_default = tkfont.Font(family="Malgun Gothic", size=10)
        self.font_bold = tkfont.Font(family="Malgun Gothic", size=10, weight="bold")
        self.font_btn = tkfont.Font(family="Malgun Gothic", size=14, weight="bold")
        self.font_entry = tkfont.Font(family="Malgun Gothic", size=12)
        self.root.option_add('*Font', self.font_default)

        # -------------------------------------------------------
        # [데이터 로딩]
        # -------------------------------------------------------
        try:
            self.df = pd.read_excel(db_file)
            
            # 날짜 및 데이터 정리
            self.df['접수일자'] = pd.to_datetime(self.df['접수일자'], errors='coerce').dt.strftime('%Y-%m-%d')
            self.df = self.df.dropna(subset=['접수일자'])
            
            if '배달운임' in self.df.columns:
                self.df['배달운임'] = (
                    self.df['배달운임'].astype(str)
                    .str.replace(',', '')
                    .str.extract(r'(\d+)', expand=False)
                    .fillna(0).astype(int)
                )
            
            self.df = self.df[(self.df['차종_최종'] != "미분류") & (~self.df['차종_최종'].str.contains('P', na=False))]
            
            title_txt = f"보라물류 통합 시스템 V{new_ver if is_updated else '1.0'}"
            if is_updated: title_txt += " (✨업데이트 완료!)"
            self.root.title(title_txt)
            
        except Exception as e:
            # 파일이 없거나 오류나면 빈 껍데기 실행
            self.df = pd.DataFrame(columns=['접수일자', '고객성명', '차종_최종', '도 착 지', '배달운임'])
            self.root.title("보라물류 통합 시스템 (데이터 없음)")

        # 업데이트 알림 메시지
        if is_updated:
            messagebox.showinfo("업데이트 성공", f"서버에서 최신 통합 엔진(v{new_ver})을 받아왔습니다!\n이제 최신 로직으로 작동합니다.")

        self.search_timer = None

        # ========================================================
        # [UI 구성]
        # ========================================================
        btn_frame = tk.Frame(root, pady=10, bg="#eee")
        btn_frame.pack(side="bottom", fill="x")
        tk.Button(btn_frame, text="선택 항목 정산 및 영수증 발행 (팝업)", command=self.open_option_popup, 
                  bg="#6c5ce7", fg="white", font=self.font_btn, height=2).pack(fill="x", padx=20, pady=5)

        header = tk.Frame(root, pady=10)
        header.pack(side="top", fill="x")
        tk.Label(header, text="💜 보라물류 배차 시스템 💜", font=self.font_header, fg="#6c5ce7").pack()

        # 검색 필터 영역
        filter_frame = tk.LabelFrame(root, text="검색 필터", font=self.font_bold)
        filter_frame.pack(side="top", pady=5, padx=20, fill="x")
        sf = tk.Frame(filter_frame, pady=5); sf.pack()

        today = datetime.now(); first_day = today.replace(day=1)

        tk.Label(sf, text="기간:", font=self.font_bold, fg="blue").grid(row=0, column=0, padx=5)
        self.ent_start = DateEntry(sf, width=12, font=self.font_entry, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.ent_start.set_date(first_day)
        self.ent_start.grid(row=0, column=1, padx=2)
        self.ent_start.bind("<<DateEntrySelected>>", lambda e: self.search())

        tk.Label(sf, text="~").grid(row=0, column=2)
        self.ent_end = DateEntry(sf, width=12, font=self.font_entry, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.ent_end.set_date(today)
        self.ent_end.grid(row=0, column=3, padx=2)
        self.ent_end.bind("<<DateEntrySelected>>", lambda e: self.search())

        tk.Label(sf, text=" |  거래처:", font=self.font_default).grid(row=0, column=4, padx=5)
        self.ent_cust = tk.Entry(sf, width=15, font=self.font_entry); self.ent_cust.grid(row=0, column=5, padx=5, ipady=3)
        self.ent_cust.bind('<KeyRelease>', lambda e: self.smart_search()) 

        tk.Label(sf, text="도착지:", font=self.font_default).grid(row=0, column=6, padx=5)
        self.ent_dest = tk.Entry(sf, width=15, font=self.font_entry); self.ent_dest.grid(row=0, column=7, padx=5, ipady=3)
        self.ent_dest.bind('<KeyRelease>', lambda e: self.smart_search())
        
        tk.Button(sf, text="조회", command=self.search, bg="#6c5ce7", fg="white", width=8, font=self.font_bold).grid(row=0, column=8, padx=15)

        # 차종 필터
        type_frame = tk.LabelFrame(root, text="차종 분류 선택", font=self.font_bold)
        type_frame.pack(side="top", pady=5, padx=20, fill="x")
        
        self.check_vars = {}
        groups = {
            "🚀 퀵서비스": ["오토바이", "다마스", "라보"],
            "📦 혼적/합짐": ["혼적", "합짐"], 
            "🚚 중형운송": ["1톤", "1.4톤"],
            "🚛 대형운송": ["2.5톤", "3.5톤", "5톤", "11톤", "16톤", "18톤", "25톤"]
        }
        
        if not self.df.empty:
            raw_types = self.df['차종_최종'].unique().tolist()
        else:
            raw_types = []

        for force_item in ["혼적", "합짐"]:
            if force_item not in raw_types: raw_types.append(force_item)
            
        for g_name, keywords in groups.items():
            g_main_f = tk.Frame(type_frame, pady=2)
            g_main_f.pack(fill="x", padx=10)
            lbl_color = "#d63031" if "혼적" in g_name else "#4834d4"
            tk.Label(g_main_f, text=g_name, font=self.font_bold, width=15, anchor="w", fg=lbl_color).pack(side="left", anchor="nw")
            cb_container = tk.Frame(g_main_f)
            cb_container.pack(side="left", fill="x", expand=True)
            
            matched_types = [t for t in raw_types if any(k in str(t) for k in keywords)]
            if "중형" in g_name: matched_types = [t for t in matched_types if not any(x in str(t) for x in ["2.5", "3.5", "5"])]
            
            def sort_key(name):
                t = str(name)
                if '2.5' in t: return 1
                if '3.5' in t: return 2
                if '5톤' in t: return 3
                return 99

            for i, t_name in enumerate(sorted(matched_types, key=sort_key)):
                var = tk.BooleanVar()
                cb = tk.Checkbutton(cb_container, text=t_name, variable=var, command=self.search, font=self.font_default)
                cb.grid(row=i//5, column=i%5, padx=5, pady=0, sticky="w")
                self.check_vars[t_name] = var

        # 트리뷰(리스트)
        list_frame = tk.Frame(root)
        list_frame.pack(side="top", pady=5, padx=20, fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(list_frame, orient="vertical")
        scrollbar_x = ttk.Scrollbar(list_frame, orient="horizontal")

        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Malgun Gothic", 10))
        style.configure("Treeview.Heading", font=("Malgun Gothic", 10, "bold"))
        
        self.tree = ttk.Treeview(list_frame, columns=("날짜", "거래처", "차종", "도착지", "단가"), show="headings", 
                                 yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.heading("날짜", text="날짜"); self.tree.heading("거래처", text="거래처명")
        self.tree.heading("차종", text="차종/옵션"); self.tree.heading("도착지", text="도착지 상세"); self.tree.heading("단가", text="기존단가")
        self.tree.column("날짜", width=100, anchor="center"); self.tree.column("거래처", width=160)
        self.tree.column("차종", width=180, anchor="center"); self.tree.column("도착지", width=500); self.tree.column("단가", width=110, anchor="e")
        self.tree.bind("<Double-1>", lambda e: self.open_option_popup())

        self.search() 
        self.ent_cust.focus_set()

    def smart_search(self):
        if self.search_timer is not None:
            self.root.after_cancel(self.search_timer)
        self.search_timer = self.root.after(300, self.search)

    def search(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        
        c = self.ent_cust.get().strip().upper()
        d = self.ent_dest.get().strip().upper()
        s_date = self.ent_start.get_date().strftime("%Y-%m-%d")
        e_date = self.ent_end.get_date().strftime("%Y-%m-%d")

        temp = self.df.copy()
        if temp.empty: return

        if s_date and e_date:
            temp = temp[(temp['접수일자'] >= s_date) & (temp['접수일자'] <= e_date)]

        if c: temp = temp[temp['고객성명'].str.contains(c, na=False)]
        if d == "혼적":
            temp = temp[temp['도 착 지'].str.contains("혼적", na=False) | temp['차종_최종'].str.contains("혼적", na=False)]
        elif d: 
            temp = temp[temp['도 착 지'].str.contains(d, na=False)]
        
        selected_types = [n for n, v in self.check_vars.items() if v.get()]
        if selected_types:
            condition = temp['차종_최종'].isin(selected_types)
            if "혼적" in selected_types or "합짐" in selected_types:
                mixed_cond = temp['도 착 지'].str.contains("혼적|합짐", na=False) | temp['차종_최종'].str.contains("혼적|합짐", na=False)
                temp = temp[condition | mixed_cond]
            else:
                temp = temp[condition]

        for _, r in temp.iterrows():
            try:
                fare_val = int(r['배달운임'])
                fare_str = f"{fare_val:,}"
            except:
                fare_str = "0"
            self.tree.insert("", "end", values=(r['접수일자'], r['고객성명'], r['차종_최종'], r['도 착 지'], fare_str))

    def open_option_popup(self):
        sel = self.tree.selection()
        if not sel: 
            messagebox.showwarning("경고", "먼저 목록에서 항목을 선택해주세요.")
            return
        
        item = self.tree.item(sel[0])['values']
        try: base_fare = int(str(item[4]).replace(",", ""))
        except: base_fare = 0
            
        car_type, cust_name, dest_name = str(item[2]), str(item[1]), str(item[3])

        pop = tk.Toplevel(self.root)
        pop.title("상세 견적 및 영수증 발행")
        pop.geometry("700x750")
        x = self.root.winfo_x() + (self.root.winfo_width()//2) - 350
        y = self.root.winfo_y() + (self.root.winfo_height()//2) - 375
        pop.geometry(f"700x750+{x}+{y}")
        pop.focus_set()

        info_frame = tk.LabelFrame(pop, text="선택된 운송 건", font=self.font_bold, padx=10, pady=10)
        info_frame.pack(fill="x", padx=10, pady=10)
        tk.Label(info_frame, text=f"날짜: {item[0]}   |   거래처: {item[1]}", font=("Malgun Gothic", 12)).pack(anchor="w")
        tk.Label(info_frame, text=f"차종: {car_type}   |   도착지: {item[3]}", font=("Malgun Gothic", 12)).pack(anchor="w")
        tk.Label(info_frame, text=f"기본 운임: {base_fare:,}원", font=("Malgun Gothic", 14, "bold"), fg="#4834d4").pack(anchor="w", pady=5)

        opt_frame = tk.LabelFrame(pop, text="추가 옵션 설정", font=self.font_bold, padx=10, pady=10)
        opt_frame.pack(fill="x", padx=10, pady=5)

        v_round, v_sun, v_tax = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar(value=True)
        v_wait_min, v_urgent, v_rack = tk.StringVar(value="0"), tk.IntVar(value=0), tk.IntVar(value=0)

        tk.Checkbutton(opt_frame, text="왕복 운행 (x1.7)", variable=v_round, font=self.font_default).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Checkbutton(opt_frame, text="휴일/야간 (+1만)", variable=v_sun, fg="red", font=self.font_default).grid(row=0, column=1, sticky="w", padx=10)
        tk.Checkbutton(opt_frame, text="부가세 별도 발행", variable=v_tax, font=self.font_bold).grid(row=0, column=2, sticky="w", padx=10)
        
        tk.Label(opt_frame, text="대기시간(분):", font=self.font_bold).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        entry_wait = tk.Entry(opt_frame, textvariable=v_wait_min, width=5, justify="center", font=self.font_default)
        entry_wait.grid(row=1, column=1, sticky="w")
        tk.Label(opt_frame, text="(10분당 1천원)", fg="gray", font=("Malgun Gothic", 8)).grid(row=1, column=2, sticky="w")

        moto_frame = tk.LabelFrame(opt_frame, text="오토바이 전용", fg="purple")
        moto_frame.grid(row=2, column=0, columnspan=3, sticky="we", pady=10)
        
        tk.Radiobutton(moto_frame, text="일반", variable=v_urgent, value=0).pack(side="left", padx=5)
        tk.Radiobutton(moto_frame, text="긴급(+1만)", variable=v_urgent, value=10000, fg="orange").pack(side="left", padx=5)
        tk.Radiobutton(moto_frame, text="우천(+2만)", variable=v_urgent, value=20000, fg="red").pack(side="left", padx=5)
        tk.Checkbutton(moto_frame, text="짐받이(+5천)", variable=v_rack, onvalue=5000, offvalue=0).pack(side="left", padx=10)

        if "오토바이" not in car_type:
            for child in moto_frame.winfo_children(): child.configure(state='disabled')

        res_frame = tk.LabelFrame(pop, text="상세 견적 내역", font=self.font_bold, fg="blue", padx=10, pady=10)
        res_frame.pack(fill="both", expand=True, padx=10, pady=5)
        lbl_detail = tk.Label(res_frame, text="계산 버튼을 누르면 상세 내역이 표시됩니다.", justify="left", font=("Malgun Gothic", 11), bg="#f1f2f6", anchor="nw", width=60, height=10)
        lbl_detail.pack(fill="both", expand=True)

        def create_receipt_excel(data_dict):
            try:
                receipt_data = [
                    ["보라물류 운송 영수증(견적서)", ""], ["", ""], ["", ""],
                    ["[ 공급자 정보 ]", ""], ["등록번호", "123-86-13156"], ["상    호", "보라물류"], ["대 표 자", "백병순"],
                    ["주    소", "경기도 군포시 당정동 103-3 1층"], ["업    태", "운수"], ["종    목", "퀵서비스, 운송주선, 화물운송"],
                    ["", ""],
                    ["[ 운송 내역 ]", ""], ["일    자", datetime.now().strftime("%Y-%m-%d")],
                    ["공급받는자", cust_name], ["운행구간", dest_name], ["차    종", car_type],
                    ["", ""],
                    ["[ 금액 산출 내역 ]", ""], ["항    목", "금    액"], ["--------------------", "--------------------"],
                    ["기본 운임", f"{data_dict['기본운임']:,}"],
                ]
                
                if data_dict['왕복할증'] > 0: receipt_data.append(["왕복 할증", f"{data_dict['왕복할증']:,}"])
                if data_dict['대기료'] > 0: receipt_data.append(["대기료", f"{data_dict['대기료']:,}"])
                if data_dict['휴일할증'] > 0: receipt_data.append(["휴일/야간 할증", f"{data_dict['휴일할증']:,}"])
                if data_dict['기타할증'] > 0: receipt_data.append(["오토바이/기타 할증", f"{data_dict['기타할증']:,}"])
                
                receipt_data.extend([["", ""], ["공급가액", f"{data_dict['공급가액']:,}"], ["부 가 세", f"{data_dict['부가세']:,}"], ["", ""], ["총 합 계", f"{data_dict['최종청구금액']:,}"], ["", ""], ["위 금액을 정히 영수(청구)합니다.", ""], ["보라물류 (인)", ""]])

                df_receipt = pd.DataFrame(receipt_data, columns=["항목", "내용"])
                user_profile = os.environ['USERPROFILE']
                save_dir = os.path.join(user_profile, 'Desktop')
                if not os.path.exists(save_dir): save_dir = os.path.join(user_profile, '바탕 화면')
                
                filename = f"{datetime.now().strftime('%Y%m%d_%H%M')}_{cust_name.replace('/', '')}_영수증.xlsx"
                save_path = os.path.join(save_dir, filename)
                df_receipt.to_excel(save_path, index=False, header=False)
                messagebox.showinfo("발행 완료", f"영수증이 저장되었습니다!\n위치: {save_dir}")
            except Exception as e: messagebox.showerror("실패", f"영수증 생성 오류: {e}")

        def calc_final():
            current_fare = base_fare
            data_row = {"기본운임": base_fare, "왕복할증": 0, "대기료": 0, "휴일할증": 0, "기타할증": 0, "공급가액": 0, "부가세": 0, "최종청구금액": 0}
            detail_text = f"■ 기본 운임: {base_fare:,}원\n" + "-" * 40 + "\n"

            if v_round.get():
                added = int(current_fare * 0.7)
                current_fare += added; data_row["왕복할증"] = added
                detail_text += f"+ [왕복 할증] 70% 추가: {added:,}원\n"
            
            try: mins = int(v_wait_min.get())
            except: mins = 0
            if mins > 0:
                wait_cost = (mins // 10) * 1000
                current_fare += wait_cost; data_row["대기료"] = wait_cost
                detail_text += f"+ [대기료] {mins}분: {wait_cost:,}원\n"

            if v_sun.get():
                current_fare += 10000; data_row["휴일할증"] = 10000
                detail_text += f"+ [휴일/야간] 할증: 10,000원\n"

            if "오토바이" in car_type:
                total_moto_add = v_urgent.get() + v_rack.get()
                if total_moto_add > 0:
                    current_fare += total_moto_add; data_row["기타할증"] = total_moto_add
                    detail_text += f"+ [오토바이] 옵션: {total_moto_add:,}원\n"

            supply_price = int(current_fare); data_row["공급가액"] = supply_price
            detail_text += "-" * 40 + "\n" + f"▶ 공급가액: {supply_price:,}원\n"
            
            final_total = supply_price
            if v_tax.get():
                vat = int(supply_price * 0.1)
                final_total += vat; data_row["부가세"] = vat
                detail_text += f"▶ 부가세(10%): {vat:,}원\n"
            
            data_row["최종청구금액"] = final_total
            detail_text += "=" * 40 + "\n" + f"💰 최종 청구 금액: {final_total:,}원"

            lbl_detail.config(text=detail_text, fg="#2d3436")
            btn_receipt.config(state="normal", command=lambda: create_receipt_excel(data_row))
        
        btn_box = tk.Frame(pop, pady=10); btn_box.pack(side="bottom", fill="x")
        tk.Button(btn_box, text="견적 산출 (Enter)", command=calc_final, bg="#6c5ce7", fg="white", font=("Malgun Gothic", 12, "bold"), height=2, width=20).pack(side="left", padx=20, expand=True)
        btn_receipt = tk.Button(btn_box, text="🖨️ 영수증(Excel) 발행", state="disabled", bg="#27ae60", fg="white", font=("Malgun Gothic", 12, "bold"), height=2, width=20)
        btn_receipt.pack(side="right", padx=20, expand=True)
        pop.bind('<Return>', lambda e: calc_final())

if __name__ == "__main__":
    root = tk.Tk(); app = BoraUltimateApp(root); root.mainloop()
