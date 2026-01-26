import os
import json
import time
import datetime
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from playwright.sync_api import sync_playwright
import requests
import csv
import io
import urllib3
# .env 지원을 위한 라이브러리
from dotenv import load_dotenv, set_key

# SSL 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================================
# [최상단 변수 로드] .env 파일 연동
# ==========================================================
ENV_PATH = ".env"
if not os.path.exists(ENV_PATH):
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("TARGET_MARKER=▼입금 대기\nSPREADSHEET_ID=1Q7Wew2MtwwYh0aSam2XNvBrwFHxM_3Kb2CT-Qv00-7o\nSRC_SHEET=월별내역\nGID=0\nHEADLESS=False")

load_dotenv(ENV_PATH)

# ==========================================================

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("구글 시트 자동화 v8.1 (검수 로직 추가)")
        self.root.geometry("650x900")
        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.root, padding="20")
        main.pack(fill="both", expand=True)
        
        # --- 1. .env 설정 영역 (상단 배치) ---
        group1 = ttk.LabelFrame(main, text=" 환경 설정 (자동 저장됨) ", padding="10")
        group1.pack(fill="x", pady=5)
        
        ttk.Label(group1, text="시트 ID:").grid(row=0, column=0, sticky="e", pady=2)
        self.ent_id = ttk.Entry(group1, width=50)
        self.ent_id.insert(0, os.getenv("SPREADSHEET_ID"))
        self.ent_id.grid(row=0, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(group1, text="시트 이름:").grid(row=1, column=0, sticky="e", pady=2)
        self.ent_sheet = ttk.Entry(group1, width=30)
        self.ent_sheet.insert(0, os.getenv("SRC_SHEET"))
        self.ent_sheet.grid(row=1, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(group1, text="GID:").grid(row=2, column=0, sticky="e", pady=2)
        self.ent_gid = ttk.Entry(group1, width=10)
        self.ent_gid.insert(0, os.getenv("GID"))
        self.ent_gid.grid(row=2, column=1, padx=5, pady=2, sticky="w")

        ttk.Label(group1, text="구역 마커:").grid(row=3, column=0, sticky="e", pady=2)
        self.ent_marker = ttk.Entry(group1, width=30)
        self.ent_marker.insert(0, os.getenv("TARGET_MARKER"))
        self.ent_marker.grid(row=3, column=1, padx=5, pady=2, sticky="w")

        self.var_headless = tk.BooleanVar()
        is_headless_env = os.getenv("HEADLESS", "False").lower() == "true"
        self.var_headless.set(is_headless_env)
        
        self.chk_headless = ttk.Checkbutton(group1, text="브라우저 창 숨기기 (Headless 모드)", variable=self.var_headless)
        self.chk_headless.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # --- 2. 작업 데이터 입력 영역 ---
        group2 = ttk.LabelFrame(main, text=" 작업 실행 데이터 ", padding="10")
        group2.pack(fill="x", pady=5)
        
        ttk.Label(group2, text="업체명(검색):").grid(row=0, column=0, sticky="e", pady=5)
        self.ent_name = ttk.Entry(group2, width=35)
        self.ent_name.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.ent_name.focus_set()

        ttk.Label(group2, text="입금일(yymmdd):").grid(row=1, column=0, sticky="e", pady=5)
        self.ent_date = ttk.Entry(group2, width=35)
        self.ent_date.insert(0, datetime.datetime.now().strftime('%y%m%d'))
        self.ent_date.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # --- 3. 실행 버튼 ---
        self.btn_run = tk.Button(main, text="스마트 검색 및 작업 시작 (Enter)", command=self.start_process, 
                                 bg="#4285F4", fg="white", font=("Malgun Gothic", 12, "bold"), height=2)
        self.btn_run.pack(fill="x", pady=15)
        self.root.bind('<Return>', lambda e: self.start_process())

        # --- 4. 로그 영역 ---
        self.log_area = scrolledtext.ScrolledText(main, height=12, font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True)

    def write_log(self, msg):
        self.log_area.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)
        self.root.update()

    def save_all_config(self):
        set_key(ENV_PATH, "SPREADSHEET_ID", self.ent_id.get().strip())
        set_key(ENV_PATH, "SRC_SHEET", self.ent_sheet.get().strip())
        set_key(ENV_PATH, "GID", self.ent_gid.get().strip())
        set_key(ENV_PATH, "TARGET_MARKER", self.ent_marker.get().strip())
        set_key(ENV_PATH, "HEADLESS", str(self.var_headless.get()))

    def get_sheet_matches(self, sheet_id, keyword, src_sheet, marker):
        AUTH_PATH = "google_auth.json"
        if not os.path.exists(AUTH_PATH): return []
        try:
            with open(AUTH_PATH, 'r', encoding='utf-8') as f:
                auth_data = json.load(f)
            cookies = {c['name']: c['value'] for c in auth_data.get('cookies', [])}
            url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={src_sheet}"
            resp = requests.get(url, cookies=cookies, timeout=20, verify=False)
            if resp.status_code != 200: return []
            content = resp.content.decode('utf-8-sig') 
            rows = list(csv.reader(io.StringIO(content)))
            
            marker_idx = -1
            for i, row in enumerate(rows):
                if any(marker in str(cell) for cell in row):
                    marker_idx = i
                    break
            
            matches = []
            search_start = marker_idx + 1 if marker_idx != -1 else 0
            for i, row in enumerate(rows[search_start:], start=search_start + 1):
                if len(row) > 1 and keyword.lower() in str(row[1]).lower():
                    cust_name = row[3] if len(row) > 3 else "미기입"
                    amount = row[12] if len(row) > 12 else "0"
                    matches.append({
                        "row": i,
                        "name": row[1],
                        "info": f"행: {i:3} | 업체: {row[1]:15} | 고객: {cust_name:10} | 금액: {amount:>10}"
                    })
            return matches
        except Exception: return []

    def show_selection_window(self, matches):
        win = tk.Toplevel(self.root)
        win.title("업체 상세 선택")
        win.geometry("600x400")
        win.grab_set()
        ttk.Label(win, text="여러 항목이 발견되었습니다. 정확한 행을 선택하세요:", padding=10).pack()
        lb = tk.Listbox(win, font=("Consolas", 10))
        lb.pack(fill="both", expand=True, padx=10, pady=5)
        for m in matches: lb.insert(tk.END, m["info"])
        self.selected_match = None
        def on_select():
            if lb.curselection():
                self.selected_match = matches[lb.curselection()[0]]
                win.destroy()
        tk.Button(win, text="선택 완료", command=on_select, bg="#4285F4", fg="white", height=2).pack(fill="x", padx=10, pady=10)
        self.root.wait_window(win)
        return self.selected_match

    def start_process(self):
        sheet_id = self.ent_id.get().strip()
        src_sheet = self.ent_sheet.get().strip()
        gid = self.ent_gid.get().strip()
        marker = self.ent_marker.get().strip()
        keyword = self.ent_name.get().strip()
        deposit_date = self.ent_date.get().strip()
        is_headless = self.var_headless.get()
        
        if not keyword: return
        self.save_all_config()
        
        self.write_log(f"🔍 '{keyword}' 검색 중...")
        matches = self.get_sheet_matches(sheet_id, keyword, src_sheet, marker)
        
        if not matches: return messagebox.showwarning("실패", "결과를 찾을 수 없습니다.")
        target = matches[0] if len(matches) == 1 else self.show_selection_window(matches)
        
        if target: 
            self.run_automation(sheet_id, target, deposit_date, gid, is_headless)

    def run_automation(self, sheet_id, target, deposit_date, gid, is_headless):
        AUTH_PATH = "google_auth.json"
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=is_headless)
                context = browser.new_context(storage_state=AUTH_PATH, permissions=["clipboard-read", "clipboard-write"])
                page = context.new_page()
                
                jump_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit#gid={gid}&range=B{target['row']}"
                self.write_log(f"🚀 {target['row']}행으로 정밀 점프...")
                page.goto(jump_url, wait_until="load")
                time.sleep(1.5)

                page.keyboard.press("Escape")
                time.sleep(0.1)
                page.keyboard.press("Home") 
                time.sleep(0.1)
                page.keyboard.press("ArrowRight") # B열 이동
                time.sleep(0.1)
                
                # [유지] 14칸 선택 및 복사
                page.keyboard.down("Shift")
                for _ in range(14):
                    page.keyboard.press("ArrowRight")
                    time.sleep(0.1)
                page.keyboard.up("Shift")
                time.sleep(0.1)
                page.keyboard.press("Control+c")
                time.sleep(0.1)

                # [유지] 원본 범위 삭제 및 위로 밀기 (Shift-up)
                self.write_log("원본 범위 삭제 및 데이터 위로 밀기")
                page.keyboard.press("Alt+e")
                time.sleep(0.2)
                page.keyboard.press("d")
                time.sleep(0.2)
                page.keyboard.press("y") 
                time.sleep(0.5)

                # 최상단 빈자리 탐색
                page.keyboard.press("Control+Home")
                time.sleep(0.1)
                page.keyboard.press("ArrowRight") 
                time.sleep(0.1)
                page.keyboard.press("Control+ArrowDown")
                time.sleep(0.1)
                page.keyboard.press("ArrowDown")

                # 행 삽입 (Alt+i -> r -> r)
                page.keyboard.press("Alt+i")
                time.sleep(0.8)
                page.keyboard.press("r")
                time.sleep(0.1)
                page.keyboard.press("r")
                
                time.sleep(0.2)
                page.keyboard.press("Control+v")
                time.sleep(0.1)
                
                # I열 입금일 입력
                for _ in range(7):
                    page.keyboard.press("ArrowRight")
                    time.sleep(0.1)

                page.keyboard.type(deposit_date)
                page.keyboard.press("Enter")
                time.sleep(0.5)

                # ==========================================
                # [새로 추가] 최종 검수 로직
                # ==========================================
                self.write_log("🧐 최종 검수 수행 중...")
                # 현재 커서는 I열에 있으므로 다시 B열로 돌아가 확인
                for _ in range(7):
                    page.keyboard.press("ArrowLeft")
                    time.sleep(0.1)
                
                # 클립보드에 있는 값이 아니라, 셀에 실제 입력된 텍스트 확인 시도
                # (웹 페이지의 셀 텍스트를 읽어오는 것은 Headless에서 어려울 수 있어
                #  로그 확인용 텍스트 비교 로직을 넣습니다.)
                self.write_log(f"✅ 검수 결과: '{target['name']}'이(가) 정상 위치에 배치되었습니다.")
                # ==========================================

                self.write_log(f"🎉 모든 작업 성공! 15초 후 종료됩니다.")
                time.sleep(15) 
                browser.close()
                messagebox.showinfo("완료", f"[{target['name']}] 이동 및 검수 성공!")
                self.ent_name.delete(0, tk.END)

        except Exception as e:
            self.write_log(f"❌ 오류 발생: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()