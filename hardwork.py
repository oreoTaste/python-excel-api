# pip install openpyxl requests python-dotenv gspread google-auth
# pyinstaller --onefile --noconsole main.py
import os
import json
import gspread
from google.oauth2.service_account import Credentials
import requests
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from dotenv import load_dotenv
import re
import datetime

# --- [1. 설정 및 상수 관리] ---
load_dotenv()
CONFIG_FILE = "settings.json"
SERVICE_ACCOUNT_FILE = "credentials.json"
TARGET_MARKER = "▼입금 대기"  # 검색 및 삽입 위치 기준 마커

DEFAULT_CONFIG = {
    "spreadsheet_id": "1Q7Wew2MtwwYh0aSam2XNvBrwFHxM_3Kb2CT-Qv00-7o", # 지정된 시트 ID
    "src_sheet": "월별내역",
    "tgt_sheet": "월별내역",
    "sum_formula_cell": "M97, N97, O97, P97", # 다중 합계 셀 요구사항 반영
    "user_id": os.getenv("COMPANY_ID", "your_id"),
    "api_login_url": os.getenv("LOGIN_URL", "https://company.com/api/login"),
    "api_save_url": os.getenv("SAVE_URL", "https://company.com/api/save")
}

def load_settings():
    """기존 설정과 기본 설정을 병합하여 로드합니다."""
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            try:
                user_config = json.load(f)
                config.update(user_config)
            except Exception:
                pass
    return config

def save_settings(config):
    """현재 설정을 JSON 파일로 저장합니다."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# --- [2. 메인 애플리케이션 클래스] ---

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("구글 시트 업무 자동화 도구 v3.5")
        self.root.geometry("550x750")
        
        self.config = load_settings()
        self.entries = {}
        self.gc = None
        
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("Malgun Gothic", 10))
        self.style.configure("TButton", font=("Malgun Gothic", 10))
        
        self.setup_ui()
        if "name" in self.entries:
            self.entries["name"].focus_set()
        self.root.bind('<Return>', lambda event: self.start_process())

    def setup_ui(self):
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill="both", expand=True)

        # 1. 시트 설정 영역
        frame = ttk.LabelFrame(main_container, text=" 구글 스프레드시트 설정 ", padding="10")
        frame.pack(fill="x", pady=(0, 10))

        ttk.Label(frame, text="시트 ID:").grid(row=0, column=0, sticky="e", pady=2)
        self.ent_sheet_id = ttk.Entry(frame, width=43)
        self.ent_sheet_id.insert(0, self.config.get("spreadsheet_id", ""))
        self.ent_sheet_id.grid(row=0, column=1, padx=5, sticky="w")

        self.ent_src = self._add_config_row(frame, "소스 시트명:", self.config.get("src_sheet"), 1)
        self.ent_tgt = self._add_config_row(frame, "타겟 시트명:", self.config.get("tgt_sheet"), 2)
        self.ent_sum_cell = self._add_config_row(frame, "합계 수식 셀들:", self.config.get("sum_formula_cell"), 3)
        
        # 2. 데이터 입력 영역
        self._create_input_ui(main_container)

        self.btn_run = tk.Button(
            main_container, text="구글 시트 작업 실행 (Enter)", command=self.start_process, 
            bg="#4285F4", fg="white", font=("Malgun Gothic", 12, "bold"), 
            height=2, relief="flat", cursor="hand2"
        )
        self.btn_run.pack(pady=15, fill="x")

        # 3. 로그 영역
        self._create_log_ui(main_container)

    def _add_config_row(self, frame, label, value, row):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="e", pady=5)
        entry = ttk.Entry(frame, width=43)
        entry.insert(0, value)
        entry.grid(row=row, column=1, padx=5, sticky="w")
        return entry

    def _create_input_ui(self, parent):
        frame = ttk.LabelFrame(parent, text=" 작업 데이터 입력 ", padding="10")
        frame.pack(fill="x", pady=5)
        fields = [
            ("업체명(검색)", "name", "(대소문자 무관)"),
            ("계약금액", "amount", "(숫자만)"),
            ("계약날짜", "cdate", "(YYYY-MM-DD)"),
            ("시작날짜", "sdate", "(YYYY-MM-DD)")
        ]
        for i, (label, key, guide) in enumerate(fields):
            ttk.Label(frame, text=f"{label}:").grid(row=i, column=0, sticky="e", pady=5)
            ent = ttk.Entry(frame, width=28)
            ent.grid(row=i, column=1, padx=5, sticky="w")
            self.entries[key] = ent
            ttk.Label(frame, text=guide, foreground="#7f8c8d", font=("Malgun Gothic", 9)).grid(row=i, column=2, sticky="w")

    def _create_log_ui(self, parent):
        frame = ttk.LabelFrame(parent, text=" 처리 로그 ", padding="5")
        frame.pack(fill="both", expand=True)
        self.log_area = scrolledtext.ScrolledText(frame, height=10, font=("Consolas", 9), state='disabled', bg="#f8f9fa")
        self.log_area.pack(fill="both", expand=True)
        self.write_log("시스템 준비 완료.")

    def write_log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def save_current_settings(self):
        self.config.update({
            "spreadsheet_id": self.ent_sheet_id.get(),
            "src_sheet": self.ent_src.get(),
            "tgt_sheet": self.ent_tgt.get(),
            "sum_formula_cell": self.ent_sum_cell.get()
        })
        save_settings(self.config)

    def authenticate_gspread(self):
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(f"인증 파일('{SERVICE_ACCOUNT_FILE}')이 필요합니다.")
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        return gspread.authorize(creds)

    def start_process(self):
        """작업리스트 마커 하단 검색 로직 수행"""
        if not messagebox.askyesno("실행 확인", "데이터를 검색하시겠습니까?"):
            return

        self.save_current_settings()
        keyword = self.entries["name"].get().strip()
        if not keyword:
            messagebox.showwarning("알림", "업체명을 입력하세요.")
            return

        try:
            if not self.gc: self.gc = self.authenticate_gspread()
            sh = self.gc.open_by_key(self.config["spreadsheet_id"])
            ws = sh.worksheet(self.config["src_sheet"])
            all_values = ws.get_all_values()
            
            # 마커 위치 찾기 (A열 또는 B열)
            marker_idx = -1
            for i, row in enumerate(all_values):
                if (len(row) > 0 and row[0] == TARGET_MARKER) or (len(row) > 1 and row[1] == TARGET_MARKER):
                    marker_idx = i
                    break
            
            start_search_idx = marker_idx + 1 if marker_idx != -1 else 1
            self.write_log(f"'{TARGET_MARKER}' 기준 {start_search_idx + 1}행부터 검색합니다.")

            matches = []
            for i, row in enumerate(all_values[start_search_idx:], start=start_search_idx + 1):
                if row and keyword.lower() in str(row[0]).lower():
                    matches.append({"idx": i, "data": row})

            if not matches:
                self.write_log(f"결과 없음: '{keyword}'")
                messagebox.showinfo("결과", "데이터를 찾지 못했습니다.")
            elif len(matches) == 1:
                self.confirm_and_run(matches[0])
            else:
                self.show_selection_window(matches)
        except Exception as e:
            self.write_log(f"오류: {str(e)}")
            messagebox.showerror("오류", str(e))

    def confirm_and_run(self, match):
        if not messagebox.askyesno("최종 확인", f"업체: '{match['data'][0]}'\n이동하시겠습니까?"):
            return
        try:
            self.write_log("구글 시트 업데이트 중...")
            new_sum_config = self._execute_gsheet_update(match)
            
            # UI 및 설정에 새 수식 위치 반영
            self.ent_sum_cell.delete(0, tk.END)
            self.ent_sum_cell.insert(0, new_sum_config)
            self.save_current_settings()
            
            self.write_log("완료: 데이터 이동 및 합계 갱신 성공")
            messagebox.showinfo("성공", "작업이 완료되었습니다!")
            self.entries["name"].delete(0, tk.END)
        except Exception as e:
            self.write_log(f"실행 오류: {str(e)}")
            messagebox.showerror("오류", str(e))

    def _execute_gsheet_update(self, match):
        """데이터 이동(A-P) 및 다중 합계 수식 일괄 업데이트 로직"""
        sh = self.gc.open_by_key(self.config["spreadsheet_id"])
        src_ws = sh.worksheet(self.config["src_sheet"])
        tgt_ws = sh.worksheet(self.config["tgt_sheet"])
        
        row_idx = match['idx']
        # 1. A-P열 데이터만 추출 및 원본 비우기
        data_to_move = (match['data'] + [""] * 16)[:16]
        src_ws.batch_clear([f"A{row_idx}:P{row_idx}"])
        
        # 2. 완료리스트 삽입 위치 찾기
        tgt_values = tgt_ws.get_all_values()
        paste_row = len(tgt_values) + 1
        for i, row in enumerate(tgt_values):
            if len(row) > 1 and TARGET_MARKER in row[1]:
                for j in range(i + 1, len(tgt_values)):
                    if not "".join(tgt_values[j]).strip():
                        paste_row = j + 1
                        break
                break
        
        tgt_ws.insert_row(data_to_move, index=paste_row)

        # 3. 다중 합계 수식 업데이트 (M, N, O, P열 등)
        updated_tgt_values = tgt_ws.get_all_values()
        sum_row_idx = -1
        for i in range(paste_row, len(updated_tgt_values)):
            if any("합계" in str(v) for v in updated_tgt_values[i]):
                sum_row_idx = i + 1
                break
        
        if sum_row_idx != -1:
            # 설정에서 열 문자 추출 (예: M, N, O, P)
            target_cols = re.findall(r'([A-Z]+)', self.config["sum_formula_cell"])
            updates = []
            new_cells = []
            for col in target_cols:
                formula = f"=SUM({col}2:{col}{sum_row_idx - 1})"
                updates.append({'range': f"{col}{sum_row_idx}", 'values': [[formula]]})
                new_cells.append(f"{col}{sum_row_idx}")
            
            tgt_ws.batch_update(updates, value_input_option='USER_ENTERED')
            return ", ".join(new_cells) # 갱신된 셀 주소 문자열 반환
        
        return self.config["sum_formula_cell"]

    def show_selection_window(self, matches):
        win = tk.Toplevel(self.root)
        win.title("업체 선택")
        lb = tk.Listbox(win, font=("Malgun Gothic", 10), width=50, height=10)
        lb.pack(padx=10, pady=5)
        for m in matches:
            lb.insert(tk.END, f"{m['data'][0]} (행: {m['idx']})")
        def on_select():
            if lb.curselection():
                idx = lb.curselection()[0]
                win.destroy()
                self.confirm_and_run(matches[idx])
        ttk.Button(win, text="선택", command=on_select).pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationApp(root)
    root.mainloop()