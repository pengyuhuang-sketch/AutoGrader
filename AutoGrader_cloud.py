import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import fitz  # PyMuPDF
import google.generativeai as genai
import pandas as pd
import os
import threading
import json
import time
import io

class AutoGraderCloud(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 閱卷系統 v8.8 - 座標對位與空白鎖定版")
        self.geometry("1100x850")
        
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        ctk.CTkLabel(self, text="AI 閱卷系統 - 終極穩定辨識模式", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        # 1. API 配置
        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(config_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_entry = ctk.CTkEntry(config_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key).grid(row=0, column=2, padx=10)

        # 2. 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word/PDF)", command=self.load_answer, width=200).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF考卷)", command=self.start_grading, fg_color="#2ecc71", width=180)
        self.btn_start.grid(row=0, column=1, padx=15, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出分頁 Excel", command=self.export_excel, state="disabled", width=180)
        self.btn_export.grid(row=0, column=2, padx=15, pady=10)

        # 3. 數據表格
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, columns=("Class", "No", "Name", "CorrectCount"), show='headings')
        self.tree.heading("Class", text="班級"); self.tree.heading("No", text="座號")
        self.tree.heading("Name", text="姓名"); self.tree.heading("CorrectCount", text="正確題數")
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var).pack(pady=5)

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f: self.api_entry.insert(0, f.read().strip())

    def save_api_key(self):
        with open(self.config_file, "w") as f: f.write(self.api_entry.get())
        messagebox.showinfo("成功", "Key 已儲存")

    def load_answer(self):
        path = filedialog.askopenfilename(filetypes=[("解答檔案", "*.docx *.pdf")])
        if not path: return
        try:
            file_ext = os.path.splitext(path)[1].lower()
            if file_ext == ".docx":
                with open(path, 'rb') as f:
                    doc = Document(io.BytesIO(f.read()))
                    self.answer_text = "\n".join([" | ".join([c.text.strip() for c in r.cells]) for t in doc.tables for r in t.rows])
            elif file_ext == ".pdf":
                doc = fitz.open(path)
                self.answer_text = "".join([page.get_text() for page in doc])
                doc.close()
            self.status_var.set("解答載入完成")
        except Exception as e:
            messagebox.showerror("錯誤", f"解答讀取失敗: {str(e)}")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if pdf_path and api_key:
            if not self.answer_text:
                messagebox.showwarning("提示", "請先載入正確解答")
                return
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("偵測 API 模型...")
            genai.configure(api_key=api_key)
            available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in available_models if "gemini-1.5-flash" in m), None) or "models/gemini-1.5-flash"

            model = genai.GenerativeModel(model_name=target_model, generation_config={"response_mime_type": "application/json"})
            
            uploaded_file = genai.upload_file(path=pdf_path)
            while uploaded_file.state.name == "PROCESSING": 
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
            
            # --- v8.8 核心修正：空白校正與嚴格座標對位 ---
            prompt = f"""
            你是一位視覺閱卷專家。影像中是學生的手寫答案卷。
            參考解答：{self.answer_text}

            【辨識指令：座標與空白鎖定】
            1. **嚴格座標對位**：
               - 每一題的辨識必須以「印刷題號」(如 1., 2., 3.) 為基準點。
               - 請確認第1題的位置，絕對不可跳過第一題。若第1題無答案，請回傳 ""，不可將第2題的答案往前填補。
            2. **空白判定基準 (核心修正)**：
               - **判定為空白 ""**：如果該題號下方的格子內沒有明顯的「深色人為墨跡」，請務必回傳空字串 ""。
               - **嚴禁腦補**：絕對不可將掃描產生的微小灰塵點、紙張纖維、或印刷的灰色格線誤判為 A, B, C。
               - 若你無法確定是不是字，只要它不像連續的筆畫，就判定為空白。
            3. **字母結構保護**：
               - 'A'：必須有交會頂點與橫線。若只有斜線無橫線，極可能是噪音或誤判，請謹慎。
               - 'C'：必須有明顯右側開口。
            4. **禁止位移**：確保 JSON 中的 q_idx 與影像中的印刷題號完全一一對應。

            嚴格回傳 JSON 格式：
            [ {{"class": "班級", "no": "座號", "name": "姓名", "questions": [{{"q_idx": 1, "s_ans": "A", "c_ans": "A", "res": "○"}}] }} ]
            """
            
            self.status_var.set("AI 進行精準座標辨識中...")
            response = model.generate_content([prompt, uploaded_file])
            self.results_data = json.loads(response.text)

            for s in self.results_data:
                correct_count = sum(1 for q in s.get('questions', []) if q.get('res') == '○')
                s['correct_sum'] = correct_count
                self.tree.insert("", "end", values=(s.get('class'), s.get('no'), s.get('name'), correct_count))
                
            self.status_var.set("批改完成！已校正位移與空白誤判。")
            self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("錯誤", f"辨識失敗: {str(e)}")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try:
            # 學生作答分頁
            export_dict = {"項目": ["班級", "座號", "姓名", "正確總題數"]}
            max_q = max([len(s.get("questions", [])) for s in self.results_data]) if self.results_data else 0
            for i in range(1, max_q + 1):
                export_dict["項目"].append(f"第{i}題_學生答案")
                export_dict["項目"].append(f"第{i}題_結果")

            for s in self.results_data:
                col_name = f"{s.get('name')}({s.get('no')})"
                vals = [s.get("class"), s.get("no"), s.get("name"), s.get("correct_sum")]
                for q in s.get("questions", []):
                    vals.extend([q.get("s_ans"), q.get("res")])
                while len(vals) < len(export_dict["項目"]): vals.append("")
                export_dict[col_name] = vals

            # 正確解答分頁
            ans_dict = {"題號": [], "標準解答": []}
            if self.results_data:
                for q in self.results_data[0].get("questions", []):
                    ans_dict["題號"].append(f"第{q.get('q_idx')}題")
                    ans_dict["標準解答"].append(q.get("c_ans"))

            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                pd.DataFrame(export_dict).to_excel(writer, sheet_name='學生作答成績', index=False)
                pd.DataFrame(ans_dict).to_excel(writer, sheet_name='正確解答', index=False)

            messagebox.showinfo("成功", "分頁報告匯出成功！")
        except Exception as e:
            messagebox.showerror("匯出錯誤", str(e))

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
