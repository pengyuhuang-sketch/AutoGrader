import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
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
        self.title("AI 雲端多生閱卷系統 v7.3 (詳盡比對版)")
        self.geometry("1100x850") # 稍微加寬以容納更多欄位
        
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        ctk.CTkLabel(self, text="AI 雲端自動閱卷系統 (Detailed Analysis)", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        # 1. API 配置區
        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(config_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_entry = ctk.CTkEntry(config_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(config_frame, text="預計總人數:").grid(row=1, column=0, padx=10, pady=5)
        self.student_count_entry = ctk.CTkEntry(config_frame, width=100)
        self.student_count_entry.insert(0, "1")
        self.student_count_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 2. 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word)", command=self.load_word, width=180).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF)", command=self.start_grading, fg_color="#2ecc71", hover_color="#27ae60", width=180)
        self.btn_start.grid(row=0, column=1, padx=15, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出 Excel", command=self.export_excel, state="disabled", width=180)
        self.btn_export.grid(row=0, column=2, padx=15, pady=10)

        # 3. 表格區域 - 新增 班級、座號、逐題詳情 欄位
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        columns = ("Class", "No", "Name", "Details", "Score")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        self.tree.heading("Class", text="班級")
        self.tree.heading("No", text="座號")
        self.tree.heading("Name", text="姓名")
        self.tree.heading("Details", text="逐題比對結果 (學生答案 vs 正解)")
        self.tree.heading("Score", text="總分")
        
        self.tree.column("Class", width=70, anchor="center")
        self.tree.column("No", width=60, anchor="center")
        self.tree.column("Name", width=100, anchor="center")
        self.tree.column("Details", width=550, anchor="w")
        self.tree.column("Score", width=80, anchor="center")
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var, font=("Microsoft JhengHei", 12)).pack(pady=5)

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                self.api_entry.insert(0, f.read().strip())

    def save_api_key(self):
        with open(self.config_file, "w") as f:
            f.write(self.api_entry.get())
        messagebox.showinfo("成功", "API Key 已儲存")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if path:
            try:
                with open(path, 'rb') as f:
                    file_stream = io.BytesIO(f.read())
                    doc = Document(file_stream)
                    content = []
                    for table in doc.tables:
                        for row in table.rows:
                            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                            if row_text:
                                content.append(" | ".join(row_text))
                    self.answer_text = "\n".join(content)
                self.status_var.set(f"解答載入成功：{os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("載入失敗", f"Word 讀取錯誤: {str(e)}")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF 考卷", "*.pdf")])
        if pdf_path and api_key:
            if not self.answer_text:
                messagebox.showwarning("提示", "請先載入解答 Word")
                return
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            self.btn_start.configure(state="disabled")
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("正在動態偵測模型...")
            genai.configure(api_key=api_key)
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in models if "1.5-flash" in m), models[0] if models else "")
            
            model = genai.GenerativeModel(
                model_name=target_model,
                generation_config={"response_mime_type": "application/json"}
            )
            
            self.status_var.set("上傳並分析 PDF 中...")
            uploaded_file = genai.upload_file(path=pdf_path)
            while uploaded_file.state.name == "PROCESSING":
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
            
            # 關鍵：更新 Prompt 要求詳盡比對資料
            prompt = f"""
            你是一位專業教師。PDF 包含約 {self.student_count_entry.get()} 位學生的考卷影像。
            標準解答：
            {self.answer_text}
            
            任務：
            1. 辨識每位學生的「班級」、「座號」與「姓名」。
            2. 針對每一題，列出「學生答案」與「正確答案」的比對結果。
            3. 計算總分。
            4. 嚴格回傳 JSON 格式：
            [
              {{
                "class": "班級名稱",
                "no": "座號",
                "student": "姓名",
                "details": "1. A(正解A) ✅; 2. B(正解C) ❌; ...",
                "score": 總分
              }}
            ]
            """
            
            response = model.generate_content([prompt, uploaded_file])
            raw_data = json.loads(response.text)
            data = raw_data if isinstance(raw_data, list) else next(iter(raw_data.values()))

            for s in data:
                c_name = s.get('class', '未知')
                no = s.get('no', '未知')
                name = s.get('student', '未知')
                details = s.get('details', '無數據')
                score = s.get('score', 0)
                
                self.tree.insert("", "end", values=(c_name, no, name, details, score))
                # 儲存到結果清單供匯出
                self.results_data.append({
                    "班級": c_name,
                    "座號": no,
                    "姓名": name,
                    "每題比對詳情": details,
                    "總計分數": score
                })
                
            self.status_var.set("批改完成！")
            self.btn_export.configure(state="normal")
            
        except Exception as e:
            messagebox.showerror("失敗", f"錯誤：{str(e)}")
            self.status_var.set("發生錯誤")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            try:
                df = pd.DataFrame(self.results_data)
                # 調整 Excel 欄位順序
                df = df[["班級", "座號", "姓名", "每題比對詳情", "總計分數"]]
                df.to_excel(path, index=False)
                messagebox.showinfo("成功", f"檔案已匯出至：{path}")
            except Exception as e:
                messagebox.showerror("匯出錯誤", str(e))

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
