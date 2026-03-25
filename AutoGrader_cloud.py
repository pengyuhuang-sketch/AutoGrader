import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import google.generativeai as genai
import pandas as pd
import os
import threading
import json
import time

class AutoGraderCloud(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 雲端多生閱卷系統 v7.1 (穩定強化版)")
        self.geometry("900x850")
        
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        ctk.CTkLabel(self, text="AI 雲端自動閱卷系統 (Gemini 3 Flash)", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

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

        # 3. 表格區域 (帶滾動條)
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.tree = ttk.Treeview(table_frame, columns=("Student", "Score", "Summary"), show='headings')
        self.tree.heading("Student", text="學生姓名/座號")
        self.tree.heading("Score", text="得分")
        self.tree.heading("Summary", text="AI 評語")
        self.tree.column("Student", width=150, anchor="center")
        self.tree.column("Score", width=100, anchor="center")
        self.tree.column("Summary", width=400, anchor="w")
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 狀態列
        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var, font=("Microsoft JhengHei", 12)).pack(pady=5)

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    self.api_entry.insert(0, f.read().strip())
            except: pass

    def save_api_key(self):
        with open(self.config_file, "w") as f:
            f.write(self.api_entry.get())
        messagebox.showinfo("成功", "API Key 已儲存至本地檔案")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if path:
            try:
                doc = Document(path)
                # 抓取所有表格中的文字作為解答
                content = []
                for table in doc.tables:
                    for row in table.rows:
                        row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        if row_text:
                            content.append(" | ".join(row_text))
                self.answer_text = "\n".join(content)
                self.status_var.set(f"解答載入完成：已從 {os.path.basename(path)} 提取數據")
            except Exception as e:
                messagebox.showerror("載入失敗", f"Word 讀取錯誤: {str(e)}")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showwarning("提示", "請輸入 API Key")
            return
        if not self.answer_text:
            messagebox.showwarning("提示", "請先載入解答 Word 檔")
            return
            
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF 考卷", "*.pdf")])
        if pdf_path:
            # 清空舊資料
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.results_data = []
            self.btn_start.configure(state="disabled")
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("正在連線 Google Gemini 雲端...")
            genai.configure(api_key=api_key)
            
            # 使用 JSON 模式，確保 AI 回傳格式穩定
            model = genai.GenerativeModel(
                model_name="gemini-1.5-flash",
                generation_config={"response_mime_type": "application/json"}
            )
            
            self.status_var.set("上傳 PDF 檔案中...")
            uploaded_file = genai.upload_file(path=pdf_path)
            
            # 檢查檔案處理狀態
            while uploaded_file.state.name == "PROCESSING":
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
            
            if uploaded_file.state.name == "FAILED":
                raise Exception("Google 雲端處理 PDF 失敗，請重試。")

            prompt = f"""
            你是一位嚴謹的批改老師。請分析 PDF 中包含的約 {self.student_count_entry.get()} 位學生的答題內容。
            
            標準解答與配分參考：
            {self.answer_text}
            
            任務：
            1. 辨識每位學生的姓名或座號。
            2. 比對答案並計算總分。
            3. 為每位學生提供一句簡短的評語。
            4. 嚴格以 JSON 格式回傳，格式範例：
               [ {{"student": "王小明", "score": 85, "summary": "選擇題表現優異，計算題需注意單位。"}} ]
            """
            
            self.status_var.set("AI 影像辨識與批改中 (這可能需要 20-40 秒)...")
            response = model.generate_content([prompt, uploaded_file])
            
            # 解析 JSON 數據
            data = json.loads(response.text)
            
            if not isinstance(data, list):
                # 有些時候 AI 可能會包一層 Key
                if isinstance(data, dict) and len(data) == 1:
                    data = list(data.values())[0]

            for s in data:
                name = s.get('student', '未知')
                score = s.get('score', 'N/A')
                summary = s.get('summary', '無評語')
                self.tree.insert("", "end", values=(name, score, summary))
                self.results_data.append({"學生": name, "分數": score, "評語": summary})
                
            self.status_var.set(f"批改完成！共處理 {len(data)} 位學生。")
            self.btn_export.configure(state="normal")
            
        except Exception as e:
            messagebox.showerror("執行錯誤", f"詳細錯誤：{str(e)}")
            self.status_var.set("發生錯誤，請檢查網路或 API Key")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        if not self.results_data: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 檔案", "*.xlsx")])
        if path:
            try:
                df = pd.DataFrame(self.results_data)
                df.to_excel(path, index=False)
                messagebox.showinfo("匯出成功", f"報告已儲存至：{path}")
            except Exception as e:
                messagebox.showerror("匯出失敗", str(e))

if __name__ == "__main__":
    app = AutoGraderCloud()
    app.mainloop()
