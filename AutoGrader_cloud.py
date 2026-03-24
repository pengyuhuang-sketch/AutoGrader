import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import google.generativeai as genai
import pandas as pd
import os
import threading
import json
import re

class AutoGraderCloud(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 雲端多生閱卷系統 v7.0 (穩定版)")
        self.geometry("900x850")
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.CTkLabel(self, text="AI 雲端自動閱卷系統", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        # 1. API 配置區
        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(config_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_entry = ctk.CTkEntry(config_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkButton(config_frame, text="儲存", width=80, command=self.save_api_key).grid(row=0, column=2, padx=10)

        ctk.CTkLabel(config_frame, text="PDF 內預計總人數:").grid(row=1, column=0, padx=10, pady=5)
        self.student_count_entry = ctk.CTkEntry(config_frame, width=100)
        self.student_count_entry.insert(0, "1")
        self.student_count_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # 2. 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word)", command=self.load_word).grid(row=0, column=0, padx=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF)", command=self.start_grading, fg_color="#2ecc71")
        self.btn_start.grid(row=0, column=1, padx=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出 Excel", command=self.export_excel, state="disabled")
        self.btn_export.grid(row=0, column=2, padx=10)

        # 3. 表格
        self.tree = ttk.Treeview(self, columns=("Student", "Score", "Status"), show='headings')
        self.tree.heading("Student", text="學生姓名/座號")
        self.tree.heading("Score", text="得分")
        self.tree.heading("Status", text="狀態")
        self.tree.pack(pady=10, padx=20, fill="both", expand=True)

        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var).pack()

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                self.api_entry.insert(0, f.read().strip())

    def save_api_key(self):
        with open(self.config_file, "w") as f: f.write(self.api_entry.get())
        messagebox.showinfo("成功", "Key 已儲存")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if path:
            doc = Document(path)
            self.answer_text = "\n".join([" | ".join([c.text.strip() for c in r.cells]) for t in doc.tables for r in t.rows])
            self.status_var.set("解答載入完成")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if pdf_path and api_key:
            for item in self.tree.get_children(): self.tree.delete(item)
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("連線 Google AI 中...")
            genai.configure(api_key=api_key)
            
            # 【關鍵修正】明確指定使用穩定版 flash
            model = genai.GenerativeModel(model_name="gemini-1.5-flash")
            
            self.status_var.set("上傳 PDF 中...")
            uploaded_file = genai.upload_file(path=pdf_path)
            
            prompt = f"""
            你是一位專業教師。PDF 包含約 {self.student_count_entry.get()} 位學生的答案卷。
            參考解答：{self.answer_text}
            請辨識每人姓名並批改，僅回傳 JSON 列表：
            [ {{"student": "姓名", "score": "分數", "summary": "評語"}} ]
            """
            
            response = model.generate_content([prompt, uploaded_file])
            json_match = re.search(r'\[.*\]', response.text, re.DOTALL)
            
            if json_match:
                data = json.loads(json_match.group())
                for s in data:
                    self.tree.insert("", "end", values=(s['student'], s['score'], "完成"))
                    self.results_data.append(s)
                self.status_var.set("批改完成！")
                self.btn_export.configure(state="normal")
            else:
                self.status_var.set("AI 回傳格式不正確")
        except Exception as e:
            messagebox.showerror("批改失敗", f"錯誤訊息：{str(e)}")
            self.status_var.set("發生錯誤")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path: pd.DataFrame(self.results_data).to_excel(path, index=False)

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
