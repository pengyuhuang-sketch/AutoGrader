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
        self.title("AI 雲端多生閱卷系統 v6.0")
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
        
        ctk.CTkLabel(config_frame, text="API Key 狀態:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.api_status_label = ctk.CTkLabel(config_frame, text="尚未載入", text_color="red")
        self.api_status_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(config_frame, text="手動更新 Key:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.api_entry = ctk.CTkEntry(config_frame, width=350, show="*")
        self.api_entry.grid(row=1, column=1, padx=10, pady=5)
        
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key_to_file).grid(row=1, column=2, padx=10)

        ctk.CTkLabel(config_frame, text="欲批改總人數:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.student_count_entry = ctk.CTkEntry(config_frame, width=100)
        self.student_count_entry.insert(0, "1")
        self.student_count_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        # 2. 功能按鈕
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word)", command=self.load_word).grid(row=0, column=0, padx=10, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF)", command=self.start_grading, fg_color="#2ecc71")
        self.btn_start.grid(row=0, column=1, padx=10, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出 Excel", command=self.export_excel, state="disabled")
        self.btn_export.grid(row=0, column=2, padx=10, pady=10)

        # 3. 成績表格
        self.tree = ttk.Treeview(self, columns=("Student", "Score", "Summary"), show='headings')
        self.tree.heading("Student", text="學生姓名/座號")
        self.tree.heading("Score", text="得分")
        self.tree.heading("Summary", text="評語")
        self.tree.pack(pady=10, padx=20, fill="both", expand=True)

        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var).pack(pady=5)

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                key = f.read().strip()
                if key:
                    self.api_entry.insert(0, key)
                    self.api_status_label.configure(text="已載入 api_key.txt", text_color="green")

    def save_api_key_to_file(self):
        key = self.api_entry.get().strip()
        with open(self.config_file, "w") as f: f.write(key)
        messagebox.showinfo("成功", "API Key 已儲存")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if path:
            doc = Document(path)
            self.answer_text = "\n".join([" | ".join([c.text.strip() for c in r.cells]) for t in doc.tables for r in t.rows])
            self.status_var.set("解答載入成功")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if pdf_path and api_key:
            for item in self.tree.get_children(): self.tree.delete(item)
            threading.Thread(target=self.process_gemini, args=(pdf_path, api_key), daemon=True).start()

    def process_gemini(self, pdf_path, api_key):
        try:
            self.status_var.set("連線至 AI 中...")
            genai.configure(api_key=api_key)
            # 使用最通用的模型名稱格式
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            uploaded_file = genai.upload_file(path=pdf_path)
            student_count = self.student_count_entry.get()
            
            prompt = f"參考解答：\n{self.answer_text}\n這份PDF有{student_count}位學生。請批改並回傳JSON列表：[{{'student': '姓名', 'score': '分數', 'summary': '評語'}}]"
            
            response = model.generate_content([prompt, uploaded_file])
            json_match = re.search(r'\[.*\]', response.text, re.DOTALL)
            
            if json_match:
                data = json.loads(json_match.group())
                for s in data:
                    self.tree.insert("", "end", values=(s['student'], s['score'], s['summary']))
                    self.results_data.append(s)
                self.status_var.set("批改完成")
                self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("錯誤", f"請檢查 API Key 或網路環境\n{str(e)}")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path: pd.DataFrame(self.results_data).to_excel(path, index=False)

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
