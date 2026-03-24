import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import google.generativeai as genai
import pandas as pd
import os
import threading
import json
import re
import sys

class AutoGraderCloud(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 雲端多生閱卷系統 v6.0")
        self.geometry("900(850")
        
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt" # 設定檔名稱
        self.setup_ui()
        self.load_api_key_from_file() # 啟動時讀取檔案

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        
        ctk.CTkLabel(self, text="AI 雲端自動閱卷系統 (API 檔案讀取版)", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        # 1. API Key 與 人數設定
        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(config_frame, text="API Key 狀態:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.api_status_label = ctk.CTkLabel(config_frame, text="尚未載入", text_color="red")
        self.api_status_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(config_frame, text="手動更新 Key:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.api_entry = ctk.CTkEntry(config_frame, width=350, show="*", placeholder_text="若需更換請在此輸入並按儲存")
        self.api_entry.grid(row=1, column=1, padx=10, pady=5)
        
        self.btn_save_key = ctk.CTkButton(config_frame, text="儲存 Key 至檔案", width=100, command=self.save_api_key_to_file)
        self.btn_save_key.grid(row=1, column=2, padx=10, pady=5)

        ctk.CTkLabel(config_frame, text="PDF 內的學生總數:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.student_count_entry = ctk.CTkEntry(config_frame, width=100)
        self.student_count_entry.insert(0, "1")
        self.student_count_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        # 2. 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        
        self.btn_load_ans = ctk.CTkButton(btn_frame, text="1. 載入 Word 解答", command=self.load_word)
        self.btn_load_ans.grid(row=0, column=0, padx=10, pady=10)

        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 PDF", command=self.start_grading, fg_color="#2ecc71")
        self.btn_start.grid(row=0, column=1, padx=10, pady=10)

        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出 Excel 成績單", command=self.export_excel, state="disabled")
        self.btn_export.grid(row=0, column=2, padx=10, pady=10)

        # 3. 表格顯示區
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.tree = ttk.Treeview(table_frame, columns=("Student", "Score", "Summary"), show='headings')
        self.tree.heading("Student", text="學生姓名/座號")
        self.tree.heading("Score", text="得分")
        self.tree.heading("Summary", text="AI 評語")
        self.tree.column("Student", width=150)
        self.tree.column("Score", width=80)
        self.tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var).pack(pady=5)

    def load_api_key_from_file(self):
        """從 txt 檔案讀取 API Key"""
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f:
                key = f.read().strip()
                if key:
                    self.api_entry.delete(0, "end")
                    self.api_entry.insert(0, key)
                    self.api_status_label.configure(text="已從 api_key.txt 載入", text_color="green")
                    return key
        return ""

    def save_api_key_to_file(self):
        """將 API Key 儲存到 txt 檔案"""
        key = self.api_entry.get().strip()
        if key:
            with open(self.config_file, "w") as f:
                f.write(key)
            self.api_status_label.configure(text="Key 已儲存至檔案", text_color="green")
            messagebox.showinfo("成功", "API Key 已成功儲存至 api_key.txt")
        else:
            messagebox.showwarning("錯誤", "請先輸入 API Key 再儲存")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if path:
            doc = Document(path)
            full_text = []
            for table in doc.tables:
                for row in table.rows:
                    full_text.append(" | ".join([c.text.strip() for c in row.cells]))
            self.answer_text = "\n".join(full_text)
            self.status_var.set("解答載入成功")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showwarning("錯誤", "API Key 為空！請在 api_key.txt 填入 Key 或在欄位中輸入。")
            return
        if not self.answer_text:
            messagebox.showwarning("錯誤", "請先載入解答檔")
            return
        
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if pdf_path:
            self.results_data = []
            for item in self.tree.get_children(): self.tree.delete(item)
            threading.Thread(target=self.process_with_gemini, args=(pdf_path, api_key), daemon=True).start()

    def process_with_gemini(self, pdf_path, api_key):
        try:
            self.status_var.set("正在與 Google AI 進行連線...")
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash-latest')
            
            student_count = self.student_count_entry.get()
            uploaded_file = genai.upload_file(path=pdf_path)
            
            prompt = f"""
            你是一位專業教師。PDF 包含約 {student_count} 位學生的答案。
            參考解答：{self.answer_text}
            請辨識每位學生姓名、批改並回傳 JSON 列表格式（嚴禁回傳解釋文字）：
            [ {{"student": "姓名/座號", "score": "分數", "summary": "評語"}} ]
            """
            
            self.status_var.set("AI 正在批改多份考卷...")
            response = model.generate_content([prompt, uploaded_file])
            
            json_match = re.search(r'\[.*\]', response.text, re.DOTALL)
            if json_match:
                students = json.loads(json_match.group())
                for s in students:
                    name = s.get("student", "未知")
                    score = s.get("score", "0")
                    summary = s.get("summary", "")
                    self.results_data.append({"學生名稱": name, "得分": score, "評語": summary})
                    self.tree.insert("", "end", values=(name, score, summary))
                self.status_var.set(f"批改完成！共 {len(students)} 位學生")
                self.btn_export.configure(state="normal")
            else:
                self.status_var.set("AI 回傳格式不正確")
        except Exception as e:
            messagebox.showerror("錯誤", f"批改失敗: {str(e)}")
            self.status_var.set("發生錯誤")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            pd.DataFrame(self.results_data).to_excel(path, index=False)
            messagebox.showinfo("成功", "成績單已匯出！")

if __name__ == "__main__":
    app = AutoGraderCloud()
    app.mainloop()
