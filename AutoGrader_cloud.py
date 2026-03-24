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
        
        ctk.CTkLabel(config_frame, text="API Key 狀態:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.api_status_label = ctk.CTkLabel(config_frame, text="尚未載入", text_color="red")
        self.api_status_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(config_frame, text="手動更新 Key:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.api_entry = ctk.CTkEntry(config_frame, width=350, show="*")
        self.api_entry.grid(row=1, column=1, padx=10, pady=5)
        
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key_to_file).grid(row=1, column=2, padx=10)

        ctk.CTkLabel(config_frame, text="PDF 內預計人數:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
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
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, columns=("Student", "Score", "Summary"), show='headings')
        self.tree.heading("Student", text="學生姓名/座號")
        self.tree.heading("Score", text="得分")
        self.tree.heading("Summary", text="AI 評語")
        self.tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.status_var = ctk.StringVar(value="狀態：請載入解答與 API Key")
        ctk.CTkLabel(self, textvariable=self.status_var).pack(pady=5)

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    key = f.read().strip()
                    if key:
                        self.api_entry.delete(0, "end")
                        self.api_entry.insert(0, key)
                        self.api_status_label.configure(text="已自動載入 API Key", text_color="green")
            except: pass

    def save_api_key_to_file(self):
        key = self.api_entry.get().strip()
        if key:
            with open(self.config_file, "w") as f: f.write(key)
            self.api_status_label.configure(text="Key 已更新並儲存", text_color="green")
            messagebox.showinfo("成功", "API Key 儲存成功！")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word 檔案", "*.docx")])
        if path:
            doc = Document(path)
            self.answer_text = "\n".join([" | ".join([c.text.strip() for c in r.cells]) for t in doc.tables for r in t.rows])
            self.status_var.set("解答載入成功")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showwarning("錯誤", "請輸入 API Key！")
            return
        path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if path:
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            threading.Thread(target=self.process_gemini, args=(path, api_key), daemon=True).start()

    def process_gemini(self, pdf_path, api_key):
        try:
            self.status_var.set("正在連線 Google AI...")
            genai.configure(api_key=api_key)
            
            # 修正：改用 models/ 前綴確保路徑正確
            model = genai.GenerativeModel('models/gemini-1.5-flash')
            
            self.status_var.set("上傳檔案中...")
            uploaded_file = genai.upload_file(path=pdf_path)
            student_count = self.student_count_entry.get()
            
            prompt = f"""
            你是一位專業教師。PDF 文件包含大約 {student_count} 位學生的答案卷。
            參考解答：
            {self.answer_text}
            
            請分別辨識每位學生的姓名、批改分數，並嚴格以 JSON 列表格式回傳（僅回傳 JSON）：
            [
              {{"student": "姓名/座號", "score": "分數", "summary": "評語"}}
            ]
            """
            
            self.status_var.set("AI 正在批改...")
            response = model.generate_content([prompt, uploaded_file])
            
            # 使用更強的 JSON 提取
            json_match = re.search(r'\[.*\]', response.text, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group())
                for s in data:
                    self.tree.insert("", "end", values=(s.get('student', '未知'), s.get('score', '0'), s.get('summary', '')))
                    self.results_data.append(s)
                self.status_var.set(f"完成！批改了 {len(data)} 位學生")
                self.btn_export.configure(state="normal")
            else:
                self.status_var.set("AI 回傳格式有誤")
                
        except Exception as e:
            error_msg = str(e)
            if "404" in error_msg:
                error_msg = "模型名稱錯誤或 API 版本不支援，請確認 API Key 權限。"
            self.status_var.set(f"錯誤: {error_msg}")
            messagebox.showerror("批改失敗", f"發生錯誤：\n{error_msg}")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            pd.DataFrame(self.results_data).to_excel(path, index=False)
            messagebox.showinfo("成功", "成績單已匯出")

if __name__ == "__main__":
    app = AutoGraderCloud()
    app.mainloop()
