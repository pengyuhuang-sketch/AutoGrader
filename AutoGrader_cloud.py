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
        self.title("AI 雲端多生閱卷系統 v7.4 (結構化數據版)")
        self.geometry("1100x850")
        
        self.answer_text = "" 
        self.results_data = [] # 存儲原始 JSON 數據
        self.config_file = "api_key.txt"
        
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        ctk.CTkLabel(self, text="AI 雲端自動閱卷系統 (Structured Data)", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        # 1. API 配置區
        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(config_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_entry = ctk.CTkEntry(config_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key).grid(row=0, column=2, padx=10)

        # 2. 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word)", command=self.load_word, width=180).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF)", command=self.start_grading, fg_color="#2ecc71", width=180)
        self.btn_start.grid(row=0, column=1, padx=15, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出 Excel (多欄位)", command=self.export_excel, state="disabled", width=180)
        self.btn_export.grid(row=0, column=2, padx=15, pady=10)

        # 3. 表格區域 (UI 僅顯示總覽)
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, columns=("Class", "No", "Name", "Score", "Status"), show='headings')
        self.tree.heading("Class", text="班級"); self.tree.heading("No", text="座號"); 
        self.tree.heading("Name", text="姓名"); self.tree.heading("Score", text="總分"); self.tree.heading("Status", text="狀態")
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.status_var = ctk.StringVar(value="狀態：準備就緒")
        ctk.CTkLabel(self, textvariable=self.status_var).pack(pady=5)

    def load_api_key_from_file(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, "r") as f: self.api_entry.insert(0, f.read().strip())

    def save_api_key(self):
        with open(self.config_file, "w") as f: f.write(self.api_entry.get())
        messagebox.showinfo("成功", "Key 已儲存")

    def load_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if path:
            with open(path, 'rb') as f:
                doc = Document(io.BytesIO(f.read()))
                self.answer_text = "\n".join([" | ".join([c.text.strip() for c in r.cells]) for t in doc.tables for r in t.rows])
            self.status_var.set("解答載入完成")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if pdf_path and api_key:
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("AI 模型連線中...")
            genai.configure(api_key=api_key)
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in models if "1.5-flash" in m), models[0])
            
            model = genai.GenerativeModel(model_name=target_model, generation_config={"response_mime_type": "application/json"})
            uploaded_file = genai.upload_file(path=pdf_path)
            while uploaded_file.state.name == "PROCESSING": time.sleep(2); uploaded_file = genai.get_file(uploaded_file.name)
            
            prompt = f"""
            你是一位自動閱卷機器人。
            參考解答：{self.answer_text}
            任務：
            1. 辨識班級、座號、姓名。
            2. 逐題比對。
            3. 嚴格回傳此 JSON 格式：
            [
              {{
                "class": "班級", "no": "座號", "name": "姓名", "total_score": 100,
                "questions": [
                  {{"q_idx": 1, "student_ans": "A", "correct_ans": "A", "result": "✅"}},
                  {{"q_idx": 2, "student_ans": "B", "correct_ans": "C", "result": "❌"}}
                ]
              }}
            ]
            """
            
            response = model.generate_content([prompt, uploaded_file])
            self.results_data = json.loads(response.text)

            for s in self.results_data:
                self.tree.insert("", "end", values=(s.get('class'), s.get('no'), s.get('name'), s.get('total_score'), "完成"))
                
            self.status_var.set("批改完成！")
            self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            final_rows = []
            for student in self.results_data:
                # 基礎資料
                row = {
                    "班級": student.get("class"),
                    "座號": student.get("no"),
                    "姓名": student.get("name"),
                    "總分": student.get("total_score")
                }
                # 展開每一題
                for q in student.get("questions", []):
                    idx = q.get("q_idx")
                    row[f"第{idx}題_學生答案"] = q.get("student_ans")
                    row[f"第{idx}題_正確答案"] = q.get("correct_ans")
                    row[f"第{idx}題_結果"] = q.get("result")
                final_rows.append(row)
            
            pd.DataFrame(final_rows).to_excel(path, index=False)
            messagebox.showinfo("成功", "多欄位 Excel 匯出成功！")

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
