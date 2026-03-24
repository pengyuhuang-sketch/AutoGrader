import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import google.generativeai as genai
import pandas as pd
import os
import threading
from PIL import Image

class AutoGraderCloud(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AI 雲端多卷批改系統 v2.0")
        self.geometry("900x800")
        
        self.answer_text = "" 
        self.results_data = [] # 儲存所有學生的成績
        self.setup_ui()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        
        # 標題
        ctk.CTkLabel(self, text="AI 雲端自動閱卷系統", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        # 1. API Key 設定區
        api_frame = ctk.CTkFrame(self)
        api_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(api_frame, text="Gemini API KEY:").grid(row=0, column=0, padx=10, pady=10)
        self.api_entry = ctk.CTkEntry(api_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=10)

        # 2. 參數設定區
        param_frame = ctk.CTkFrame(self)
        param_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(param_frame, text="欲批改考卷數量:").grid(row=0, column=0, padx=10, pady=10)
        self.count_entry = ctk.CTkEntry(param_frame, width=100)
        self.count_entry.insert(0, "1")
        self.count_entry.grid(row=0, column=1, padx=10, pady=10)

        # 3. 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        
        self.btn_load_ans = ctk.CTkButton(btn_frame, text="1. 載入 Word 解答", command=self.load_word)
        self.btn_load_ans.grid(row=0, column=0, padx=10, pady=10)

        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 PDF", command=self.start_batch_grading, fg_color="#2ecc71")
        self.btn_start.grid(row=0, column=1, padx=10, pady=10)

        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出 Excel 成績單", command=self.export_excel, state="disabled")
        self.btn_export.grid(row=0, column=2, padx=10, pady=10)

        # 4. 表格顯示區 (使用 ttk.Treeview)
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.tree = ttk.Treeview(table_frame, columns=("FileName", "Score", "Status"), show='headings')
        self.tree.heading("FileName", text="學生檔名")
        self.tree.heading("Score", text="得分")
        self.tree.heading("Status", text="批改狀態")
        self.tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.status_var = ctk.StringVar(value="狀態：等待操作")
        ctk.CTkLabel(self, textvariable=self.status_var).pack(pady=5)

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

    def start_batch_grading(self):
        api_key = self.api_entry.get()
        if not api_key or not self.answer_text:
            messagebox.showwarning("錯誤", "請確認 API Key 與解答檔皆已備齊")
            return
        
        try:
            num_files = int(self.count_entry.get())
        except:
            messagebox.showwarning("錯誤", "請輸入正確的數字")
            return

        files = filedialog.askopenfilenames(filetypes=[("PDF 檔案", "*.pdf")])
        if not files: return
        
        # 只取使用者設定的數量
        files_to_process = files[:num_files]
        self.results_data = []
        for item in self.tree.get_children(): self.tree.delete(item)
        
        genai.configure(api_key=api_key)
        threading.Thread(target=self.batch_process, args=(files_to_process,), daemon=True).start()

    def batch_process(self, files):
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        for file_path in files:
            file_name = os.path.basename(file_path)
            self.status_var.set(f"正在批改: {file_name}...")
            
            try:
                sample_file = genai.upload_file(path=file_path)
                prompt = f"參考解答：\n{self.answer_text}\n請批改此卷並僅回傳 JSON 格式：{{'score': 總分, 'summary': '簡短評語'}}"
                
                response = model.generate_content([prompt, sample_file])
                # 這裡假設 AI 回傳包含分數的文字，實務上可再加正則表達式擷取數字
                res_text = response.text
                score = "".join(filter(str.isdigit, res_text[:50])) # 簡單擷取數字
                
                self.results_data.append({"學生檔名": file_name, "得分": score, "狀態": "完成"})
                self.tree.insert("", "end", values=(file_name, score, "批改完成"))
            except Exception as e:
                self.tree.insert("", "end", values=(file_name, "N/A", f"錯誤: {str(e)}"))

        self.status_var.set("全部批改完成！")
        self.btn_export.configure(state="normal")

    def export_excel(self):
        if not self.results_data: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 檔案", "*.xlsx")])
        if path:
            df = pd.DataFrame(self.results_data)
            df.to_excel(path, index=False)
            messagebox.showinfo("匯出成功", f"成績單已儲存至：\n{path}")

if __name__ == "__main__":
    app = AutoGraderCloud()
    app.mainloop()