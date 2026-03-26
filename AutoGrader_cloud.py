import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from docx import Document
import fitz  # PyMuPDF: 用於讀取 PDF 文字
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
        self.title("AI 雲端閱卷系統 v7.6 (Word/PDF 雙解答支援)")
        self.geometry("1100x850")
        
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        ctk.CTkLabel(self, text="AI 閱卷系統 - 診斷報告模式", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

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
        # 修改：調整為支援兩種格式
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word/PDF)", command=self.load_answer, width=200).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF考卷)", command=self.start_grading, fg_color="#2ecc71", width=180)
        self.btn_start.grid(row=0, column=1, padx=15, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出轉置 Excel", command=self.export_excel, state="disabled", width=180)
        self.btn_export.grid(row=0, column=2, padx=15, pady=10)

        # 3. 畫面顯示 (UI 顯示簡略清單)
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

    # --- 修改後的解答載入函數 ---
    def load_answer(self):
        path = filedialog.askopenfilename(filetypes=[("解答檔案", "*.docx *.pdf")])
        if not path:
            return
            
        try:
            file_ext = os.path.splitext(path)[1].lower()
            
            if file_ext == ".docx":
                with open(path, 'rb') as f:
                    doc = Document(io.BytesIO(f.read()))
                    # 讀取表格文字（舊邏輯）
                    self.answer_text = "\n".join([" | ".join([c.text.strip() for c in r.cells]) for t in doc.tables for r in t.rows])
                self.status_var.set("Word 解答載入成功")
                
            elif file_ext == ".pdf":
                doc = fitz.open(path)
                full_text = ""
                for page in doc:
                    full_text += page.get_text()
                self.answer_text = full_text
                doc.close()
                self.status_var.set("PDF 解答載入成功")
                
            if not self.answer_text.strip():
                messagebox.showwarning("警告", "無法從檔案中提取到文字內容，請確認檔案是否為掃描圖檔或空白。")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"解答檔案讀取失敗: {str(e)}")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if pdf_path and api_key:
            if not self.answer_text:
                messagebox.showwarning("提示", "請先載入正確解答 (Word 或 PDF)")
                return
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("AI 辨識中 (採用精準視覺演算法)...")
            genai.configure(api_key=api_key)
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            # 優先使用 1.5-pro 如果有的話，辨識手寫更準，否則用 flash
            target_model = next((m for m in models if "1.5-pro" in m), next((m for m in models if "1.5-flash" in m), models[0]))
            
            model = genai.GenerativeModel(model_name=target_model, generation_config={"response_mime_type": "application/json"})
            uploaded_file = genai.upload_file(path=pdf_path)
            while uploaded_file.state.name == "PROCESSING": time.sleep(2); uploaded_file = genai.get_file(uploaded_file.name)
            
            # 使用先前針對 3 號學生優化過的精準 Prompt
            prompt = f"""
            你是一位極其嚴謹、專門處理模糊手寫稿的閱卷專家。
            目前影像中是學生的手寫選擇題答案（A, B, C, D）。
            參考解答：{self.answer_text}

            【字形結構判斷準則】
            1. 在辨識時，請主動忽略所有細碎的、類似格線的噪音線段，使用下述的判斷準則判斷即可
            2. 判定為 'A'：看到兩條斜線在頂部交會，頂部連貫寫成尖銳形狀**（類似頂部尖銳的 8），而非標準的圓弧，中間有明顯橫線，結構無兩個圓弧堆疊，請判定為 **A**。
            3. 判定為 'B'：
               - **結構特徵**：只要字體是由「上下兩個閉合或半閉合的圓圈」組成，一律判定為 **B**。
               - **排除尖頂干擾**：有些學生寫 B 時頂部會變尖，或左側直線沒連上，只要它是雙層結構，就不是 A 也不是 C，請判定為 **B**。
            4. 判定為 'C'：有明顯右側開口。學生寫字母「C」時極度潦草，呈現出一個非常扁平、無力的「∠」或「<」形狀。在該形狀周圍，有一些細碎的、類似格線的噪音筆跡。如果忽略噪音後只剩下一個單純的圓弧或扁平的「<」形狀，即便看起來有些斷裂或不標準，也請將其判定為「C」。
            5. 判定為 'D'：左側一直線，右側一個大圓弧，中間無明顯橫線，請判定為 **D**。
            

            【環境排除干擾】
            - 忽略印刷格線與格線左上角的題號。
            - 請主動忽略所有細碎的、類似格線的噪音線段。
            - 以學生手寫筆跡（藍色或深灰色）為準。

            任務目標：
            1. 辨識班級、座號、姓名。
            2. 比對正確答案。正確請用 '○'，錯誤請用 '╳'。
            3. 嚴格回傳此 JSON 格式：
            [
              {{
                "class": "班級", "no": "座號", "name": "姓名",
                "questions": [
                  {{"q_idx": 1, "s_ans": "A", "c_ans": "A", "res": "○"}},
                  ...
                ]
              }}
            ]
            """
            
            response = model.generate_content([prompt, uploaded_file])
            self.results_data = json.loads(response.text)

            for s in self.results_data:
                correct_count = sum(1 for q in s.get('questions', []) if q.get('res') == '○')
                s['correct_sum'] = correct_count
                self.tree.insert("", "end", values=(s.get('class'), s.get('no'), s.get('name'), correct_count))
                
            self.status_var.set("批改完成！")
            self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("錯誤", f"批改過程出錯: {str(e)}")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            try:
                export_dict = {"項目": ["班級", "座號", "姓名", "正確總題數"]}
                max_q = 0
                for s in self.results_data:
                    max_q = max(max_q, len(s.get("questions", [])))
                
                for i in range(1, max_q + 1):
                    export_dict["項目"].append(f"第{i}題_學生答案")
                    export_dict["項目"].append(f"第{i}題_正確答案")
                    export_dict["項目"].append(f"第{i}題_比對結果")

                for s in self.results_data:
                    col_name = f"{s.get('name')}({s.get('no')})"
                    student_values = [s.get("class"), s.get("no"), s.get("name"), s.get("correct_sum")]
                    for q in s.get("questions", []):
                        student_values.append(q.get("s_ans"))
                        student_values.append(q.get("c_ans"))
                        student_values.append(q.get("res"))
                    while len(student_values) < len(export_dict["項目"]):
                        student_values.append("")
                    export_dict[col_name] = student_values

                df = pd.DataFrame(export_dict)
                df.to_excel(path, index=False)
                messagebox.showinfo("成功", f"轉置報告匯出成功！")
            except Exception as e:
                messagebox.showerror("匯出錯誤", str(e))

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
