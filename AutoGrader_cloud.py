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
        self.title("AI 雲端閱卷系統 v9.0 - 容錯校正與筆跡增益版")
        self.geometry("1100x850")
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        ctk.CTkLabel(self, text="AI 閱卷系統 - 最終容錯辨識模式", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(config_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_entry = ctk.CTkEntry(config_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key).grid(row=0, column=2, padx=10)

        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word/PDF)", command=self.load_answer, width=200).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF考卷)", command=self.start_grading, fg_color="#2ecc71", width=180)
        self.btn_start.grid(row=0, column=1, padx=15, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出分頁 Excel", command=self.export_excel, state="disabled", width=180)
        self.btn_export.grid(row=0, column=2, padx=15, pady=10)

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
            self.status_var.set("解答載入成功")
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取失敗: {str(e)}")

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
            self.status_var.set("自動偵測可用模型...")
            genai.configure(api_key=api_key)
            # 自動抓取可用的 flash 模型
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in models if "1.5-flash" in m), models[0] if models else "models/gemini-1.5-flash")

            model = genai.GenerativeModel(model_name=target_model, generation_config={"response_mime_type": "application/json"})
            uploaded_file = genai.upload_file(path=pdf_path)
            while uploaded_file.state.name == "PROCESSING": 
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
            
            # --- v9.0 容錯核心 Prompt ---
            prompt = f"""
            你是一位視覺閱卷專家。影像中是學生的手寫答案卷。
            參考解答清單：{self.answer_text}

            【辨識強化指令：不漏抓、不位移】
            1. **容錯對位**：以印刷題號為中心，搜尋其鄰近格內的筆劃。即使影像稍微傾斜或題號沒對齊格子，也請找出對應的答案。
            2. **筆跡增益保護 (解決空白誤判)**：
               - 格子內無論是藍色、黑色或淺灰色墨跡，只要有「人為書寫軌跡」，絕對不准判定為空白。
               - 即使筆劃極淡或斷裂，也請根據輪廓嘗試判定為 A, B, C 或 D。
               - 只有在該區域「完全純白」且無任何墨跡時，才可回傳 ""。
            3. **特徵優先**：
               - 'A'：頂部尖或圓，中間有橫線（若橫線極淡也算 A）。
               - 'C'：左側圓弧，右側明顯開口。
            4. **禁止位移遞補**：每一題號 (q_idx) 必須與影像中的題號標籤嚴格對應。

            任務：辨識班級、座號、姓名，並逐題比對。
            回傳 JSON：
            [ {{"class": "班級", "no": "座號", "name": "姓名", "questions": [{{"q_idx": 1, "s_ans": "A", "c_ans": "A", "res": "○"}}] }} ]
            """
            
            self.status_var.set("AI 深度容錯辨識中...")
            response = model.generate_content([prompt, uploaded_file])
            self.results_data = json.loads(response.text)

            for s in self.results_data:
                correct_count = sum(1 for q in s.get('questions', []) if q.get('res') == '○')
                s['correct_sum'] = correct_count
                self.tree.insert("", "end", values=(s.get('class'), s.get('no'), s.get('name'), correct_count))
                
            self.status_var.set(f"批改完成！已使用模型: {target_model}")
            self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("錯誤", f"辨識失敗: {str(e)}")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try:
            # 學生作答成績分頁
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
            ans_dict = {"題號": [], "正確解答": []}
            if self.results_data:
                for q in self.results_data[0].get("questions", []):
                    ans_dict["題號"].append(f"第{q.get('q_idx')}題")
                    ans_dict["正確解答"].append(q.get("c_ans"))

            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                pd.DataFrame(export_dict).to_excel(writer, sheet_name='學生作答成績', index=False)
                pd.DataFrame(ans_dict).to_excel(writer, sheet_name='正確解答', index=False)

            messagebox.showinfo("成功", "分頁報告匯出成功！")
        except Exception as e:
            messagebox.showerror("匯出錯誤", str(e))

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
