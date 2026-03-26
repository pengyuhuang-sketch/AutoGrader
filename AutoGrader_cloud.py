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
        self.title("AI 雲端閱卷系統 v8.6 - 動態特徵平衡版")
        self.geometry("1100x850")
        
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        ctk.CTkLabel(self, text="AI 閱卷系統 - 視覺辨識與筆跡保護模式", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

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
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word/PDF)", command=self.load_answer, width=200).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF考卷)", command=self.start_grading, fg_color="#2ecc71", width=180)
        self.btn_start.grid(row=0, column=1, padx=15, pady=10)
        self.btn_export = ctk.CTkButton(btn_frame, text="3. 匯出分頁 Excel", command=self.export_excel, state="disabled", width=180)
        self.btn_export.grid(row=0, column=2, padx=15, pady=10)

        # 3. 畫面顯示清單
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
                self.status_var.set("Word 解答載入成功")
            elif file_ext == ".pdf":
                doc = fitz.open(path)
                full_text = "".join([page.get_text() for page in doc])
                self.answer_text = full_text
                doc.close()
                self.status_var.set("PDF 解答載入成功")
        except Exception as e:
            messagebox.showerror("錯誤", f"解答檔案開啟失敗: {str(e)}")

    def start_grading(self):
        api_key = self.api_entry.get().strip()
        pdf_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if pdf_path and api_key:
            if not self.answer_text:
                messagebox.showwarning("提示", "請先載入正確解答檔案")
                return
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            threading.Thread(target=self.run_grading, args=(pdf_path, api_key), daemon=True).start()

    def run_grading(self, pdf_path, api_key):
        try:
            self.status_var.set("正在獲取可用 AI 模型路徑...")
            genai.configure(api_key=api_key)
            
            # --- 自動模型路徑抓取 ---
            available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in available_models if "gemini-1.5-flash" in m), None)
            if not target_model: target_model = available_models[0] if available_models else "models/gemini-1.5-flash"

            model = genai.GenerativeModel(model_name=target_model, generation_config={"response_mime_type": "application/json"})
            
            self.status_var.set("上傳考卷中並準備辨識...")
            uploaded_file = genai.upload_file(path=pdf_path)
            while uploaded_file.state.name == "PROCESSING": 
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
            
            # --- 動態平衡版 Prompt (避免誤傷細微筆跡) ---
            prompt = f"""
            你是一位極其敏銳的視覺閱卷專家。任務是辨識學生手寫答案。
            參考解答清單：{self.answer_text}

            【辨識準則：筆跡優先於去噪】
            1. **保護細微筆跡**：
               - 部分學生書寫極輕，線條可能呈現淡灰色或斷續。只要其軌跡構成字母特徵（如 A 的橫線、B 的接縫），請務必視為有效答案，不可當作雜訊過濾。
            2. **嚴謹判定空白**：
               - **判定為空 ""**：僅在格子內完全無人為書寫墨跡時回傳。
               - **判定為有值**：只要格內有明顯區別於印刷背景的人為墨跡（即便很淡），就必須依據結構歸類。
            3. **結構特徵鑑定**：
               - 'A'：尋找斜線匯合特徵，留意中間極淡的橫線。
               - 'B'：尋找上下兩個閉合圓弧的輪廓。
               - 'C'：左側圓弧右側開口。若格內只有一條微弱弧形筆畫，優先判定為 C。
            4. **處理塗改**：忽略劃掉的斜線，辨識最清晰或重疊在上方的最新筆畫。

            【回傳規範】
            - 若為空值，s_ans 給 ""，結果 res 給 "╳"。
            - 嚴格遵守 JSON 格式。

            JSON 格式：
            [ {{"class": "班級", "no": "座號", "name": "姓名", "questions": [{{"q_idx": 1, "s_ans": "A", "c_ans": "A", "res": "○"}}] }} ]
            """
            
            self.status_var.set(f"AI 深度辨識中 ({target_model})...")
            response = model.generate_content([prompt, uploaded_file])
            self.results_data = json.loads(response.text)

            for s in self.results_data:
                correct_count = sum(1 for q in s.get('questions', []) if q.get('res') == '○')
                s['correct_sum'] = correct_count
                self.tree.insert("", "end", values=(s.get('class'), s.get('no'), s.get('name'), correct_count))
                
            self.status_var.set("批改完成！(已啟用筆跡保護機制)")
            self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("批改出錯", f"詳細原因: {str(e)}")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try:
            # 1. 學生作答成績分頁
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
            df_main = pd.DataFrame(export_dict)

            # 2. 正確解答對照分頁
            ans_dict = {"題號": [], "標準解答": []}
            if self.results_data:
                for q in self.results_data[0].get("questions", []):
                    ans_dict["題號"].append(f"第{q.get('q_idx')}題")
                    ans_dict["標準解答"].append(q.get("c_ans"))
            df_ans = pd.DataFrame(ans_dict)

            # 使用 ExcelWriter 分頁匯出
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df_main.to_excel(writer, sheet_name='學生作答成績', index=False)
                df_ans.to_excel(writer, sheet_name='正確解答', index=False)

            messagebox.showinfo("成功", "Excel 報告已成功匯出（含正確解答分頁）")
        except Exception as e:
            messagebox.showerror("匯出錯誤", f"無法儲存 Excel: {str(e)}")

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
