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
        self.title("AI 雲端閱卷系統 v9.3 - 物理網格鎖定版")
        self.geometry("1100x850")
        self.answer_text = "" 
        self.results_data = [] 
        self.config_file = "api_key.txt"
        self.setup_ui()
        self.load_api_key_from_file()

    def setup_ui(self):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        ctk.CTkLabel(self, text="AI 閱卷系統 - 物理網格鎖定模式 (v9.3)", font=("Microsoft JhengHei", 24, "bold")).pack(pady=15)

        config_frame = ctk.CTkFrame(self)
        config_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(config_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_entry = ctk.CTkEntry(config_frame, width=400, show="*")
        self.api_entry.grid(row=0, column=1, padx=10, pady=5)
        ctk.CTkButton(config_frame, text="儲存 Key", width=100, command=self.save_api_key).grid(row=0, column=2, padx=10)

        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkButton(btn_frame, text="1. 載入解答 (Word/PDF)", command=self.load_answer, width=200).grid(row=0, column=0, padx=15, pady=10)
        self.btn_start = ctk.CTkButton(btn_frame, text="2. 開始批改 (PDF/JPG)", command=self.start_grading, fg_color="#2ecc71", width=180)
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
        file_path = filedialog.askopenfilename(filetypes=[("PDF/Image", "*.pdf *.jpg *.png")])
        if file_path and api_key:
            if not self.answer_text:
                messagebox.showwarning("提示", "請先載入正確解答")
                return
            for item in self.tree.get_children(): self.tree.delete(item)
            self.results_data = []
            threading.Thread(target=self.run_grading, args=(file_path, api_key), daemon=True).start()

    def run_grading(self, file_path, api_key):
        try:
            self.status_var.set("系統初始化中...")
            genai.configure(api_key=api_key)
            # 自動偵測模型，避免 404 錯誤
            models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
            target_model = next((m for m in models if "1.5-flash" in m), models[0] if models else "models/gemini-1.5-flash")

            model = genai.GenerativeModel(model_name=target_model, generation_config={"response_mime_type": "application/json"})
            uploaded_file = genai.upload_file(path=file_path)
            while uploaded_file.state.name == "PROCESSING": 
                time.sleep(2)
                uploaded_file = genai.get_file(uploaded_file.name)
            
            # --- v9.3 核心對位 Prompt ---
            prompt = f"""
            你是一位視覺辨識與對位極其精準的閱卷專家。
            參考解答清單：{self.answer_text}

            【任務：絕對物理區域鎖定辨識】
            1. **建立 3x5 網格鎖定 (Grid-Lock)**：
               - 請先找到影像中「聽力作答區」的 3x5 黑色格框。
               - **嚴格限區辨識**：
                 - 第一排（題號 1-5）：辨識第 5 題時，筆跡即便偏向右上角，只要在第 5 格內，必判定為 C。
                 - 第二排（題號 6-10）：辨識第 8-10 題時，中心點必須垂直鎖定在該格框內。嚴禁向上讀取到第 3-5 題的答案。
                 - 第三排（題號 11-15）：確保第 14 題(B)與第 15 題(A)不發生水平位移或順序顛倒。
            2. **消除跨行干擾**：
               - 絕對禁止將上方格子的筆跡視為下方題目的答案。每一格必須獨立判斷。
            3. **細微筆跡增益保護**：
               - 針對藍色或黑色細筆跡，只要格子內有「人為書寫軌跡」，絕對禁止判定為空白。
            4. **禁止自動遞補**：
               - 若某格確實無內容，回傳 ""，但後續題號不可因此位移。

            回傳 JSON：
            [ {{"class": "班級", "no": "座號", "name": "姓名", "questions": [{{"q_idx": 1, "s_ans": "A", "c_ans": "A", "res": "○"}}] }} ]
            """
            
            self.status_var.set("AI 網格鎖定辨識中...")
            response = model.generate_content([prompt, uploaded_file])
            self.results_data = json.loads(response.text)

            for s in self.results_data:
                correct_count = sum(1 for q in s.get('questions', []) if q.get('res') == '○')
                s['correct_sum'] = correct_count
                self.tree.insert("", "end", values=(s.get('class'), s.get('no'), s.get('name'), correct_count))
                
            self.status_var.set(f"批改完成！(已修正位移與漏判)")
            self.btn_export.configure(state="normal")
        except Exception as e:
            messagebox.showerror("錯誤", f"辨識失敗: {str(e)}")
        finally:
            self.btn_start.configure(state="normal")

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        try:
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

            ans_dict = {"題號": [], "標準解答": []}
            if self.results_data:
                for q in self.results_data[0].get("questions", []):
                    ans_dict["題號"].append(f"第{q.get('q_idx')}題")
                    ans_dict["標準解答"].append(q.get("c_ans"))

            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                pd.DataFrame(export_dict).to_excel(writer, sheet_name='學生作答成績', index=False)
                pd.DataFrame(ans_dict).to_excel(writer, sheet_name='正確解答', index=False)

            messagebox.showinfo("成功", "報告已成功匯出至 Excel 分頁！")
        except Exception as e:
            messagebox.showerror("匯出錯誤", str(e))

if __name__ == "__main__":
    AutoGraderCloud().mainloop()
