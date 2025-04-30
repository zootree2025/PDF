import os
import tkinter as tk
from tkinter import filedialog, messagebox

class FileHandler:
    def __init__(self, app):  # 確保接受 app 參數
        self.app = app
        self.last_dir = os.path.expanduser("~/Desktop")

    def select_file(self):
        filename = filedialog.askopenfilename(filetypes=[("所有支援格式", "*.pdf *.docx *.txt")], initialdir=self.last_dir)
        if filename:
            self.app.input_path = filename
            self.last_dir = os.path.dirname(filename)
            self.app.input_entry.delete(0, tk.END)
            self.app.input_entry.insert(0, filename)
            self.save_ppt()

    def select_save_location(self):
        if not self.app.input_path:
            messagebox.showerror("錯誤", "請先選擇輸入文件！")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            initialfile=f"{os.path.splitext(os.path.basename(self.app.input_path))[0]}.pptx",
            filetypes=[("PowerPoint 文件", "*.pptx")],
            initialdir=self.last_dir
        )
        if filename:
            self.app.ppt_path = filename
            self.last_dir = os.path.dirname(filename)
            self.app.ppt_entry.delete(0, tk.END)
            self.app.ppt_entry.insert(0, filename)

    def save_ppt(self):
        if self.app.input_path:
            base = os.path.splitext(os.path.basename(self.app.input_path))[0]
            template_suffix = "_template" if self.app.template_path else ""
            self.app.ppt_path = os.path.join(os.path.dirname(self.app.input_path), 
                                       f"{base}{template_suffix}.pptx")
            self.app.ppt_entry.delete(0, tk.END)
            self.app.ppt_entry.insert(0, self.app.ppt_path)

    def select_template(self):
        filename = filedialog.askopenfilename(
            filetypes=[("PPT模板", "*.pptx")],
            initialdir=self.last_dir
        )
        if filename:
            self.app.template_path = filename
            self.app.template_combobox.set(os.path.basename(filename))

    def open_text_input(self):
        # 創建新視窗
        text_window = tk.Toplevel(self.app.root)
        text_window.title("輸入文字")
        text_window.geometry("800x600")  # 設置更大的初始尺寸
        text_window.configure(bg="#ADD8E6")
        
        # 設置視窗最大化
        text_window.state('zoomed')  # 在 Windows 上最大化視窗
        
        # 添加說明標籤
        tk.Label(text_window, text="請輸入或貼上文字內容：", bg="#ADD8E6", font=(self.app.font_name, 12)).pack(pady=10)
        
        # 添加文字框
        text_frame = tk.Frame(text_window, bg="#ADD8E6")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        text_box = tk.Text(text_frame, font=(self.app.font_name, 12), wrap=tk.WORD)
        text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滾動條
        scrollbar = tk.Scrollbar(text_frame, command=text_box.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_box.config(yscrollcommand=scrollbar.set)
        
        # 添加按鈕框架
        button_frame = tk.Frame(text_window, bg="#ADD8E6")
        button_frame.pack(pady=15)
        
        # 添加確認按鈕
        tk.Button(button_frame, text="確認", font=(self.app.font_name, 12), 
                  command=lambda: self.process_text_input(text_box.get("1.0", tk.END), text_window),
                  relief="flat", bg="#4CAF50", fg="#FFFFFF", width=8).pack(side=tk.LEFT, padx=10)
        
        # 添加取消按鈕
        tk.Button(button_frame, text="取消", font=(self.app.font_name, 12), command=text_window.destroy,
                  relief="flat", bg="#808080", fg="#FFFFFF", width=8).pack(side=tk.LEFT, padx=10)

    def process_text_input(self, text, window):
        if not text.strip():
            messagebox.showerror("錯誤", "請輸入文字內容！")
            return
        
        # 創建臨時文字文件
        temp_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp")
        os.makedirs(temp_dir, exist_ok=True)
        temp_file = os.path.join(temp_dir, "temp_text_input.txt")
        
        with open(temp_file, "w", encoding="utf-8") as f:
            f.write(text)
        
        # 設置輸入路徑為臨時文件
        self.app.input_path = temp_file
        self.app.input_type = "txt"  # 明確設置檔案類型
        self.app.input_entry.delete(0, tk.END)
        self.app.input_entry.insert(0, "直接輸入的文字")
        
        # 確保輸出目錄存在
        output_dir = os.path.dirname(os.path.abspath(__file__))  # 使用當前程式所在目錄
        
        # 設置輸出路徑
        self.app.ppt_path = os.path.join(output_dir, "文字輸入.pptx")
        self.app.ppt_entry.delete(0, tk.END)
        self.app.ppt_entry.insert(0, self.app.ppt_path)
        
        # 關閉視窗
        window.destroy()
        
        # 提示用戶
        messagebox.showinfo("成功", "文字已準備好，請點擊「開始轉檔」按鈕進行轉換。")