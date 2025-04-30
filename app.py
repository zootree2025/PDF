import tkinter as tk
import os
from ui_components import UIComponents
from file_handler import FileHandler
from converter import Converter
from typing import Optional

class PDFToPPTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF/Word 轉 PPT - 防呆版")
        self.root.configure(bg="#ADD8E6")
        self.root.geometry("450x360")
        self.root.resizable(False, False)

        # 初始化屬性
        self.input_path: Optional[str] = None
        self.ppt_path: Optional[str] = None
        self.input_type: Optional[str] = None
        self.template_path: Optional[str] = None
        self.converting = False
        self.success_flag = {'ok': True}
        self.last_dir = os.path.expanduser("~/Desktop")

        # 初始化樣式設定
        self.font_size = 24
        self.font_color = (0, 0, 0)
        self.font_name = "標楷體"
        self.page_bg_color = None
        self.text_align = tk.StringVar(value="LEFT")
        self.aspect_ratio = tk.StringVar(value="16:9")
        self.font_size_var = tk.StringVar(value="24")

        # 初始化元件
        self.ui = UIComponents(self)
        self.file_handler = FileHandler(self)
        self.converter = Converter(self)

        # 創建 UI
        self.ui.create_widgets()