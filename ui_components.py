import tkinter as tk
from tkinter import ttk, colorchooser
from tkinter.font import families

class UIComponents:
    def __init__(self, app):
        self.app = app
        self.root = app.root

    def create_widgets(self):
        font_setting = (self.app.font_name, 10)

        # 輸入框架
        input_frame = tk.Frame(self.root, bg="#ADD8E6")
        input_frame.pack(pady=15)
        tk.Label(input_frame, text="輸入", bg="#ADD8E6", font=font_setting).pack(side=tk.LEFT)
        self.app.input_entry = tk.Entry(input_frame, width=30, font=font_setting)
        self.app.input_entry.pack(side=tk.LEFT, padx=10)
        tk.Button(input_frame, text="瀏覽", font=font_setting, command=self.app.file_handler.select_file,
                  relief="flat", bg="#4CAF50", fg="#FFFFFF").pack(side=tk.LEFT)
        
        # 貼上文字按鈕
        tk.Button(input_frame, text="貼上文字", font=font_setting, command=self.app.file_handler.open_text_input,
                  relief="flat", bg="#4CAF50", fg="#FFFFFF").pack(side=tk.LEFT, padx=5)

        # 輸出框架
        ppt_frame = tk.Frame(self.root, bg="#ADD8E6")
        ppt_frame.pack(pady=15)
        tk.Label(ppt_frame, text="結果", bg="#ADD8E6", font=font_setting).pack(side=tk.LEFT)
        self.app.ppt_entry = tk.Entry(ppt_frame, width=30, font=font_setting)
        self.app.ppt_entry.pack(side=tk.LEFT, padx=10)
        tk.Button(ppt_frame, text="另存", font=font_setting, command=self.app.file_handler.select_save_location,
                  relief="flat", bg="#4CAF50", fg="#FFFFFF").pack(side=tk.LEFT)

        # 配置框架
        config_frame = tk.Frame(self.root, bg="#ADD8E6")
        config_frame.pack(pady=10)

        # 字體選擇
        font_family_label = tk.Label(config_frame, text="字體", bg="#ADD8E6", font=font_setting)
        font_family_label.pack(side=tk.LEFT, padx=5)
        self.app.font_family_combobox = ttk.Combobox(config_frame, values=families(), font=font_setting, width=8)
        self.app.font_family_combobox.pack(side=tk.LEFT, padx=5)
        self.app.font_family_combobox.set("微軟正黑體")
        self.app.font_family_combobox.bind("<<ComboboxSelected>>", self.update_font_name)

        # 字體顏色
        font_color_label = tk.Label(config_frame, text="字色", bg="#ADD8E6", font=font_setting)
        font_color_label.pack(side=tk.LEFT, padx=5)
        self.app.font_color_btn = tk.Button(config_frame, text="選色", command=self.choose_font_color,
                                        relief="flat", bg="#808080", fg="#FFFFFF", font=font_setting)
        self.app.font_color_btn.pack(side=tk.LEFT, padx=5)

        # 頁面背景顏色
        page_bg_color_label = tk.Label(config_frame, text="頁色", bg="#ADD8E6", font=font_setting)
        page_bg_color_label.pack(side=tk.LEFT, padx=5)
        self.app.page_bg_color_btn = tk.Button(config_frame, text="選色", command=self.choose_page_bg_color,
                                           relief="flat", bg="#808080", fg="#FFFFFF", font=font_setting)
        self.app.page_bg_color_btn.pack(side=tk.LEFT, padx=5)

        # 文字對齊選項
        align_label = tk.Label(config_frame, text="對齊", bg="#ADD8E6", font=font_setting)
        align_label.pack(side=tk.LEFT, padx=5)
        self.app.align_dropdown = ttk.Combobox(config_frame, textvariable=self.app.text_align,
                                          values=["LEFT", "CENTER", "RIGHT"], state="readonly", width=6, font=font_setting)
        self.app.align_dropdown.pack(side=tk.LEFT, padx=5)

        # 比例框架
        ratio_frame = tk.Frame(self.root, bg="#ADD8E6")
        ratio_frame.pack(pady=10)
        tk.Label(ratio_frame, text="比例", bg="#ADD8E6", font=font_setting).pack(side=tk.LEFT)
        self.app.aspect_dropdown = ttk.Combobox(ratio_frame, textvariable=self.app.aspect_ratio,
                                            values=["16:9", "4:3", "10:16"], state="readonly", width=6, font=font_setting)
        self.app.aspect_dropdown.pack(side=tk.LEFT, padx=5)

        # 字體大小框架
        font_size_frame = tk.Frame(ratio_frame, bg="#ADD8E6")
        font_size_frame.pack(side=tk.LEFT, padx=10)
        tk.Label(font_size_frame, text="字體大小", bg="#ADD8E6", font=font_setting).pack(side=tk.LEFT)
        font_sizes = [str(i) for i in range(8, 74, 2)]  # 8到72的偶數
        self.app.font_size_dropdown = ttk.Combobox(font_size_frame, textvariable=self.app.font_size_var,
                                               values=font_sizes, state="readonly", width=4, font=font_setting)
        self.app.font_size_dropdown.pack(side=tk.LEFT, padx=5)
        self.app.font_size_dropdown.bind("<<ComboboxSelected>>", self.update_font_size)

        # 模板選擇
        template_label = tk.Label(font_size_frame, text="模板", bg="#ADD8E6", font=font_setting)
        template_label.pack(side=tk.LEFT, padx=5)
        self.app.template_combobox = ttk.Combobox(font_size_frame, values=["空白模板", "自定義..."], 
                                        state="readonly", width=8, font=font_setting)
        self.app.template_combobox.pack(side=tk.LEFT, padx=5)
        self.app.template_combobox.set("空白模板")
        self.app.template_combobox.bind("<<ComboboxSelected>>", self.update_template)
        
        # 模板瀏覽按鈕
        self.app.template_browse_btn = tk.Button(font_size_frame, text="瀏覽", command=self.app.file_handler.select_template,
                                        relief="flat", bg="#808080", fg="#FFFFFF", font=font_setting)
        self.app.template_browse_btn.pack(side=tk.LEFT, padx=5)

        # 轉換按鈕
        self.app.convert_btn = tk.Button(self.root, text="開始轉檔", font=("標楷體", 18),
                                     command=self.app.converter.start_conversion, relief="flat", width=8, height=1, bg="#4CAF50", fg="#FFFFDD")
        self.app.convert_btn.pack(pady=15)

        # 載入標籤
        self.app.loading_label = tk.Label(self.root, text="", bg="#ADD8E6", font=("標楷體", 14))
        self.app.loading_label.pack(pady=5)

    def update_font_name(self, event=None):
        self.app.font_name = self.app.font_family_combobox.get()

    def update_font_size(self, event=None):
        self.app.font_size = int(self.app.font_size_var.get())

    def choose_font_color(self):
        color_code = colorchooser.askcolor(title="選擇字體顏色")
        if color_code[0] is not None:
            rgb = color_code[0]
            if isinstance(rgb, tuple) and len(rgb) == 3:
                self.app.font_color = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
                self.app.font_color_btn.config(
                    bg=f'#{int(rgb[0]):02x}{int(rgb[1]):02x}{int(rgb[2]):02x}')

    def choose_page_bg_color(self):
        color_code = colorchooser.askcolor(title="選擇頁面背景顏色")
        if color_code[0] is not None:
            rgb = color_code[0]
            if isinstance(rgb, tuple) and len(rgb) == 3:
                self.app.page_bg_color = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
                self.app.page_bg_color_btn.config(
                    bg=f'#{int(rgb[0]):02x}{int(rgb[1]):02x}{int(rgb[2]):02x}')

    def update_template(self, event=None):
        if self.app.template_combobox.get() == "空白模板":
            self.app.template_path = None
            self.app.template_browse_btn.config(state=tk.NORMAL)