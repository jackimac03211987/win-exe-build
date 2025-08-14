import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageEnhance, ImageFilter, ImageChops
import pandas as pd
import os
import io
import tempfile
from pathlib import Path
import sys
from pdf2image import convert_from_path, convert_from_bytes
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import Color
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import shutil
import threading
import re
import decimal
import platform
import glob
import numpy as np
from math import sin, cos, radians
import json
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import time


class ModernGlassUI:
    """现代化磨砂玻璃UI样式管理器"""
    
    def __init__(self):
        self.colors = {
            'primary': '#2C3E50',
            'secondary': '#3498DB',
            'accent': '#E74C3C',
            'success': '#27AE60',
            'warning': '#F39C12',
            'info': '#3498DB',
            'dark': '#1A252F',
            'light': '#ECF0F1',
            'white': '#FFFFFF',
            'glass_bg': 'rgba(255, 255, 255, 0.1)',
            'glass_border': 'rgba(255, 255, 255, 0.2)',
            'shadow': 'rgba(0, 0, 0, 0.1)',
            'text_primary': '#2C3E50',
            'text_secondary': '#7F8C8D'
        }
        
        self.fonts = {
            'title': ('Segoe UI', 12, 'bold'),
            'heading': ('Segoe UI', 10, 'bold'),
            'normal': ('Segoe UI', 9),
            'small': ('Segoe UI', 8)
        }
        
        self.setup_styles()
    
    def setup_styles(self):
        """设置现代化样式"""
        style = ttk.Style()
        
        # 配置主题颜色
        bg_color = '#F8F9FA'
        fg_color = self.colors['text_primary']
        
        # 设置主窗口样式
        style.configure('TFrame', background=bg_color)
        style.configure('TLabelframe', background=bg_color, 
                       bordercolor=self.colors['glass_border'], relief='solid', borderwidth=1)
        style.configure('TLabelframe.Label', background=bg_color, 
                       foreground=self.colors['text_primary'], font=self.fonts['heading'])
        
        # 按钮样式 - 现代化设计
        style.configure('TButton', 
                       background=self.colors['secondary'],
                       foreground=self.colors['text_primary'],
                       borderwidth=0,
                       focuscolor='none',
                       font=self.fonts['normal'],
                       padding=(12, 8))
        style.map('TButton',
                 background=[('active', '#2980B9'), ('pressed', '#21618C')],
                 foreground=[('active', self.colors['text_primary']), ('pressed', self.colors['text_primary']), ('disabled', 'gray')])
        
        # 标签样式
        style.configure('TLabel', background=bg_color, 
                       foreground=self.colors['text_primary'], font=self.fonts['normal'])
        style.configure('Title.TLabel', background=bg_color,
                       foreground=self.colors['primary'], font=self.fonts['title'])
        style.configure('Heading.TLabel', background=bg_color,
                       foreground=self.colors['text_primary'], font=self.fonts['heading'])
        
        # 输入框样式
        style.configure('TEntry', 
                       fieldbackground='white',
                       bordercolor=self.colors['glass_border'],
                       borderwidth=1,
                       relief='solid',
                       font=self.fonts['normal'])
        style.map('TEntry',
                 bordercolor=[('focus', self.colors['secondary'])])
        
        # 下拉框样式
        style.configure('TCombobox', 
                       fieldbackground='white',
                       bordercolor=self.colors['glass_border'],
                       borderwidth=1,
                       relief='solid',
                       font=self.fonts['normal'])
        style.map('TCombobox',
                 bordercolor=[('focus', self.colors['secondary'])])
        
        # 复选框样式
        style.configure('TCheckbutton', 
                       background=bg_color,
                       foreground=self.colors['text_primary'],
                       font=self.fonts['normal'])
        
        # 单选框样式
        style.configure('TRadiobutton', 
                       background=bg_color,
                       foreground=self.colors['text_primary'],
                       font=self.fonts['normal'])
        
        # 滚动条样式
        style.configure('TScrollbar', 
                       background=bg_color,
                       bordercolor=self.colors['glass_border'],
                       troughcolor=self.colors['glass_border'],
                       width=8)
        style.map('TScrollbar',
                 background=[('active', self.colors['secondary'])])
        
        # 进度条样式
        style.configure('TProgressbar', 
                       background=self.colors['secondary'],
                       troughcolor=self.colors['glass_border'],
                       borderwidth=0)
        
        # 选项卡样式
        style.configure('TNotebook', background=bg_color, borderwidth=0)
        style.configure('TNotebook.Tab', 
                       background='#F0F0F0',
                       foreground='#000000',
                       padding=(16, 8),
                       borderwidth=1,
                       relief='solid',
                       font=self.fonts['normal'])
        style.map('TNotebook.Tab',
                 background=[('selected', self.colors['secondary'])])
        
        # 分隔线样式
        style.configure('TSeparator', background=self.colors['glass_border'])


class PDFWatermarkTool:
    def __init__(self, master):
        # 设置Poppler路径（在其他初始化之前）
        self.setup_poppler_path()

        self.master = master
        master.title("PDF批量水印工具+邮件自动发送系统")
        master.geometry("1280x800")
        
        # 初始化UI样式管理器
        self.ui = ModernGlassUI()
        
        # 设置现代化主窗口背景
        self.configure_window_background()

        # 初始化变量
        self.pdf_path = None
        self.excel_path = None
        self.company_names = []
        self.company_emails = []  # 公司邮箱列表，支持一对多关系
        self.company_email_map = {}  # 公司名称到邮箱列表的映射
        self.watermarked_pdfs = []
        self.preview_image = None
        self.temp_dir = tempfile.mkdtemp()  # 创建临时目录

        # 邮件相关变量
        self.email_subject = tk.StringVar(value="您的加水印文件已准备好")
        self.email_body = tk.StringVar(
            value="尊敬的{company}，\n\n您的加水印文件已处理完成，请查收附件。\n\n如有任何问题，请及时联系我们。\n\n祝好！")
        self.smtp_server = tk.StringVar(value="smtp.exmail.qq.com")
        self.smtp_port = tk.IntVar(value=465)
        self.smtp_username = tk.StringVar()
        self.smtp_password = tk.StringVar()
        self.sender_name = tk.StringVar(value="系统管理员")
        self.enable_email = tk.BooleanVar(value=False)  # 是否启用邮件发送

        # 初始化默认值
        self.text_color = "#FF0000"  # 默认红色
        self.prefix_text = tk.StringVar(value="IDC圈：仅限")  # 设置默认前缀
        self.suffix_text = tk.StringVar(value="内部使用，转发侵权")  # 设置默认后缀
        self.font_family = tk.StringVar()  # 将在load_default_settings中设置
        self.font_size = tk.IntVar()  # 将在load_default_settings中设置
        self.watermark_angle = tk.IntVar()  # 将在load_default_settings中设置
        self.watermark_density = tk.IntVar()  # 将在load_default_settings中设置
        self.watermark_position = tk.StringVar()  # 将在load_default_settings中设置
        self.conversion_quality = tk.IntVar()  # 将在load_default_settings中设置
        self.compression_level = tk.IntVar()  # 将在load_default_settings中设置
        self.enable_rasterize = tk.BooleanVar()  # 将在load_default_settings中设置

        # 新增高级效果参数
        self.effect_type = tk.StringVar()  # 将在load_default_settings中设置
        self.outline_width = tk.IntVar()  # 将在load_default_settings中设置
        self.shadow_offset = tk.IntVar()  # 将在load_default_settings中设置
        self.effect_intensity = tk.IntVar()  # 将在load_default_settings中设置
        self.pattern_density = tk.IntVar()  # 将在load_default_settings中设置
        self.filename_pattern = tk.StringVar()  # 将在load_default_settings中设置

        # 创建颜色按钮存储列表
        self.color_buttons = []

        # 获取系统中文字体 - 移到UI设置之前
        self.system_fonts = self.get_system_fonts()

        # 设置样式（使用现代化样式）
        self.setup_ui_styles()

        # 设置UI
        self.setup_ui()

        # 加载默认设置
        self.load_default_settings()

        # 加载邮件设置
        self.load_email_settings()
    
    def configure_window_background(self):
        """配置现代化窗口背景"""
        self.master.configure(bg='#F8F9FA')
        
        # 创建渐变背景效果
        bg_frame = tk.Frame(self.master, bg='#F8F9FA')
        bg_frame.place(relwidth=1, relheight=1)
        
        # 添加阴影效果和现代感
        self.master.configure(highlightbackground='#E9ECEF', highlightthickness=1)
    
    def setup_ui_styles(self):
        """设置UI样式"""
        # 现代化样式已在ModernGlassUI中配置
        pass

    def setup_poppler_path(self):
        """确保Poppler工具在应用程序路径中可用"""
        # 获取应用程序运行路径
        if getattr(sys, 'frozen', False):
            # 如果是打包后的应用
            application_path = os.path.dirname(sys.executable)

            # 如果是Mac应用包
            if platform.system() == 'Darwin':
                # 检查各种可能的Mac应用包结构
                if os.path.isdir(os.path.join(application_path, '../Resources')):
                    application_path = os.path.join(application_path, '../Resources')
                elif os.path.isdir(os.path.join(application_path, 'Contents/Resources')):
                    application_path = os.path.join(application_path, 'Contents/Resources')
                elif os.path.isdir(os.path.join(application_path, 'Contents/MacOS')):
                    application_path = os.path.join(application_path, 'Contents/MacOS')
        else:
            # 如果是直接运行Python脚本
            application_path = os.path.dirname(os.path.abspath(__file__))

        # 将应用程序路径添加到系统PATH中，使pdf2image能找到poppler二进制文件
        os.environ['PATH'] = application_path + os.pathsep + os.environ.get('PATH', '')

        # 打印调试信息
        print(f"设置Poppler路径: {application_path}")
        print(f"系统PATH: {os.environ['PATH']}")

        # 检查是否有从PyInstaller打包时添加的Poppler二进制文件
        poppler_files = glob.glob(os.path.join(application_path, "pdftoppm*"))
        if poppler_files:
            print(f"找到Poppler文件: {poppler_files}")
        else:
            print("未在应用程序目录中找到Poppler文件，将使用系统Poppler")

    def get_system_fonts(self):
        """获取系统安装的字体"""
        system_fonts = {}

        # Windows系统字体路径
        if platform.system() == 'Windows':
            font_dir = 'C:\\Windows\\Fonts'
            system_fonts = {
                "宋体": "C:\\Windows\\Fonts\\simsun.ttc",
                "黑体": "C:\\Windows\\Fonts\\simhei.ttf",
                "微软雅黑": "C:\\Windows\\Fonts\\msyh.ttc",
                "微软雅黑粗体": "C:\\Windows\\Fonts\\msyhbd.ttc",  # 添加粗体版本
                "Arial": "C:\\Windows\\Fonts\\arial.ttf",
                "Times New Roman": "C:\\Windows\\Fonts\\times.ttf",
                "Arial Black": "C:\\Windows\\Fonts\\ariblk.ttf",  # 添加非常宽的字体
                "Impact": "C:\\Windows\\Fonts\\impact.ttf"  # 添加非常宽的字体
            }

        # MacOS系统字体路径
        elif platform.system() == 'Darwin':  # macOS
            font_dirs = [
                '/System/Library/Fonts',
                '/Library/Fonts',
                os.path.expanduser('~/Library/Fonts'),
                '/System/Library/Fonts/Supplemental'  # 添加补充字体
            ]

            # 增强的 macOS 中文字体映射
            potential_fonts = {
                "宋体": ["STSong", "Songti", "SimSun", "Songti.ttc", "宋体", "STSongti-SC"],
                "黑体": ["STHeiti", "Heiti", "SimHei", "Heiti.ttc", "黑体", "STHeiti-Medium"],
                "微软雅黑": ["Microsoft YaHei", "MicrosoftYaHei", "微软雅黑", "PingFang", "PingFangSC"],
                "微软雅黑粗体": ["Microsoft YaHei Bold", "PingFang SC Bold", "PingFang-SC-Bold"],
                "Arial": ["Arial", "Arial.ttf", "ArialMT"],
                "Times New Roman": ["Times New Roman", "Times", "TimesNewRoman"],
                "Arial Black": ["Arial Black", "Arial-Black"],
                "Impact": ["Impact", "Impact.ttf"]
            }

            # 查找字体文件
            for font_name, patterns in potential_fonts.items():
                for font_dir in font_dirs:
                    for pattern in patterns:
                        matches = glob.glob(f"{font_dir}/**/*{pattern}*", recursive=True)
                        if matches:
                            system_fonts[font_name] = matches[0]
                            break
                    if font_name in system_fonts:
                        break

            # 确保我们至少有一些基本字体
            if "黑体" not in system_fonts:
                # 尝试查找任何中文字体作为备用
                for font_dir in font_dirs:
                    cn_fonts = glob.glob(f"{font_dir}/**/华文*.ttf", recursive=True)
                    cn_fonts += glob.glob(f"{font_dir}/**/STSong*.ttf", recursive=True)
                    cn_fonts += glob.glob(f"{font_dir}/**/Songti*.ttc", recursive=True)
                    cn_fonts += glob.glob(f"{font_dir}/**/Heiti*.ttc", recursive=True)
                    if cn_fonts:
                        system_fonts["黑体"] = cn_fonts[0]
                        break

        # Linux系统字体路径
        elif platform.system() == 'Linux':
            font_dirs = [
                '/usr/share/fonts/',
                '/usr/local/share/fonts/',
                os.path.expanduser('~/.fonts/')
            ]

            # 常见中文字体映射
            potential_fonts = {
                "宋体": ["SimSun", "simsun", "song"],
                "黑体": ["SimHei", "simhei"],
                "微软雅黑": ["Microsoft YaHei", "msyh"],
                "微软雅黑粗体": ["Microsoft YaHei Bold", "msyhbd"],
                "Arial": ["Arial", "arial"],
                "Times New Roman": ["Times New Roman", "times"],
                "Arial Black": ["Arial Black", "ariblk"],
                "Impact": ["Impact"]
            }

            # 查找字体文件
            for font_name, patterns in potential_fonts.items():
                for font_dir in font_dirs:
                    for pattern in patterns:
                        matches = glob.glob(f"{font_dir}/**/*{pattern}*", recursive=True)
                        if matches:
                            system_fonts[font_name] = matches[0]
                            break
                    if font_name in system_fonts:
                        break

        # 记录字体查找结果
        for font_name, path in system_fonts.items():
            print(f"找到字体: {font_name} -> {path}")

        return system_fonts

    def setup_ui(self):
        # 创建现代化选项卡容器
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # 创建各选项卡
        self.pdf_tab = self.create_modern_frame(self.notebook)
        self.excel_tab = self.create_modern_frame(self.notebook)
        self.watermark_tab = self.create_modern_frame(self.notebook)
        self.email_tab = self.create_modern_frame(self.notebook)
        self.batch_tab = self.create_modern_frame(self.notebook)

        self.notebook.add(self.pdf_tab, text="PDF选择")
        self.notebook.add(self.excel_tab, text="数据导入")
        self.notebook.add(self.watermark_tab, text="水印设置")
        self.notebook.add(self.email_tab, text="邮件设置")
        self.notebook.add(self.batch_tab, text="批量处理")

        # 设置各选项卡内容
        self.setup_pdf_tab()
        self.setup_excel_tab()
        self.setup_watermark_tab()
        self.setup_email_tab()
        self.setup_batch_tab()

        # 底部状态栏 - 现代化设计
        status_container = tk.Frame(self.master, bg='#F8F9FA', height=40)
        status_container.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(0, 15))
        
        self.status_bar = ttk.Label(status_container, text="就绪", relief=tk.FLAT, 
                                   style='Heading.TLabel')
        self.status_bar.pack(side=tk.LEFT, padx=10, pady=10)

        # 进度条 - 现代化设计
        progress_container = tk.Frame(self.master, bg='#F8F9FA')
        progress_container.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(0, 10))
        
        self.progress = ttk.Progressbar(progress_container, orient=tk.HORIZONTAL, 
                                      length=100, mode='determinate', style='TProgressbar')
        self.progress.pack(side=tk.BOTTOM, fill=tk.X)

    def create_modern_frame(self, parent):
        """创建现代化框架"""
        frame = ttk.Frame(parent, style='TFrame')
        return frame

    def setup_pdf_tab(self):
        # 左侧控制面板
        left_frame = self.create_modern_frame(self.pdf_tab)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # 右侧预览区域
        right_frame = self.create_modern_frame(self.pdf_tab)
        right_frame.grid(row=0, column=1, sticky="nsew")

        self.pdf_tab.grid_columnconfigure(1, weight=4)  # 预览区域占更多空间
        self.pdf_tab.grid_columnconfigure(0, weight=1)
        self.pdf_tab.grid_rowconfigure(0, weight=1)

        # 左侧控制区域 - 现代化设计
        control_group = ttk.LabelFrame(left_frame, text="PDF文件控制", padding="15")
        control_group.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(control_group, text="选择PDF文件:").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Button(control_group, text="浏览...", command=self.load_pdf).grid(row=0, column=1, sticky="e", pady=8)

        self.pdf_path_label = ttk.Label(control_group, text="未选择文件", style='TLabel')
        self.pdf_path_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Separator(control_group, orient=tk.HORIZONTAL).grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)

        ttk.Label(control_group, text="PDF页数:").grid(row=3, column=0, sticky="w", pady=5)
        self.page_count_label = ttk.Label(control_group, text="0", style='TLabel')
        self.page_count_label.grid(row=3, column=1, sticky="e", pady=5)

        ttk.Label(control_group, text="预览页面:").grid(row=4, column=0, sticky="w", pady=5)
        self.preview_page = tk.StringVar(value="1")
        preview_spin = ttk.Spinbox(control_group, from_=1, to=1, textvariable=self.preview_page, width=5,
                                   command=self.update_preview)
        preview_spin.grid(row=4, column=1, sticky="e", pady=5)
        self.preview_spin = preview_spin

        ttk.Button(control_group, text="预览水印效果", command=self.preview_watermark).grid(row=5, column=0,
                                                                                         columnspan=2, pady=15)

        # 右侧预览区域
        self.setup_preview_area(right_frame)

    def setup_excel_tab(self):
        frame = self.create_modern_frame(self.excel_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Excel文件选择 - 现代化分组
        excel_group = ttk.LabelFrame(frame, text="Excel数据导入", padding="15")
        excel_group.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(excel_group, text="选择Excel文件:").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Button(excel_group, text="浏览...", command=self.load_excel).grid(row=0, column=1, sticky="w", pady=8)

        self.excel_path_label = ttk.Label(excel_group, text="未选择文件", style='TLabel')
        self.excel_path_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=5)

        # Excel列选择 - 现代化分组
        column_group = ttk.LabelFrame(frame, text="列映射设置", padding="15")
        column_group.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(column_group, text="选择公司名称列:").grid(row=0, column=0, sticky="w", pady=8)
        self.column_combobox = ttk.Combobox(column_group, state="readonly", width=15)
        self.column_combobox.grid(row=0, column=1, sticky="w", pady=8)
        self.column_combobox.bind("<<ComboboxSelected>>", self.load_company_names)

        ttk.Button(column_group, text="刷新公司列表", command=self.load_company_names).grid(row=0, column=2, sticky="w",
                                                                                     pady=8)

        # 新增：邮箱列选择
        ttk.Label(column_group, text="选择公司邮箱列:").grid(row=1, column=0, sticky="w", pady=8)
        self.email_column_combobox = ttk.Combobox(column_group, state="readonly", width=15)
        self.email_column_combobox.grid(row=1, column=1, sticky="w", pady=8)
        self.email_column_combobox.bind("<<ComboboxSelected>>", self.load_company_emails)

        ttk.Button(column_group, text="刷新邮箱列表", command=self.load_company_emails).grid(row=1, column=2, sticky="w",
                                                                                      pady=8)

        # 水印文字模板 - 现代化分组
        template_group = ttk.LabelFrame(frame, text="水印模板设置", padding="15")
        template_group.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(template_group, text="水印文字模板:").grid(row=0, column=0, sticky="w", pady=8)

        template_frame = self.create_modern_frame(template_group)
        template_frame.grid(row=0, column=1, columnspan=2, sticky="w", pady=8)

        # 直接使用__init__中已初始化的变量，不再重新创建
        ttk.Label(template_frame, text="前缀:").pack(side=tk.LEFT, padx=(0, 8))
        prefix_entry = ttk.Entry(template_frame, textvariable=self.prefix_text, width=15)
        prefix_entry.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(template_frame, text="[公司名]").pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(template_frame, text="后缀:").pack(side=tk.LEFT, padx=(0, 8))
        suffix_entry = ttk.Entry(template_frame, textvariable=self.suffix_text, width=15)
        suffix_entry.pack(side=tk.LEFT)

        # 公司名称列表 - 现代化分组
        list_group = ttk.LabelFrame(frame, text="公司列表", padding="15")
        list_group.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # 创建带滚动条的列表框
        list_frame = self.create_modern_frame(list_group)
        list_frame.pack(fill=tk.BOTH, expand=True)
        list_group.grid_rowconfigure(0, weight=1)
        list_group.grid_columnconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.company_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=15, width=40,
                                         bg='white', relief=tk.FLAT, borderwidth=1, font=('Segoe UI', 9))
        self.company_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.company_listbox.yview)

        # 添加使用说明 - 现代化分组
        info_group = ttk.LabelFrame(frame, text="使用说明", padding="15")
        info_group.pack(fill=tk.X, pady=(0, 15))

        usage_info = """Excel表格格式要求：
• 公司名称列：每行一个公司名称
• 邮箱地址列：支持多邮箱，用分号(;)分隔

示例：
公司A    admin@companyA.com;sales@companyA.com
公司B    contact@companyB.com
公司C    info@companyC.com;support@companyC.com;tech@companyC.com

导入后，公司列表将显示：
• 公司A (admin@companyA.com; sales@companyA.com)
• 公司B (contact@companyB.com)  
• 公司C (info@companyC.com; support@companyC.com... 等3个)

系统将为每个公司生成相同的水印PDF，并发送给该公司的所有邮箱地址。"""

        ttk.Label(info_group, text=usage_info, justify=tk.LEFT, font=('Segoe UI', 9)).pack(anchor="w")

    def setup_watermark_tab(self):
        frame = self.create_modern_frame(self.watermark_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # 创建现代化的滚动区域
        canvas = tk.Canvas(frame, bg='#F8F9FA', highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')

        def update_scrollregion(event=None):
            """防抖更新滚动区域"""
            if hasattr(self, '_scrollregion_timer'):
                self.master.after_cancel(self._scrollregion_timer)
            self._scrollregion_timer = self.master.after(100, lambda: canvas.configure(scrollregion=canvas.bbox("all")))
        
        scrollable_frame.bind("<Configure>", update_scrollregion)

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 水印透明度 - 现代化分组
        opacity_group = ttk.LabelFrame(scrollable_frame, text="透明度设置", padding="15")
        opacity_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(opacity_group, text="水印透明度:").grid(row=0, column=0, sticky="w", pady=5)
        self.opacity_scale = ttk.Scale(opacity_group, from_=0, to=100, orient=tk.HORIZONTAL, length=200)
        self.opacity_scale.grid(row=0, column=1, sticky="w", pady=5)
        self.opacity_value = ttk.Label(opacity_group, text="26%", style='TLabel')
        self.opacity_value.grid(row=0, column=2, sticky="w", pady=5)
        self.opacity_scale.bind("<Motion>", self.update_opacity_value)
        self.opacity_scale.bind("<ButtonRelease-1>", self.update_opacity_value)

        # 水印角度 - 现代化分组
        angle_group = ttk.LabelFrame(scrollable_frame, text="旋转角度", padding="15")
        angle_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(angle_group, text="水印角度:").grid(row=0, column=0, sticky="w", pady=5)
        angle_frame = self.create_modern_frame(angle_group)
        angle_frame.grid(row=0, column=1, sticky="w", pady=5)
        ttk.Radiobutton(angle_frame, text="0°", variable=self.watermark_angle, value=0,
                        command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(angle_frame, text="30°", variable=self.watermark_angle, value=30,
                        command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(angle_frame, text="45°", variable=self.watermark_angle, value=45,
                        command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)
        ttk.Radiobutton(angle_frame, text="90°", variable=self.watermark_angle, value=90,
                        command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 字体大小 - 现代化分组
        size_group = ttk.LabelFrame(scrollable_frame, text="字体设置", padding="15")
        size_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(size_group, text="字体大小:").grid(row=0, column=0, sticky="w", pady=5)
        size_frame = self.create_modern_frame(size_group)
        size_frame.grid(row=0, column=1, sticky="w", pady=5)
        sizes = [36, 48, 60, 72]  # 增大字体大小选项
        for size in sizes:
            ttk.Radiobutton(size_frame, text=str(size), variable=self.font_size, value=size,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 字体选择 - 继续字体设置
        ttk.Label(size_group, text="字体选择:").grid(row=1, column=0, sticky="w", pady=5)
        # 只保留宽体字体
        fonts = ["黑体", "微软雅黑", "微软雅黑粗体", "Arial Black", "Impact"]
        available_fonts = [font for font in fonts if font in self.system_fonts]
        if not available_fonts:
            available_fonts = ["默认字体"]

        font_combo = ttk.Combobox(size_group, textvariable=self.font_family, values=available_fonts, state="readonly",
                                  width=15)
        font_combo.grid(row=1, column=1, sticky="w", pady=5)
        font_combo.bind("<<ComboboxSelected>>", lambda e: self.update_preview_on_change())
        font_combo.current(0)  # 默认选择第一个字体

        # 字体颜色 - 现代化分组
        color_group = ttk.LabelFrame(scrollable_frame, text="颜色设置", padding="15")
        color_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(color_group, text="字体颜色:").grid(row=0, column=0, sticky="w", pady=5)
        color_frame = self.create_modern_frame(color_group)
        color_frame.grid(row=0, column=1, sticky="w", pady=5)

        # 当前选中的颜色显示
        self.color_indicator = tk.Frame(color_frame, width=20, height=20, bg=self.text_color,
                                        relief=tk.RAISED, borderwidth=2)
        self.color_indicator.pack(side=tk.LEFT, padx=8)

        self.color_btn = tk.Button(color_frame, text="选择颜色", width=8, command=self.choose_color,
                                  bg=self.ui.colors['secondary'], fg='white', relief=tk.FLAT, font=('Segoe UI', 9))
        self.color_btn.pack(side=tk.LEFT, padx=8)

        # 常用颜色快捷按钮 - 只保留红色、蓝色和灰色
        self.color_buttons = []  # 存储颜色按钮以便更新选中状态
        colors = [("#FF0000", "红色"), ("#0000FF", "蓝色"), ("#808080", "灰色")]  # 只保留三种颜色

        for color_code, color_name in colors:
            btn_frame = tk.Frame(color_frame)
            btn_frame.pack(side=tk.LEFT, padx=4)

            # 创建颜色按钮
            btn = tk.Button(btn_frame, text=color_name, width=6, bg="white", fg="black",
                            command=lambda c=color_code: self.set_color(c), relief=tk.FLAT, font=('Segoe UI', 8))
            btn.pack(side=tk.TOP)

            # 创建颜色指示器
            indicator = tk.Frame(btn_frame, width=35, height=5, bg=color_code)
            indicator.pack(side=tk.BOTTOM, fill=tk.X)

            # 存储按钮引用
            self.color_buttons.append((btn, color_code))

        # 高级效果类型 - 现代化分组
        effect_group = ttk.LabelFrame(scrollable_frame, text="高级效果", padding="15")
        effect_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(effect_group, text="效果类型:").grid(row=0, column=0, sticky="w", pady=5)
        effect_frame = self.create_modern_frame(effect_group)
        effect_frame.grid(row=0, column=1, sticky="w", pady=5)
        effects = [("轮廓效果", "outline"), ("阴影效果", "shadow"), ("浮雕效果", "emboss"), ("纹理效果", "texture")]
        for text, value in effects:
            ttk.Radiobutton(effect_frame, text=text, variable=self.effect_type, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 效果强度
        ttk.Label(effect_group, text="效果强度:").grid(row=1, column=0, sticky="w", pady=5)
        intensity_scale = ttk.Scale(effect_group, from_=0, to=100, orient=tk.HORIZONTAL, length=200)
        intensity_scale.grid(row=1, column=1, sticky="w", pady=5)
        self.intensity_value = ttk.Label(effect_group, text="70%", style='TLabel')
        self.intensity_value.grid(row=1, column=2, sticky="w", pady=5)
        intensity_scale.config(variable=self.effect_intensity)
        intensity_scale.bind("<Motion>", self.update_intensity_value)
        intensity_scale.bind("<ButtonRelease-1>", self.update_preview_on_change)

        # 轮廓宽度
        ttk.Label(effect_group, text="轮廓宽度:").grid(row=2, column=0, sticky="w", pady=5)
        outline_frame = self.create_modern_frame(effect_group)
        outline_frame.grid(row=2, column=1, sticky="w", pady=5)
        outline_options = [("细", 1), ("中", 2), ("粗", 3), ("特粗", 5)]
        for text, value in outline_options:
            ttk.Radiobutton(outline_frame, text=text, variable=self.outline_width, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 阴影偏移
        ttk.Label(effect_group, text="阴影偏移:").grid(row=3, column=0, sticky="w", pady=5)
        shadow_frame = self.create_modern_frame(effect_group)
        shadow_frame.grid(row=3, column=1, sticky="w", pady=5)
        shadow_options = [("小", 2), ("中", 3), ("大", 5)]
        for text, value in shadow_options:
            ttk.Radiobutton(shadow_frame, text=text, variable=self.shadow_offset, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 图案密度
        ttk.Label(effect_group, text="图案密度:").grid(row=4, column=0, sticky="w", pady=5)
        pattern_frame = self.create_modern_frame(effect_group)
        pattern_frame.grid(row=4, column=1, sticky="w", pady=5)
        pattern_options = [("稀疏", 3), ("中等", 5), ("密集", 8)]
        for text, value in pattern_options:
            ttk.Radiobutton(pattern_frame, text=text, variable=self.pattern_density, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 水印密度 - 使用5档密度设置，最高对应10x10
        ttk.Label(effect_group, text="水印密度:").grid(row=5, column=0, sticky="w", pady=5)
        density_frame = self.create_modern_frame(effect_group)
        density_frame.grid(row=5, column=1, sticky="w", pady=5)

        # 使用5档密度设置
        density_options = [
            ("低(1×1)", 1),
            ("中(2×2)", 2),
            ("高(4×4)", 4),
            ("超高(7×7)", 7),
            ("极高(10×10)", 10)
        ]

        for text, value in density_options:
            ttk.Radiobutton(density_frame, text=text, variable=self.watermark_density, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=5)

        # 水印位置
        ttk.Label(effect_group, text="水印位置:").grid(row=6, column=0, sticky="w", pady=5)
        position_frame = self.create_modern_frame(effect_group)
        position_frame.grid(row=6, column=1, sticky="w", pady=5)
        positions = [("居中", "center"), ("平铺", "tile")]
        for text, value in positions:
            ttk.Radiobutton(position_frame, text=text, variable=self.watermark_position, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 转换质量
        ttk.Label(effect_group, text="图片转换质量:").grid(row=7, column=0, sticky="w", pady=5)
        quality_frame = self.create_modern_frame(effect_group)
        quality_frame.grid(row=7, column=1, sticky="w", pady=5)
        qualities = [(f"{q}dpi", q) for q in [100, 200, 300]]
        for text, value in qualities:
            ttk.Radiobutton(quality_frame, text=text, variable=self.conversion_quality, value=value,
                            command=self.update_preview_on_change).pack(side=tk.LEFT, padx=8)

        # 设置为默认按钮
        ttk.Button(scrollable_frame, text="设置为默认", command=self.save_default_settings).pack(pady=15)

    def setup_batch_tab(self):
        frame = self.create_modern_frame(self.batch_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # 输出设置 - 现代化分组
        output_group = ttk.LabelFrame(frame, text="输出设置", padding="15")
        output_group.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(output_group, text="输出文件夹:").grid(row=0, column=0, sticky="w", pady=8)
        self.output_dir = tk.StringVar()
        ttk.Entry(output_group, textvariable=self.output_dir, width=40).grid(row=0, column=1, sticky="w", pady=8)
        ttk.Button(output_group, text="浏览...", command=self.select_output_dir).grid(row=0, column=2, sticky="w", pady=8)

        # 文件命名规则
        ttk.Label(output_group, text="文件命名规则:").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(output_group, textvariable=self.filename_pattern, width=40).grid(row=1, column=1, sticky="w", pady=8)
        ttk.Label(output_group, text="文件名表示上传文件名").grid(row=1, column=2, sticky="w", pady=8)

        # 图片化处理选项
        ttk.Label(output_group, text="图片化处理:").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Checkbutton(output_group, text="启用PDF图片化处理(防止去水印)", variable=self.enable_rasterize).grid(row=2,
                                                                                                          column=1,
                                                                                                          sticky="w",
                                                                                                          pady=8)

        # 添加PDF压缩选项
        ttk.Label(output_group, text="PDF压缩级别:").grid(row=3, column=0, sticky="w", pady=8)
        compression_frame = self.create_modern_frame(output_group)
        compression_frame.grid(row=3, column=1, sticky="w", pady=8)
        compression_options = [
            ("不压缩", 0),
            ("轻度压缩", 1),
            ("中度压缩", 2),
            ("高度压缩", 3)
        ]
        for text, value in compression_options:
            ttk.Radiobutton(compression_frame, text=text, variable=self.compression_level, value=value).pack(
                side=tk.LEFT, padx=8)

        # 邮件发送设置显示
        email_display_frame = ttk.LabelFrame(frame, text="邮件发送状态", padding="15")
        email_display_frame.pack(fill=tk.X, pady=(0, 10))

        self.email_status_label = ttk.Label(email_display_frame, text="邮件发送：未启用", style='TLabel')
        self.email_status_label.pack(side=tk.LEFT)

        ttk.Button(email_display_frame, text="邮件设置",
                   command=lambda: self.notebook.select(3)).pack(side=tk.RIGHT)

        # 预览与处理按钮 - 现代化按钮组
        button_group = ttk.LabelFrame(frame, text="操作控制", padding="15")
        button_group.pack(fill=tk.X, pady=(0, 10))

        button_frame = self.create_modern_frame(button_group)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="批量处理", command=self.batch_process).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="打开输出文件夹", command=self.open_output_folder).pack(side=tk.LEFT, padx=(0, 10))

        # 字体和系统信息
        system_info = f"系统: {platform.system()} {platform.version()}"
        font_info = f"已加载字体数: {len(self.system_fonts)}"
        ttk.Label(button_frame, text=f"{system_info} | {font_info}", style='TLabel').pack(side=tk.LEFT, padx=(20, 0))

        # 处理日志 - 现代化分组
        log_group = ttk.LabelFrame(frame, text="处理日志", padding="15")
        log_group.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        log_frame = self.create_modern_frame(log_group)
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_group.grid_rowconfigure(0, weight=1)
        log_group.grid_columnconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, height=15,
                               bg='white', relief=tk.FLAT, borderwidth=1, font=('Consolas', 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 优化日志文本框性能
        self.log_text.config(spacing1=2, spacing2=2, spacing3=2)  # 减少行间距

        scrollbar.config(command=self.log_text.yview)

    def setup_email_tab(self):
        """设置邮件选项卡"""
        frame = self.create_modern_frame(self.email_tab)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # 邮件发送开关 - 现代化分组
        enable_frame = ttk.LabelFrame(frame, text="邮件发送设置", padding="15")
        enable_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(enable_frame, text="启用邮件发送功能", variable=self.enable_email,
                        command=self.update_email_status).pack(anchor="w")

        # SMTP服务器设置 - 现代化分组
        smtp_frame = ttk.LabelFrame(frame, text="SMTP服务器设置", padding="15")
        smtp_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(smtp_frame, text="SMTP服务器:").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(smtp_frame, textvariable=self.smtp_server, width=25).grid(row=0, column=1, sticky="w", pady=8)
        ttk.Label(smtp_frame, text="(企业微信邮箱: smtp.exmail.qq.com)").grid(row=0, column=2, sticky="w", pady=8)

        ttk.Label(smtp_frame, text="端口:").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(smtp_frame, textvariable=self.smtp_port, width=25).grid(row=1, column=1, sticky="w", pady=8)
        ttk.Label(smtp_frame, text="(SSL: 465, TLS: 587)").grid(row=1, column=2, sticky="w", pady=8)

        ttk.Label(smtp_frame, text="邮箱账号:").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(smtp_frame, textvariable=self.smtp_username, width=25).grid(row=2, column=1, sticky="w", pady=8)
        ttk.Label(smtp_frame, text="(完整邮箱地址)").grid(row=2, column=2, sticky="w", pady=8)

        ttk.Label(smtp_frame, text="邮箱密码:").grid(row=3, column=0, sticky="w", pady=8)
        password_entry = ttk.Entry(smtp_frame, textvariable=self.smtp_password, width=25, show="*")
        password_entry.grid(row=3, column=1, sticky="w", pady=8)
        ttk.Label(smtp_frame, text="(授权码或密码)").grid(row=3, column=2, sticky="w", pady=8)

        ttk.Label(smtp_frame, text="发件人姓名:").grid(row=4, column=0, sticky="w", pady=8)
        ttk.Entry(smtp_frame, textvariable=self.sender_name, width=25).grid(row=4, column=1, sticky="w", pady=8)

        # 邮件内容设置 - 现代化分组
        content_frame = ttk.LabelFrame(frame, text="邮件内容设置", padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        ttk.Label(content_frame, text="邮件主题:").grid(row=0, column=0, sticky="nw", pady=8)
        subject_entry = ttk.Entry(content_frame, textvariable=self.email_subject, width=60)
        subject_entry.grid(row=0, column=1, sticky="ew", pady=8)

        ttk.Label(content_frame, text="邮件内容:").grid(row=1, column=0, sticky="nw", pady=8)

        # 创建Text widget并绑定到StringVar
        self.body_text = tk.Text(content_frame, width=60, height=8, wrap=tk.WORD,
                                bg='white', relief=tk.FLAT, borderwidth=1, font=('Segoe UI', 9))
        self.body_text.grid(row=1, column=1, sticky="nsew", pady=8)
        self.body_text.insert("1.0", self.email_body.get())

        # 绑定内容变化事件
        def update_email_body(event=None):
            self.email_body.set(self.body_text.get("1.0", tk.END).strip())

        self.body_text.bind("<KeyRelease>", update_email_body)
        self.body_text.bind("<Button-1>", update_email_body)

        content_frame.grid_columnconfigure(1, weight=1)
        content_frame.grid_rowconfigure(1, weight=1)

        # 测试邮件按钮 - 现代化按钮组
        test_group = ttk.LabelFrame(frame, text="邮件测试", padding="15")
        test_group.pack(fill=tk.X, pady=(0, 10))

        test_frame = self.create_modern_frame(test_group)
        test_frame.pack(fill=tk.X)

        ttk.Button(test_frame, text="测试邮件设置", command=self.test_email_settings).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(test_frame, text="保存邮件设置", command=self.save_email_settings).pack(side=tk.LEFT, padx=(0, 10))

        # 添加预设模板按钮
        ttk.Button(test_frame, text="加载邮件模板", command=self.load_email_template).pack(side=tk.LEFT, padx=(0, 10))

        # 说明文字 - 现代化分组
        info_group = ttk.LabelFrame(frame, text="使用说明", padding="15")
        info_group.pack(fill=tk.X, pady=(0, 10))

        info_text = """邮件内容支持以下变量：
{company} - 公司名称
{filename} - 文件名

多邮箱支持：
在Excel邮箱列中，一个公司可对应多个邮箱，用分号(;)分隔
例如：admin@company.com;sales@company.com;support@company.com

企业微信邮箱设置：
- SMTP服务器：smtp.exmail.qq.com
- 端口：465 (SSL) 或 587 (TLS)
- 需要在企业微信邮箱开启SMTP功能并获取授权码

Gmail邮箱优化：
已针对Gmail邮箱优化邮件主题显示，确保主题正确呈现。

请确保Excel中包含有效的邮箱地址列，并在"数据导入"选项卡中选择正确的邮箱列。"""

        ttk.Label(info_group, text=info_text, justify=tk.LEFT, font=('Segoe UI', 9)).pack(anchor="w")

    def setup_preview_area(self, parent):
        # 现代化预览区域
        preview_container = ttk.LabelFrame(parent, text="PDF预览", padding="15")
        preview_container.pack(fill=tk.BOTH, expand=True)

        self.preview_canvas = tk.Canvas(preview_container, bg="white", highlightthickness=0,
                                       relief=tk.FLAT, borderwidth=1)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)

        # 底部信息栏
        info_frame = self.create_modern_frame(preview_container)
        info_frame.pack(fill=tk.X, pady=(10, 0))

        scrollbar_y = ttk.Scrollbar(preview_container, orient="vertical", command=self.preview_canvas.yview)
        scrollbar_x = ttk.Scrollbar(preview_container, orient="horizontal", command=self.preview_canvas.xview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.preview_canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.preview_canvas.bind('<Configure>', self.on_canvas_configure)

        # 预览标签
        self.preview_label = ttk.Label(info_frame, text="无预览", style='TLabel')
        self.preview_label.pack(side=tk.LEFT)

    def on_canvas_configure(self, event):
        """防抖更新预览画布滚动区域"""
        if hasattr(self, 'preview_image') and self.preview_image:
            if hasattr(self, '_preview_timer'):
                self.master.after_cancel(self._preview_timer)
            self._preview_timer = self.master.after(100, lambda: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all")))

    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.pdf_path_label.config(text=os.path.basename(file_path))

            # 读取PDF页数
            try:
                with open(file_path, 'rb') as f:
                    pdf = PdfReader(f)
                    page_count = len(pdf.pages)
                    self.page_count_label.config(text=str(page_count))

                    # 更新预览页面选择器
                    self.preview_spin.config(from_=1, to=page_count)

                    # 显示第一页预览
                    self.update_preview()

                self.log(f"已加载PDF: {os.path.basename(file_path)}，共 {page_count} 页")
            except Exception as e:
                messagebox.showerror("错误", f"无法读取PDF文件: {str(e)}")
                self.log(f"错误: {str(e)}")

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_path = file_path
            self.excel_path_label.config(text=os.path.basename(file_path))

            try:
                # 读取Excel文件
                df = pd.read_excel(file_path)

                # 获取列名
                columns = df.columns.tolist()

                # 更新列选择下拉框
                self.column_combobox['values'] = columns
                self.email_column_combobox['values'] = columns
                if columns:
                    self.column_combobox.current(0)
                    if len(columns) > 1:
                        self.email_column_combobox.current(1)
                    else:
                        self.email_column_combobox.current(0)

                self.log(f"已加载Excel: {os.path.basename(file_path)}，包含 {len(columns)} 列")

                # 加载公司名称和邮箱
                self.load_company_names()
                self.load_company_emails()

            except Exception as e:
                messagebox.showerror("错误", f"无法读取Excel文件: {str(e)}")
                self.log(f"错误: {str(e)}")

    def load_company_names(self, event=None):
        if not self.excel_path:
            return

        selected_column = self.column_combobox.get()
        if not selected_column:
            return

        try:
            # 读取选定列的公司名称
            df = pd.read_excel(self.excel_path)
            company_names = df[selected_column].dropna().tolist()
            self.company_names = company_names

            # 更新公司列表显示（包含邮箱信息）
            self.update_company_list_display()

            self.log(f"已加载 {len(company_names)} 个公司名称")

            # 更新邮件状态显示
            self.update_email_status()

        except Exception as e:
            messagebox.showerror("错误", f"无法加载公司名称: {str(e)}")
            self.log(f"错误: {str(e)}")

    def update_company_list_display(self):
        """更新公司名称列表显示，整合邮箱信息"""
        # 清空列表
        self.company_listbox.delete(0, tk.END)

        # 重新填充列表，包含邮箱信息
        for company_name in self.company_names:
            display_text = company_name

            # 如果有邮箱映射，添加邮箱信息
            if hasattr(self, 'company_email_map') and company_name in self.company_email_map:
                emails = self.company_email_map[company_name]
                if emails:
                    if len(emails) == 1:
                        display_text += f" ({emails[0]})"
                    else:
                        # 多个邮箱时显示数量和前两个邮箱
                        email_preview = f"{emails[0]}; {emails[1]}" if len(emails) > 1 else emails[0]
                        if len(emails) > 2:
                            display_text += f" ({email_preview}... 等{len(emails)}个)"
                        else:
                            display_text += f" ({email_preview})"
                else:
                    display_text += " (无邮箱)"

            self.company_listbox.insert(tk.END, display_text)

    def load_company_emails(self, event=None):
        """加载公司邮箱列表，支持一对多关系（分号分隔）"""
        if not self.excel_path:
            return

        selected_column = self.email_column_combobox.get()
        if not selected_column:
            return

        try:
            # 读取选定列的公司邮箱
            df = pd.read_excel(self.excel_path)
            company_emails_raw = df[selected_column].dropna().tolist()

            # 清空之前的映射
            self.company_email_map = {}
            all_valid_emails = []
            total_email_count = 0
            invalid_emails = []

            for i, email_cell in enumerate(company_emails_raw):
                if i < len(self.company_names):
                    company_name = self.company_names[i]

                    # 解析分号分隔的邮箱地址，支持中文全角和英文半角分号
                    email_list = []
                    if pd.notna(email_cell):
                        # 先将中文全角分号替换为英文半角分号，然后统一分割
                        email_cell_normalized = str(email_cell).replace('；', ';')
                        emails = email_cell_normalized.split(';')
                        for email in emails:
                            email_str = email.strip()
                            if email_str and self.is_valid_email(email_str):
                                email_list.append(email_str)
                                all_valid_emails.append(email_str)
                                total_email_count += 1
                            elif email_str:
                                invalid_emails.append(email_str)

                    # 建立公司名称到邮箱列表的映射
                    self.company_email_map[company_name] = email_list

            self.company_emails = all_valid_emails

            # 更新公司列表显示（现在包含邮箱信息）
            self.update_company_list_display()

            if invalid_emails:
                self.log(
                    f"发现 {len(invalid_emails)} 个无效邮箱地址，已自动过滤: {', '.join(invalid_emails[:5])}{'...' if len(invalid_emails) > 5 else ''}")

            self.log(
                f"已加载 {total_email_count} 个有效邮箱地址，覆盖 {len([k for k, v in self.company_email_map.items() if v])} 个公司")

            # 显示详细的映射信息
            for company, emails in self.company_email_map.items():
                if emails:
                    self.log(
                        f"  {company}: {len(emails)} 个邮箱 ({', '.join(emails[:2])}{'...' if len(emails) > 2 else ''})")

            # 更新邮件状态显示
            self.update_email_status()

        except Exception as e:
            messagebox.showerror("错误", f"无法加载公司邮箱: {str(e)}")
            self.log(f"错误: {str(e)}")

    def is_valid_email(self, email):
        """验证邮箱地址格式 - 增强版RFC 5322合规检查"""
        if not email or not isinstance(email, str):
            return False
        
        email = email.strip()
        
        # 基本长度检查
        if len(email) > 254:  # RFC 5322规定最大长度
            return False
        
        # 更严格的正则表达式，符合RFC 5322标准
        import re
        pattern = r'^[a-zA-Z0-9!#$%&\'*+/=?^_`{|}~-]+(?:\.[a-zA-Z0-9!#$%&\'*+/=?^_`{|}~-]+)*@(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?$'
        
        if not re.match(pattern, email):
            return False
        
        # 额外检查：确保域名部分有效
        local_part, domain_part = email.split('@', 1)
        
        # 检查本地部分长度
        if len(local_part) > 64:  # RFC 5322规定本地部分最大长度
            return False
        
        # 检查域名部分
        if len(domain_part) > 253 or '.' not in domain_part:
            return False
        
        # 检查域名各部分
        domain_parts = domain_part.split('.')
        if any(len(part) == 0 or len(part) > 63 for part in domain_parts):
            return False
        
        return True

    def update_email_status(self):
        """更新邮件状态显示"""
        if hasattr(self, 'email_status_label'):
            if self.enable_email.get():
                total_emails = len(self.company_emails) if hasattr(self, 'company_emails') else 0
                company_count = len(self.company_names) if hasattr(self, 'company_names') else 0
                companies_with_email = len(
                    [k for k, v in getattr(self, 'company_email_map', {}).items() if v]) if hasattr(self,
                                                                                                    'company_email_map') else 0

                status_text = f"邮件发送：已启用 (总邮箱：{total_emails}, 公司数：{company_count}, 有邮箱公司：{companies_with_email})"

                if company_count > 0 and companies_with_email < company_count:
                    no_email_count = company_count - companies_with_email
                    status_text += f" [缺失邮箱：{no_email_count}个公司]"
            else:
                status_text = "邮件发送：未启用"
            self.email_status_label.config(text=status_text)

    def send_email(self, company_name, email_address, file_path, filename):
        """发送邮件，专门针对Gmail和其他邮箱优化"""
        try:
            # 创建邮件对象，使用mixed类型以确保Gmail兼容性
            msg = MIMEMultipart('mixed')

            # 设置基本邮件头 - 确保RFC 5322合规
            sender_name = self.sender_name.get().strip()
            sender_email = self.smtp_username.get().strip()
            
            # 验证邮箱地址格式
            if not self.is_valid_email(sender_email):
                raise Exception(f"发件人邮箱地址格式无效: {sender_email}")
            
            # 安全地设置From头部，处理特殊字符
            try:
                # 尝试直接设置（ASCII字符）
                sender_name.encode('ascii')
                # 如果是ASCII字符，可以直接使用
                msg['From'] = f"{sender_name} <{sender_email}>"
            except UnicodeEncodeError:
                # 如果包含非ASCII字符，需要使用RFC 2047编码
                from email.header import Header
                encoded_name = Header(sender_name, 'utf-8').encode()
                msg['From'] = f"{encoded_name} <{sender_email}>"
            
            msg['To'] = email_address

            # 处理邮件主题，确保Gmail能正确显示
            subject = self.email_subject.get().replace("{company}", company_name).replace("{filename}", filename)

            # Gmail特殊处理：确保主题编码正确
            try:
                # 先尝试ASCII编码
                subject.encode('ascii')
                msg['Subject'] = subject
            except UnicodeEncodeError:
                # 如果包含非ASCII字符，使用base64编码
                from email.header import Header
                msg['Subject'] = Header(subject, 'utf-8', header_name='Subject').encode()

            # 添加Gmail友好的邮件头
            import email.utils
            msg['Date'] = email.utils.formatdate(localtime=True)
            msg['Message-ID'] = email.utils.make_msgid()
            msg['X-Mailer'] = 'PDF-Watermark-Tool'
            msg['X-Priority'] = '3'
            msg['MIME-Version'] = '1.0'

            # 替换邮件内容中的变量
            body = self.email_body.get().replace("{company}", company_name).replace("{filename}", filename)

            # 创建文本部分，使用quoted-printable编码以提高Gmail兼容性
            from email.mime.text import MIMEText
            text_part = MIMEText(body, 'plain', 'utf-8')
            text_part.set_charset('utf-8')
            msg.attach(text_part)

            # 添加PDF附件，特殊处理确保Gmail能正确识别
            with open(file_path, 'rb') as f:
                pdf_data = f.read()
                attachment = MIMEApplication(pdf_data, _subtype='pdf', name=f"{filename}.pdf")
                attachment.add_header('Content-Disposition', 'attachment', filename=f"{filename}.pdf")
                attachment.add_header('Content-Type', 'application/pdf', name=f"{filename}.pdf")
                attachment.add_header('Content-Transfer-Encoding', 'base64')
                msg.attach(attachment)

            # 连接SMTP服务器
            server = None
            try:
                if self.smtp_port.get() == 465:  # SSL
                    server = smtplib.SMTP_SSL(self.smtp_server.get(), self.smtp_port.get())
                else:  # TLS (587 for Gmail)
                    server = smtplib.SMTP(self.smtp_server.get(), self.smtp_port.get())
                    server.starttls()

                # 登录
                server.login(self.smtp_username.get(), self.smtp_password.get())

                # 发送邮件，使用更兼容的方法
                from_addr = self.smtp_username.get()
                to_addr = [email_address]

                # 将邮件转换为字符串，确保编码正确
                msg_string = msg.as_string()

                # 发送邮件
                server.sendmail(from_addr, to_addr, msg_string.encode('utf-8'))

            finally:
                if server:
                    server.quit()

            return True

        except Exception as e:
            raise Exception(f"发送邮件失败: {str(e)}")

    def test_email_settings(self):
        """测试邮件设置"""
        if not self.smtp_username.get() or not self.smtp_password.get():
            messagebox.showerror("错误", "请先填写邮箱账号和密码")
            return
        
        # 验证邮箱地址格式
        if not self.is_valid_email(self.smtp_username.get()):
            messagebox.showerror("错误", f"发件人邮箱地址格式无效: {self.smtp_username.get()}")
            return
        
        # 验证发件人姓名
        sender_name = self.sender_name.get().strip()
        if not sender_name:
            messagebox.showerror("错误", "发件人姓名不能为空")
            return

        try:
            # 测试SMTP连接
            if self.smtp_port.get() == 465:  # SSL
                server = smtplib.SMTP_SSL(self.smtp_server.get(), self.smtp_port.get())
            else:  # TLS
                server = smtplib.SMTP(self.smtp_server.get(), self.smtp_port.get())
                server.starttls()

            server.login(self.smtp_username.get(), self.smtp_password.get())
            server.quit()

            messagebox.showinfo("成功", "邮件设置测试成功！SMTP连接正常。")
            self.log("邮件设置测试成功")

        except Exception as e:
            messagebox.showerror("错误", f"邮件设置测试失败: {str(e)}")
            self.log(f"邮件设置测试失败: {str(e)}")

    def save_email_settings(self):
        """保存邮件设置"""
        
        # 验证邮箱地址格式
        if not self.is_valid_email(self.smtp_username.get()):
            messagebox.showerror("错误", f"发件人邮箱地址格式无效: {self.smtp_username.get()}")
            return
        
        # 验证发件人姓名
        sender_name = self.sender_name.get().strip()
        if not sender_name:
            messagebox.showerror("错误", "发件人姓名不能为空")
            return
        
        # 验证SMTP服务器设置
        if not self.smtp_server.get().strip():
            messagebox.showerror("错误", "SMTP服务器地址不能为空")
            return
        
        if self.smtp_port.get() not in [465, 587, 25]:
            messagebox.showerror("错误", "SMTP端口必须是 465(SSL)、587(TLS) 或 25")
            return
        
        settings = {
            "smtp_server": self.smtp_server.get(),
            "smtp_port": self.smtp_port.get(),
            "smtp_username": self.smtp_username.get(),
            "smtp_password": self.smtp_password.get(),  # 注意：密码明文保存，生产环境建议加密
            "sender_name": self.sender_name.get(),
            "email_subject": self.email_subject.get(),
            "email_body": self.email_body.get(),
            "enable_email": self.enable_email.get()
        }

        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "邮件设置已保存")
            self.log("邮件设置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存邮件设置失败: {str(e)}")

    def load_email_settings(self):
        """加载邮件设置"""
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")
        if not os.path.exists(config_path):
            return

        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)

            self.smtp_server.set(settings.get("smtp_server", "smtp.exmail.qq.com"))
            self.smtp_port.set(settings.get("smtp_port", 465))
            self.smtp_username.set(settings.get("smtp_username", ""))
            self.smtp_password.set(settings.get("smtp_password", ""))
            self.sender_name.set(settings.get("sender_name", "系统管理员"))
            self.email_subject.set(settings.get("email_subject", "您的加水印文件已准备好"))
            self.email_body.set(settings.get("email_body",
                                             "尊敬的{company}，\n\n您的加水印文件已处理完成，请查收附件。\n\n如有任何问题，请及时联系我们。\n\n祝好！"))
            self.enable_email.set(settings.get("enable_email", False))

            # 更新Text widget内容
            if hasattr(self, 'body_text'):
                self.body_text.delete("1.0", tk.END)
                self.body_text.insert("1.0", self.email_body.get())

        except Exception as e:
            self.log(f"加载邮件设置失败: {str(e)}")

    def load_email_template(self):
        """加载邮件模板"""
        templates = {
            "商务正式": {
                "subject": "您的文档已完成处理",
                "body": "尊敬的{company}，\n\n您好！\n\n您委托处理的文档已完成加密水印处理，请查收附件。\n\n文件名：{filename}\n处理时间：" + time.strftime(
                    "%Y-%m-%d %H:%M:%S") + "\n\n如有任何疑问，请随时与我们联系。\n\n此致\n敬礼！"
            },
            "友好简洁": {
                "subject": "您的文件已准备好 - {filename}",
                "body": "Hi {company}，\n\n您的文件已处理完成！\n\n附件：{filename}\n\n如有问题随时联系我们哦～\n\n谢谢！"
            },
            "技术规范": {
                "subject": "文档处理完成通知 - {filename}",
                "body": "致：{company}\n\n文档处理状态：已完成\n文件名称：{filename}\n处理类型：水印添加\n处理时间：" + time.strftime(
                    "%Y-%m-%d %H:%M:%S") + "\n安全等级：高\n\n请检查附件并确认文档完整性。\n\n技术支持团队"
            }
        }

        # 创建模板选择窗口
        template_window = tk.Toplevel(self.master)
        template_window.title("选择邮件模板")
        template_window.geometry("400x300")
        template_window.transient(self.master)
        template_window.grab_set()
        
        # 应用现代化样式
        template_window.configure(bg='#F8F9FA')

        ttk.Label(template_window, text="请选择邮件模板：", style='Title.TLabel').pack(pady=15)

        # 模板列表
        template_frame = self.create_modern_frame(template_window)
        template_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        for template_name, template_data in templates.items():
            frame = ttk.LabelFrame(template_frame, text=template_name, padding="15")
            frame.pack(fill=tk.X, pady=8)

            ttk.Label(frame, text=f"主题：{template_data['subject']}", style='TLabel').pack(anchor="w")

            body_preview = template_data['body'][:100] + "..." if len(template_data['body']) > 100 else template_data[
                'body']
            ttk.Label(frame, text=f"内容：{body_preview}", wraplength=350, style='TLabel').pack(anchor="w")

            ttk.Button(frame, text="使用此模板",
                       command=lambda t=template_data: self.apply_email_template(t, template_window)).pack(anchor="e", pady=(8, 0))

        ttk.Button(template_window, text="取消", command=template_window.destroy).pack(pady=15)

    def apply_email_template(self, template_data, window):
        """应用邮件模板"""
        self.email_subject.set(template_data['subject'])
        self.email_body.set(template_data['body'])

        # 更新Text widget内容
        if hasattr(self, 'body_text'):
            self.body_text.delete("1.0", tk.END)
            self.body_text.insert("1.0", template_data['body'])

        window.destroy()
        messagebox.showinfo("成功", "邮件模板已应用")

    def update_preview(self):
        if not self.pdf_path:
            return

        try:
            # 获取选择的页面
            page_num = int(self.preview_page.get())

            # 使用当前设置的质量设置进行预览 - 修复预览质量问题
            quality = int(self.conversion_quality.get())

            # 将PDF页面转换为图像，使用与最终输出相同的质量设置
            images = convert_from_path(self.pdf_path, first_page=page_num, last_page=page_num, dpi=quality)
            if images:
                image = images[0]
                self.show_preview(image)
                self.preview_label.config(text=f"预览: 第 {page_num} 页 ({quality}dpi)")

            # 保留原有预览方式，不自动显示水印效果
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {str(e)}")
            self.log(f"错误: {str(e)}")

    def show_preview(self, image):
        self.preview_canvas.delete("all")

        # 调整图像大小以适应Canvas
        self.master.update_idletasks()
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()

        # 保存原始图像用于水印预览
        self.preview_image = image

        # 调整图像大小
        image_ratio = image.width / image.height
        canvas_ratio = canvas_width / canvas_height

        if image_ratio > canvas_ratio:
            new_width = canvas_width
            new_height = int(canvas_width / image_ratio)
        else:
            new_height = canvas_height
            new_width = int(canvas_height * image_ratio)

        resized_image = image.copy()
        resized_image.thumbnail((new_width, new_height), Image.LANCZOS)

        # 显示预览图像
        self.photo = ImageTk.PhotoImage(resized_image)
        self.preview_canvas.create_image(canvas_width // 2, canvas_height // 2, anchor="center", image=self.photo)
        self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))

    def update_preview_on_change(self):
        """当设置改变时更新预览"""
        if hasattr(self, 'preview_image') and self.preview_image:
            # 使用延迟更新，避免频繁刷新
            self.master.after(300, self.preview_watermark)

    def preview_watermark(self):
        if not self.preview_image:
            messagebox.showinfo("提示", "请先选择PDF文件")
            return

        sample_company = "示例公司"

        if self.company_names and len(self.company_names) > 0:
            sample_company = self.company_names[0]

        # 生成水印文本
        watermark_text = f"{self.prefix_text.get()}{sample_company}{self.suffix_text.get()}"

        # 从原始预览图像创建副本
        image = self.preview_image.copy()

        # 添加水印
        watermarked_image = self.add_text_watermark_to_image(
            image,
            watermark_text,
            opacity=int(self.opacity_scale.get()) / 100,
            angle=self.watermark_angle.get(),
            font_size=self.font_size.get(),
            font_family=self.font_family.get(),
            color=self.text_color,
            density=self.watermark_density.get(),
            position=self.watermark_position.get(),
            effect_type=self.effect_type.get(),
            outline_width=self.outline_width.get(),
            shadow_offset=self.shadow_offset.get(),
            effect_intensity=self.effect_intensity.get(),
            pattern_density=self.pattern_density.get()
        )

        # 显示水印预览
        self.show_preview(watermarked_image)

        # 简化预览标签，邮箱信息已在公司列表中显示
        self.preview_label.config(text=f"水印预览: {watermark_text}")

    # 保存默认设置
    def save_default_settings(self):
        """保存当前水印设置为默认设置"""
        settings = {
            "opacity": self.opacity_scale.get(),
            "angle": self.watermark_angle.get(),
            "font_size": self.font_size.get(),
            "font_family": self.font_family.get(),
            "text_color": self.text_color,
            "effect_type": self.effect_type.get(),
            "effect_intensity": self.effect_intensity.get(),
            "outline_width": self.outline_width.get(),
            "shadow_offset": self.shadow_offset.get(),
            "pattern_density": self.pattern_density.get(),
            "watermark_density": self.watermark_density.get(),
            "watermark_position": self.watermark_position.get(),
            "conversion_quality": self.conversion_quality.get(),
            "compression_level": self.compression_level.get(),
            "filename_pattern": self.filename_pattern.get(),
            "enable_rasterize": self.enable_rasterize.get()
        }

        # 保存到配置文件
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "watermark_config.json")
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "默认设置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存默认设置失败: {str(e)}")

    # 加载默认设置
    def load_default_settings(self):
        """加载保存的默认设置"""
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "watermark_config.json")
        if not os.path.exists(config_path):
            # 如果配置文件不存在，使用默认值
            self.apply_default_values()
            return

        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)

            # 应用加载的设置到UI
            self.opacity_scale.set(settings.get("opacity", 24))
            self.watermark_angle.set(settings.get("angle", 45))
            self.font_size.set(settings.get("font_size", 48))
            self.font_family.set(settings.get("font_family", "黑体"))
            self.text_color = settings.get("text_color", "#FF0000")
            self.effect_type.set(settings.get("effect_type", "outline"))
            self.effect_intensity.set(settings.get("effect_intensity", 70))
            self.outline_width.set(settings.get("outline_width", 2))
            self.shadow_offset.set(settings.get("shadow_offset", 3))
            self.pattern_density.set(settings.get("pattern_density", 5))
            self.watermark_density.set(settings.get("watermark_density", 7))
            self.watermark_position.set(settings.get("watermark_position", "tile"))
            self.conversion_quality.set(settings.get("conversion_quality", 200))
            self.compression_level.set(settings.get("compression_level", 2))
            self.filename_pattern.set(settings.get("filename_pattern", "文件名{company}"))
            self.enable_rasterize.set(settings.get("enable_rasterize", True))

            # 更新UI显示
            self.update_color_button(self.text_color)
            self.update_opacity_value()
            self.update_intensity_value()

        except Exception as e:
            print(f"加载默认设置失败: {str(e)}")
            # 如果加载失败，使用默认值
            self.apply_default_values()

    # 应用默认值
    def apply_default_values(self):
        """确保所有默认值正确应用到UI元素"""
        # 确保前缀和后缀文本显示正确的默认值
        self.prefix_text.set("IDC圈：仅限")  # 设置默认前缀
        self.suffix_text.set("内部使用，转发侵权")  # 设置默认后缀

        # 确保透明度滑块正确设置
        if not self.opacity_scale.get():
            self.opacity_scale.set(24)
            self.opacity_value.config(text="24%")

        # 确保水印角度选项正确选择
        if not self.watermark_angle.get():
            self.watermark_angle.set(45)

        # 确保字体大小选项正确选择
        if not self.font_size.get():
            self.font_size.set(48)

        # 确保效果类型选项正确选择
        if not self.effect_type.get():
            self.effect_type.set("outline")

        # 确保效果强度正确设置
        if not self.effect_intensity.get():
            self.effect_intensity.set(70)
            self.intensity_value.config(text="70%")

        # 确保轮廓宽度选项正确选择
        if not self.outline_width.get():
            self.outline_width.set(2)

        # 确保阴影偏移选项正确选择
        if not self.shadow_offset.get():
            self.shadow_offset.set(3)

        # 确保图案密度选项正确选择
        if not self.pattern_density.get():
            self.pattern_density.set(5)

        # 确保水印密度选项正确选择
        if not self.watermark_density.get():
            self.watermark_density.set(7)

        # 确保水印位置选项正确选择
        if not self.watermark_position.get():
            self.watermark_position.set("tile")

        # 确保转换质量选项正确选择
        if not self.conversion_quality.get():
            self.conversion_quality.set(200)

        # 确保图片化处理选项正确选择
        if not self.enable_rasterize.get():
            self.enable_rasterize.set(True)

        # 确保压缩级别选项正确选择
        if not self.compression_level.get():
            self.compression_level.set(2)

        # 确保文件命名规则正确设置
        if not self.filename_pattern.get():
            self.filename_pattern.set("文件名{company}")

    # 获取与背景颜色对比度较高的文本颜色
    def get_contrasting_text_color(self, hex_color):
        # 将十六进制颜色转换为RGB分量
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)

        # 计算颜色亮度 (基于ITU-R BT.709标准)
        luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255

        # 根据亮度返回黑色或白色文本
        return "#000000" if luminance > 0.5 else "#FFFFFF"

    # 更新颜色按钮的外观
    def update_color_button(self, color):
        # 更新颜色指示器
        self.color_indicator.config(bg=color)

        # 更新选中状态
        for btn, btn_color in self.color_buttons:
            if btn_color == color:
                btn.config(relief=tk.SUNKEN)
            else:
                btn.config(relief=tk.RAISED)

    def update_intensity_value(self, event=None):
        value = int(self.effect_intensity.get())
        self.intensity_value.config(text=f"{value}%")

    def add_text_watermark_to_image(self, image, text, opacity=0.5, angle=30, font_size=36,
                                    font_family="宋体", color="#FF0000", density=1, position="center",
                                    effect_type="outline", outline_width=2, shadow_offset=3,
                                    effect_intensity=70, pattern_density=5):
        # 确保所有参数为正确类型
        float_opacity = float(opacity)
        int_angle = int(angle)
        int_font_size = int(font_size)
        int_density = int(density)
        int_outline_width = int(outline_width)
        int_shadow_offset = int(shadow_offset)
        int_effect_intensity = int(effect_intensity)
        int_pattern_density = int(pattern_density)

        # 转换为RGBA模式以支持透明度
        if image.mode != 'RGBA':
            image = image.convert('RGBA')

        # 创建透明图层
        watermark = Image.new('RGBA', image.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(watermark)

        # 解析颜色（支持十六进制）- 修复颜色问题
        if color.startswith('#'):
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
        else:
            # 为常见颜色名称提供映射
            color_map = {
                "red": (255, 0, 0),
                "green": (0, 255, 0),
                "blue": (0, 0, 255),
                "black": (0, 0, 0),
                "white": (255, 255, 255),
                "gray": (128, 128, 128)
            }
            r, g, b = color_map.get(color.lower(), (255, 0, 0))  # 默认红色

        # 设置水印文字透明度（0-255）
        alpha = int(float_opacity * 255)

        # 尝试加载字体
        font = None
        try:
            # 尝试从系统字体映射获取字体路径
            font_path = None
            if font_family in self.system_fonts:
                font_path = self.system_fonts[font_family]

            # 根据字体粗细调整字体路径
            if font_path and os.path.exists(font_path):
                font = ImageFont.truetype(font_path, int_font_size)
                self.log(f"使用字体: {font_family} ({font_path})")
            else:
                # 尝试常见中文字体路径
                if platform.system() == 'Darwin':  # macOS
                    possible_fonts = [
                        '/System/Library/Fonts/PingFang.ttc',
                        '/Library/Fonts/Arial Unicode.ttf',
                        '/System/Library/Fonts/STHeiti Medium.ttc',
                        '/System/Library/Fonts/STHeiti Bold.ttc',
                        '/System/Library/Fonts/Hiragino Sans GB.ttc'
                    ]
                    for path in possible_fonts:
                        if os.path.exists(path):
                            font = ImageFont.truetype(path, int_font_size)
                            self.log(f"使用备用字体: {path}")
                            break

                # 如果找不到合适的字体，使用默认字体
                if font is None:
                    font = ImageFont.load_default()
                    self.log("使用系统默认字体")
                    int_font_size = 24  # 调整默认字体大小
        except Exception as e:
            self.log(f"加载字体时出错: {str(e)}")
            font = ImageFont.load_default()
            int_font_size = 24

        # 获取文本尺寸
        if font is not None:
            if hasattr(draw, 'textsize'):
                text_width, text_height = draw.textsize(text, font=font)
            elif hasattr(font, 'getsize'):
                text_width, text_height = font.getsize(text)
            else:
                text_width, text_height = int_font_size * len(text), int_font_size * 1.5
        else:
            text_width, text_height = int_font_size * len(text), int_font_size * 1.5

        # 根据位置和密度设置水印
        if position == "tile":
            # 计算平铺的水印间距
            x_spacing = max(text_width * 1.5, image.width // int_density)  # 确保足够的间距，防止拥挤
            y_spacing = max(text_height * 1.5, image.height // int_density)

            # 计算水印覆盖的行列数
            cols = max(int_density, int(image.width / x_spacing) + 1)  # 确保至少有density个水印，并覆盖整个宽度
            rows = max(int_density, int(image.height / y_spacing) + 1)  # 确保至少有density个水印，并覆盖整个高度

            for i in range(cols):
                for j in range(rows):
                    # 计算水印位置，确保均匀分布
                    x = i * image.width / cols
                    y = j * image.height / rows

                    # 创建旋转后的水印文字
                    txt = Image.new('RGBA', (int(text_width + 60), int(text_height + 60)), (0, 0, 0, 0))
                    d = ImageDraw.Draw(txt)

                    # 根据效果类型应用不同的水印效果
                    if effect_type == "outline":
                        # 轮廓效果
                        self._apply_outline_effect(d, text, font, r, g, b, alpha, int_outline_width,
                                                   int_effect_intensity)
                    elif effect_type == "shadow":
                        # 阴影效果
                        self._apply_shadow_effect(d, text, font, r, g, b, alpha, int_shadow_offset,
                                                  int_effect_intensity)
                    elif effect_type == "emboss":
                        # 浮雕效果
                        self._apply_emboss_effect(d, text, font, r, g, b, alpha, int_effect_intensity)
                    elif effect_type == "texture":
                        # 纹理效果
                        self._apply_texture_effect(d, text, font, r, g, b, alpha, int_pattern_density,
                                                   int_effect_intensity)
                    else:
                        # 默认效果
                        d.text((30, 30), text, font=font, fill=(r, g, b, alpha))

                    rotated = txt.rotate(int_angle, expand=True)

                    # 将旋转后的水印粘贴到透明图层
                    watermark.paste(rotated, (int(x), int(y)), rotated)
        else:  # center
            # 居中放置单个水印
            x = (image.width - text_width) // 2
            y = (image.height - text_height) // 2

            # 创建旋转后的水印文字
            txt = Image.new('RGBA', (int(text_width + 60), int(text_height + 60)), (0, 0, 0, 0))
            d = ImageDraw.Draw(txt)

            # 根据效果类型应用不同的水印效果
            if effect_type == "outline":
                # 轮廓效果
                self._apply_outline_effect(d, text, font, r, g, b, alpha, int_outline_width, int_effect_intensity)
            elif effect_type == "shadow":
                # 阴影效果
                self._apply_shadow_effect(d, text, font, r, g, b, alpha, int_shadow_offset, int_effect_intensity)
            elif effect_type == "emboss":
                # 浮雕效果
                self._apply_emboss_effect(d, text, font, r, g, b, alpha, int_effect_intensity)
            elif effect_type == "texture":
                # 纹理效果
                self._apply_texture_effect(d, text, font, r, g, b, alpha, int_pattern_density, int_effect_intensity)
            else:
                # 默认效果
                d.text((30, 30), text, font=font, fill=(r, g, b, alpha))

            rotated = txt.rotate(int_angle, expand=True)

            # 将旋转后的水印粘贴到透明图层
            watermark.paste(rotated, (int(x - rotated.width // 2), int(y - rotated.height // 2)), rotated)

        # 将水印叠加到原图
        return Image.alpha_composite(image, watermark)

    def _apply_outline_effect(self, draw, text, font, r, g, b, alpha, outline_width, intensity):
        """应用轮廓效果"""
        # 计算轮廓颜色（稍深于主颜色）
        outline_r = max(0, r - 50)
        outline_g = max(0, g - 50)
        outline_b = max(0, b - 50)

        # 计算轮廓透明度（基于强度）
        outline_alpha = int(alpha * (intensity / 100))

        # 绘制多层轮廓
        for i in range(outline_width, 0, -1):
            # 绘制8个方向的轮廓
            draw.text((30 + i, 30 + i), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30 + i, 30 - i), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30 - i, 30 + i), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30 - i, 30 - i), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30 + i, 30), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30 - i, 30), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30, 30 + i), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))
            draw.text((30, 30 - i), text, font=font, fill=(outline_r, outline_g, outline_b, outline_alpha))

        # 绘制主文字
        draw.text((30, 30), text, font=font, fill=(r, g, b, alpha))

    def _apply_shadow_effect(self, draw, text, font, r, g, b, alpha, shadow_offset, intensity):
        """应用阴影效果"""
        # 计算阴影透明度（基于强度）
        shadow_alpha = int(alpha * 0.5 * (intensity / 100))

        # 绘制阴影
        draw.text((30 + shadow_offset, 30 + shadow_offset), text, font=font, fill=(0, 0, 0, shadow_alpha))

        # 绘制主文字
        draw.text((30, 30), text, font=font, fill=(r, g, b, alpha))

    def _apply_emboss_effect(self, draw, text, font, r, g, b, alpha, intensity):
        """应用浮雕效果"""
        # 计算高光和阴影颜色
        highlight_r = min(255, r + 100)
        highlight_g = min(255, g + 100)
        highlight_b = min(255, b + 100)

        shadow_r = max(0, r - 100)
        shadow_g = max(0, g - 100)
        shadow_b = max(0, b - 100)

        # 计算效果透明度（基于强度）
        effect_alpha = int(alpha * (intensity / 100))

        # 绘制高光（左上角）
        draw.text((28, 28), text, font=font, fill=(highlight_r, highlight_g, highlight_b, effect_alpha))

        # 绘制阴影（右下角）
        draw.text((32, 32), text, font=font, fill=(shadow_r, shadow_g, shadow_b, effect_alpha))

        # 绘制主文字（中间）
        draw.text((30, 30), text, font=font, fill=(r, g, b, alpha))

    def _apply_texture_effect(self, draw, text, font, r, g, b, alpha, pattern_density, intensity):
        """应用纹理效果"""
        # 创建临时图像用于纹理
        temp_img = Image.new('RGBA', (200, 200), (0, 0, 0, 0))
        temp_draw = ImageDraw.Draw(temp_img)

        # 绘制主文字
        draw.text((30, 30), text, font=font, fill=(r, g, b, alpha))

        # 添加纹理点
        point_alpha = int(alpha * 0.7 * (intensity / 100))
        point_size = max(1, pattern_density // 3)

        # 在文字周围随机添加点
        import random
        for _ in range(pattern_density * 10):
            x = random.randint(20, 180)
            y = random.randint(20, 180)
            temp_draw.ellipse([x, y, x + point_size, y + point_size], fill=(r, g, b, point_alpha))

        # 将纹理图像应用到主文字
        texture = temp_img.rotate(0, expand=True)
        draw.bitmap((30, 30), texture, fill=(r, g, b, point_alpha))

    def select_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(directory)

    def open_output_folder(self):
        output_dir = self.output_dir.get()
        if output_dir and os.path.exists(output_dir):
            # 根据操作系统打开文件夹
            if os.name == 'nt':  # Windows
                os.startfile(output_dir)
            elif os.name == 'posix':  # macOS, Linux
                if platform.system() == 'Darwin':  # macOS
                    os.system(f'open "{output_dir}"')
                else:  # Linux
                    os.system(f'xdg-open "{output_dir}"')
        else:
            messagebox.showinfo("提示", "请先选择有效的输出文件夹")

    def update_opacity_value(self, event=None):
        value = int(self.opacity_scale.get())
        self.opacity_value.config(text=f"{value}%")

    def choose_color(self):
        color = colorchooser.askcolor(initialcolor=self.text_color)
        if color[1]:  # 用户选择了颜色而不是取消
            self.text_color = color[1]
            # 使用新方法更新颜色按钮
            self.update_color_button(color[1])

            # 更新预览（如果有）
            if hasattr(self, 'preview_image') and self.preview_image:
                self.preview_watermark()

    def set_color(self, color):
        self.text_color = color
        # 使用新方法更新颜色按钮
        self.update_color_button(color)

        # 更新预览（如果有）
        if hasattr(self, 'preview_image') and self.preview_image:
            self.preview_watermark()

    def batch_process(self):
        # 创建开始按钮引用
        if not hasattr(self, 'start_button'):
            # 找到批量处理按钮并保存引用
            for widget in self.batch_tab.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Button) and "批量处理" in str(child.cget('text')):
                            self.start_button = child
                            break
        
        if not self.pdf_path:
            messagebox.showinfo("提示", "请先选择PDF文件")
            return

        if not self.company_names:
            messagebox.showinfo("提示", "请先导入公司名称")
            return

        # 如果启用邮件发送，检查邮箱映射
        if self.enable_email.get():
            if not hasattr(self, 'company_email_map') or not self.company_email_map:
                messagebox.showinfo("提示", "请先导入公司邮箱地址")
                return

            companies_with_email = [k for k, v in self.company_email_map.items() if v]
            companies_without_email = [k for k in self.company_names if
                                       k not in self.company_email_map or not self.company_email_map[k]]

            if companies_without_email:
                result = messagebox.askyesno("警告",
                                             f"发现 {len(companies_without_email)} 个公司没有邮箱地址：\n{', '.join(companies_without_email[:5])}{'...' if len(companies_without_email) > 5 else ''}\n\n是否继续处理？（将跳过没有邮箱的公司的邮件发送）")
                if not result:
                    return

        output_dir = self.output_dir.get()
        if not output_dir:
            messagebox.showinfo("提示", "请选择输出文件夹")
            return

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出文件夹: {str(e)}")
                return

        # 开始处理前清空进度条和日志
        self.progress['value'] = 0
        self.status_bar.config(text="正在处理...")
        
        # 禁用开始按钮，防止重复点击
        if hasattr(self, 'start_button'):
            self.start_button.config(state='disabled')

        # 在新线程中处理，以避免界面冻结
        threading.Thread(target=self.process_watermarks, daemon=True).start()

    def process_watermarks(self):
        try:
            total_companies = len(self.company_names)
            self.log(f"开始处理 {total_companies} 个公司的水印PDF...")

            # 获取处理参数
            opacity = int(self.opacity_scale.get()) / 100  # 转换为0-1范围的浮点数
            angle = int(self.watermark_angle.get())
            font_size = int(self.font_size.get())
            font_family = self.font_family.get()
            color = self.text_color
            density = int(self.watermark_density.get())
            position = self.watermark_position.get()
            quality = int(self.conversion_quality.get())
            enable_rasterize = self.enable_rasterize.get()
            compression_level = int(self.compression_level.get())
            effect_type = self.effect_type.get()
            outline_width = int(self.outline_width.get())
            shadow_offset = int(self.shadow_offset.get())
            effect_intensity = int(self.effect_intensity.get())
            pattern_density = int(self.pattern_density.get())

            # 获取输出目录和文件命名规则
            output_dir = self.output_dir.get()
            filename_pattern = self.filename_pattern.get()

            # 统计邮件发送结果
            email_sent_count = 0
            email_failed_count = 0

            for i, company_name in enumerate(self.company_names):
                # 更新进度
                progress_value = int((i / total_companies) * 100)
                self.master.after(0, lambda v=progress_value: self.progress.config(value=v))
                self.master.after(0, lambda c=company_name, i=i + 1, t=total_companies:
                self.status_bar.config(text=f"处理中: {c} ({i}/{t})"))

                # 生成水印文本
                watermark_text = f"{self.prefix_text.get()}{company_name}{self.suffix_text.get()}"

                # 处理PDF
                try:
                    # 创建输出文件名
                    # 获取上传文件的名称（不含扩展名）
                    base_filename = os.path.splitext(os.path.basename(self.pdf_path))[0]
                    # 替换命名规则中的占位符
                    output_filename = filename_pattern.replace("文件名", base_filename).replace("{company}",
                                                                                                company_name)
                    # 处理文件名中的无效字符
                    output_filename = re.sub(r'[\\/*?:"<>|]', "_", output_filename)
                    output_path = os.path.join(output_dir, f"{output_filename}.pdf")

                    # 添加水印
                    self.apply_watermark_to_pdf(
                        self.pdf_path,
                        output_path,
                        watermark_text,
                        opacity=opacity,
                        angle=angle,
                        font_size=font_size,
                        font_family=font_family,
                        color=color,
                        density=density,
                        position=position,
                        quality=quality,
                        rasterize=enable_rasterize,
                        compression_level=compression_level,
                        effect_type=effect_type,
                        outline_width=outline_width,
                        shadow_offset=shadow_offset,
                        effect_intensity=effect_intensity,
                        pattern_density=pattern_density
                    )

                    self.log(f"已完成: {company_name} -> {output_filename}.pdf")

                    # 如果启用邮件发送，发送邮件到该公司的所有邮箱
                    if self.enable_email.get() and hasattr(self, 'company_email_map'):
                        company_emails = self.company_email_map.get(company_name, [])
                        if company_emails:
                            for email in company_emails:
                                try:
                                    self.send_email(company_name, email, output_path, output_filename)
                                    self.log(f"邮件已发送: {company_name} -> {email}")
                                    email_sent_count += 1
                                except Exception as e:
                                    self.log(f"发送邮件给 {company_name} ({email}) 失败: {str(e)}")
                                    email_failed_count += 1
                        else:
                            self.log(f"跳过邮件发送: {company_name} (无邮箱地址)")

                except Exception as e:
                    self.log(f"处理 {company_name} 时出错: {str(e)}")

            # 处理完成
            self.master.after(0, lambda: self.progress.config(value=100))
            self.master.after(0, lambda: self.status_bar.config(text="处理完成"))
            # 重新启用开始按钮
            self.master.after(0, lambda: self.start_button.config(state='normal') if hasattr(self, 'start_button') else None)

            # 生成完成报告
            completion_msg = f"已完成所有 {total_companies} 个PDF的水印添加"
            if self.enable_email.get():
                completion_msg += f"\n邮件发送统计："
                completion_msg += f"\n✓ 成功发送: {email_sent_count} 封"
                if email_failed_count > 0:
                    completion_msg += f"\n✗ 发送失败: {email_failed_count} 封"
                completion_msg += f"\n总计: {email_sent_count + email_failed_count} 封"

                # 显示每个公司的邮件发送详情
                if hasattr(self, 'company_email_map'):
                    companies_with_multiple_emails = [(k, len(v)) for k, v in self.company_email_map.items() if
                                                      len(v) > 1]
                    if companies_with_multiple_emails:
                        completion_msg += f"\n\n多邮箱公司："
                        for company, count in companies_with_multiple_emails[:3]:  # 只显示前3个
                            completion_msg += f"\n• {company}: {count} 个邮箱"
                        if len(companies_with_multiple_emails) > 3:
                            completion_msg += f"\n... 还有 {len(companies_with_multiple_emails) - 3} 个公司"

            self.master.after(0, lambda: messagebox.showinfo("完成", completion_msg))

        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror("错误", f"批量处理失败: {str(e)}"))
            self.master.after(0, lambda: self.status_bar.config(text="处理失败"))
            # 重新启用开始按钮
            self.master.after(0, lambda: self.start_button.config(state='normal') if hasattr(self, 'start_button') else None)
            self.log(f"批量处理失败: {str(e)}")

    def apply_watermark_to_pdf(self, input_path, output_path, text, opacity=0.5, angle=30,
                               font_size=36, font_family="宋体", color="#FF0000",
                               density=1, position="center", quality=100, rasterize=True,
                               compression_level=0, effect_type="outline", outline_width=2,
                               shadow_offset=3, effect_intensity=70, pattern_density=5):
        if rasterize:
            # 将PDF转换为图像，添加水印，然后转回PDF
            self.rasterize_pdf_with_watermark(
                input_path, output_path, text, opacity, angle, font_size,
                font_family, color, density, position, quality, compression_level,
                effect_type, outline_width, shadow_offset, effect_intensity, pattern_density
            )
        else:
            # 直接添加水印到PDF（不图片化）
            self.add_watermark_to_pdf(
                input_path, output_path, text, opacity, angle, font_size,
                font_family, color, density, position, compression_level,
                effect_type, outline_width, shadow_offset, effect_intensity, pattern_density
            )

    def rasterize_pdf_with_watermark(self, input_path, output_path, text, opacity, angle,
                                     font_size, font_family, color, density, position, quality,
                                     compression_level=0, effect_type="outline", outline_width=2,
                                     shadow_offset=3, effect_intensity=70, pattern_density=5):
        # 创建临时目录
        temp_img_dir = os.path.join(self.temp_dir, "temp_images")
        if not os.path.exists(temp_img_dir):
            os.makedirs(temp_img_dir, exist_ok=True)

        try:
            # 确保关键参数为正确类型
            int_quality = int(quality)  # 保证是整数
            float_opacity = float(opacity)  # 保证是浮点数
            int_angle = int(angle)  # 保证是整数
            int_font_size = int(font_size)  # 保证是整数
            int_density = int(density)  # 保证是整数
            int_compression = int(compression_level)  # 保证是整数
            int_outline_width = int(outline_width)  # 保证是整数
            int_shadow_offset = int(shadow_offset)  # 保证是整数
            int_effect_intensity = int(effect_intensity)  # 保证是整数
            int_pattern_density = int(pattern_density)  # 保证是整数

            # 将PDF转换为图像
            images = convert_from_path(input_path, dpi=int_quality)

            # 添加水印到每个图像
            watermarked_images = []

            # 为每个图像单独创建临时PDF文件
            pdf_files = []

            for i, img in enumerate(images):
                # 添加水印
                watermarked = self.add_text_watermark_to_image(
                    img, text, float_opacity, int_angle, int_font_size, font_family,
                    color, int_density, position, effect_type, int_outline_width,
                    int_shadow_offset, int_effect_intensity, int_pattern_density
                )

                # 根据压缩级别设置JPEG质量
                jpg_quality = self.get_jpg_quality_from_compression_level(int_compression)

                # 将水印图像保存为临时JPG文件
                temp_jpg = os.path.join(temp_img_dir, f"page_{i}.jpg")
                watermarked = watermarked.convert('RGB')  # 转换为RGB模式
                watermarked.save(temp_jpg, "JPEG", quality=jpg_quality)

                # 从JPG创建单页PDF
                temp_pdf = os.path.join(temp_img_dir, f"page_{i}.pdf")
                c = canvas.Canvas(temp_pdf, pagesize=(watermarked.width, watermarked.height))
                c.drawImage(temp_jpg, 0, 0, watermarked.width, watermarked.height)
                c.save()

                pdf_files.append(temp_pdf)

            # 合并所有PDF页面
            merger = PdfMerger()
            for pdf_file in pdf_files:
                merger.append(pdf_file)

            # 保存最终PDF
            merger.write(output_path)
            merger.close()

        except Exception as e:
            self.log(f"图片化PDF处理出错: {str(e)}")
            raise e
        finally:
            # 清理临时文件
            try:
                shutil.rmtree(temp_img_dir)
            except Exception as e:
                self.log(f"清理临时文件出错: {str(e)}")

    def add_watermark_to_pdf(self, input_path, output_path, text, opacity, angle,
                             font_size, font_family, color, density, position, compression_level=0,
                             effect_type="outline", outline_width=2, shadow_offset=3,
                             effect_intensity=70, pattern_density=5):
        try:
            # 确保参数类型正确
            float_opacity = float(opacity)  # 保证是浮点数
            int_angle = int(angle)  # 保证是整数
            int_font_size = int(font_size)  # 保证是整数
            int_density = int(density)  # 保证是整数
            int_compression = int(compression_level)  # 保证是整数
            int_outline_width = int(outline_width)  # 保证是整数
            int_shadow_offset = int(shadow_offset)  # 保证是整数
            int_effect_intensity = int(effect_intensity)  # 保证是整数
            int_pattern_density = int(pattern_density)  # 保证是整数

            # 读取输入PDF
            input_pdf = PdfReader(open(input_path, 'rb'))
            output_pdf = PdfWriter()

            # 查找报告中可用的中文字体
            reportlab_font = "Helvetica-Bold"  # 默认使用粗体字体
            if font_family in self.system_fonts:
                font_path = self.system_fonts[font_family]
                self.log(f"使用字体: {font_family} ({font_path})")
                # 对于ReportLab，需要使用字体的基本名称，而不是完整路径
                if font_family == "宋体":
                    reportlab_font = "SimSun-Bold"
                elif font_family == "黑体":
                    reportlab_font = "SimHei-Bold"
                elif font_family == "微软雅黑":
                    reportlab_font = "MicrosoftYaHei-Bold"
                elif font_family == "微软雅黑粗体":
                    reportlab_font = "MicrosoftYaHei-Bold"
                elif font_family == "Arial Black":
                    reportlab_font = "Arial-Black"
                elif font_family == "Impact":
                    reportlab_font = "Impact"
                else:
                    reportlab_font = font_family + "-Bold"

            # 解析颜色 - 修复颜色问题
            if color.startswith('#'):
                r = int(color[1:3], 16) / 255.0
                g = int(color[3:5], 16) / 255.0
                b = int(color[5:7], 16) / 255.0
                # 添加日志，确认颜色值
                self.log(f"水印颜色: RGB({r:.2f}, {g:.2f}, {b:.2f}) 来自 {color}")
            else:
                # 为常见颜色名称提供映射
                color_map = {
                    "red": (1, 0, 0),
                    "green": (0, 1, 0),
                    "blue": (0, 0, 1),
                    "black": (0, 0, 0),
                    "white": (1, 1, 1),
                    "gray": (0.5, 0.5, 0.5)
                }
                r, g, b = color_map.get(color.lower(), (1, 0, 0))  # 默认红色
                self.log(f"水印颜色: 使用命名颜色 {color} -> RGB({r}, {g}, {b})")

            for page_num in range(len(input_pdf.pages)):
                page = input_pdf.pages[page_num]

                # 创建水印
                packet = io.BytesIO()

                # 获取页面尺寸并转换为float类型
                page_width = float(page.mediabox.width)
                page_height = float(page.mediabox.height)

                c = canvas.Canvas(packet, pagesize=(page_width, page_height))

                # 设置透明度和颜色 - 确保颜色正确设置
                c.setFillColorRGB(r, g, b, alpha=float_opacity)

                # 设置字体
                try:
                    c.setFont(reportlab_font, int_font_size)
                except:
                    c.setFont("Helvetica-Bold", int_font_size)
                    self.log(f"ReportLab无法加载字体 {reportlab_font}，使用默认Helvetica-Bold")

                # 保存当前图形状态
                c.saveState()

                if position == "tile":
                    # 测量文本尺寸（在ReportLab中）
                    text_width = c.stringWidth(text, reportlab_font, int_font_size)
                    text_height = int_font_size * 1.2  # 估计文本高度

                    # 计算平铺的水印间距，确保足够的覆盖
                    x_spacing = max(text_width * 1.5, page_width / int_density)
                    y_spacing = max(text_height * 1.5, page_height / int_density)

                    # 计算水印覆盖的行列数
                    cols = max(int_density, int(page_width / x_spacing) + 1)
                    rows = max(int_density, int(page_height / y_spacing) + 1)

                    for i in range(cols):
                        for j in range(rows):
                            # 计算水印位置，确保均匀分布
                            x = i * page_width / cols
                            y = j * page_height / rows

                            # 绘制旋转的文字
                            c.saveState()
                            c.translate(float(x), float(y))
                            c.rotate(int_angle)

                            # 根据效果类型应用不同的水印效果
                            if effect_type == "outline":
                                # 轮廓效果
                                c.setLineWidth(int_outline_width)
                                c.setStrokeColorRGB(r * 0.7, g * 0.7, b * 0.7, alpha=float_opacity * 0.7)
                                c.setFillColorRGB(r, g, b, alpha=float_opacity)
                                c.drawString(0, 0, text)
                            elif effect_type == "shadow":
                                # 阴影效果
                                c.setFillColorRGB(0, 0, 0, alpha=float_opacity * 0.5)
                                c.drawString(int_shadow_offset, -int_shadow_offset, text)
                                c.setFillColorRGB(r, g, b, alpha=float_opacity)
                                c.drawString(0, 0, text)
                            elif effect_type == "emboss":
                                # 浮雕效果
                                c.setFillColorRGB(min(1, r + 0.4), min(1, g + 0.4), min(1, b + 0.4),
                                                  alpha=float_opacity * 0.7)
                                c.drawString(-1, 1, text)
                                c.setFillColorRGB(max(0, r - 0.4), max(0, g - 0.4), max(0, b - 0.4),
                                                  alpha=float_opacity * 0.7)
                                c.drawString(1, -1, text)
                                c.setFillColorRGB(r, g, b, alpha=float_opacity)
                                c.drawString(0, 0, text)
                            elif effect_type == "texture":
                                # 纹理效果
                                c.setFillColorRGB(r, g, b, alpha=float_opacity)
                                c.drawString(0, 0, text)
                                # 添加纹理点
                                c.setFillColorRGB(r, g, b, alpha=float_opacity * 0.7)
                                for _ in range(int_pattern_density):
                                    import random
                                    tx = random.randint(-20, 20)
                                    ty = random.randint(-20, 20)
                                    c.circle(tx, ty, 1, fill=1, stroke=0)
                            else:
                                # 默认效果
                                c.drawString(0, 0, text)

                            c.restoreState()
                else:  # center
                    # 居中放置单个水印
                    c.saveState()
                    c.translate(page_width / 2, page_height / 2)
                    c.rotate(int_angle)

                    # 根据效果类型应用不同的水印效果
                    if effect_type == "outline":
                        # 轮廓效果
                        c.setLineWidth(int_outline_width)
                        c.setStrokeColorRGB(r * 0.7, g * 0.7, b * 0.7, alpha=float_opacity * 0.7)
                        c.setFillColorRGB(r, g, b, alpha=float_opacity)
                        c.drawCentredString(0, 0, text)
                    elif effect_type == "shadow":
                        # 阴影效果
                        c.setFillColorRGB(0, 0, 0, alpha=float_opacity * 0.5)
                        c.drawCentredString(int_shadow_offset, -int_shadow_offset, text)
                        c.setFillColorRGB(r, g, b, alpha=float_opacity)
                        c.drawCentredString(0, 0, text)
                    elif effect_type == "emboss":
                        # 浮雕效果
                        c.setFillColorRGB(min(1, r + 0.4), min(1, g + 0.4), min(1, b + 0.4), alpha=float_opacity * 0.7)
                        c.drawCentredString(-1, 1, text)
                        c.setFillColorRGB(max(0, r - 0.4), max(0, g - 0.4), max(0, b - 0.4), alpha=float_opacity * 0.7)
                        c.drawCentredString(1, -1, text)
                        c.setFillColorRGB(r, g, b, alpha=float_opacity)
                        c.drawCentredString(0, 0, text)
                    elif effect_type == "texture":
                        # 纹理效果
                        c.setFillColorRGB(r, g, b, alpha=float_opacity)
                        c.drawCentredString(0, 0, text)
                        # 添加纹理点
                        c.setFillColorRGB(r, g, b, alpha=float_opacity * 0.7)
                        for _ in range(int_pattern_density):
                            import random
                            tx = random.randint(-20, 20)
                            ty = random.randint(-20, 20)
                            c.circle(tx, ty, 1, fill=1, stroke=0)
                    else:
                        # 默认效果
                        c.drawCentredString(0, 0, text)

                    c.restoreState()

                # 恢复图形状态
                c.restoreState()
                c.save()

                # 将水印叠加到原始页面
                packet.seek(0)
                watermark_pdf = PdfReader(packet)
                page.merge_page(watermark_pdf.pages[0])

                # 添加处理后的页面到输出PDF
                output_pdf.add_page(page)

            # 保存输出PDF，应用压缩
            with open(output_path, 'wb') as output_file:
                output_pdf.write(output_file)

        except Exception as e:
            raise Exception(f"添加水印到PDF失败: {str(e)}")

    # 根据压缩级别返回JPEG质量
    def get_jpg_quality_from_compression_level(self, compression_level):
        # 压缩级别：0=不压缩，1=轻度，2=中度，3=高度
        quality_map = {
            0: 95,  # 高质量，几乎不压缩
            1: 85,  # 轻度压缩
            2: 70,  # 中度压缩
            3: 50  # 高度压缩
        }
        return quality_map.get(compression_level, 95)  # 默认高质量

    def log(self, message):
        # 在UI线程中更新日志
        self.master.after(0, lambda: self._append_log(message))

    def _append_log(self, message):
        # 优化日志性能：限制日志长度，避免内存溢出
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        
        # 限制日志长度为1000行，超过则删除最早的行
        line_count = int(self.log_text.index('end-1c').split('.')[0])
        if line_count > 1000:
            self.log_text.delete('1.0', '2.0')
        
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def __del__(self):
        # 清理临时目录
        try:
            if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except:
            pass


def main():
    root = tk.Tk()
    app = PDFWatermarkTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
