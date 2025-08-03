# -*- coding: utf-8 -*-
"""
主窗口UI
提供Excel批量处理的图形用户界面
"""

import os
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import logging
from typing import Optional

from controllers.processor import ProcessorController
from controllers.teacher_processor import TeacherProcessorController
from controllers.salary_processor import SalaryProcessorController
from views.salary_config_dialog import SalaryConfigDialog
from config.settings import UI_CONFIG, APP_NAME, APP_VERSION, TEMPLATE_FILENAME


class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.processor = ProcessorController()
        self.teacher_processor = TeacherProcessorController()
        self.salary_processor = SalaryProcessorController()
        self.source_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.template_file = tk.StringVar()  # 添加模板文件路径变量
        self.progress_var = tk.DoubleVar()
        self.progress_text = tk.StringVar(value="就绪")
        
        # 处理模式选择
        self.processing_mode = tk.StringVar(value="normal")  # normal, teacher, 或 salary

        # 设置日志器
        self.logger = logging.getLogger(__name__)

        self.setup_ui()
        self.setup_logging()
        
        # 初始化时设置正确的界面显示状态
        self.on_mode_changed()
        
        # 检查并显示已保存的工资配置状态
        self.check_saved_salary_config()
        
    def setup_ui(self):
        """设置用户界面"""
        # 窗口配置
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        
        # 居中显示窗口
        self._center_window()
        
        self.root.resizable(True, True)
        
        # 设置字体
        default_font = (UI_CONFIG['font_family'], UI_CONFIG['font_size'])
        self.root.option_add('*Font', default_font)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)  # 调整权重行号，因为增加了手工费操作表一行
        
        # 源目录选择
        ttk.Label(main_frame, text="源文件夹:", font=('微软雅黑', 10, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        source_frame.columnconfigure(0, weight=1)
        
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_dir, font=default_font)
        self.source_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(source_frame, text="浏览", command=self.browse_source_dir).grid(row=0, column=1)
        
        # 模板文件选择 (仅在国韩报税模式下显示)
        self.template_label = ttk.Label(main_frame, text="模板文件:", font=('微软雅黑', 10, 'bold'))
        self.template_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        template_frame = ttk.Frame(main_frame)
        template_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        template_frame.columnconfigure(0, weight=1)
        self.template_frame = template_frame  # 保存引用以便后续显示/隐藏
        
        self.template_entry = ttk.Entry(template_frame, textvariable=self.template_file, font=default_font)
        self.template_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(template_frame, text="浏览", command=self.browse_template_file).grid(row=0, column=1)
        
        # 手工费操作表选择 (仅在工资模式下显示)
        self.operation_table_label = ttk.Label(main_frame, text="手工费操作表:", font=('微软雅黑', 10, 'bold'))
        self.operation_table_label.grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        operation_table_frame = ttk.Frame(main_frame)
        operation_table_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        operation_table_frame.columnconfigure(0, weight=1)
        self.operation_table_frame = operation_table_frame
        
        self.operation_table_file = tk.StringVar()
        self.operation_table_entry = ttk.Entry(operation_table_frame, textvariable=self.operation_table_file, font=default_font)
        self.operation_table_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(operation_table_frame, text="浏览", command=self.browse_operation_table).grid(row=0, column=1)
        
        # 输出目录选择
        self.output_label = ttk.Label(main_frame, text="输出文件夹:", font=('微软雅黑', 10, 'bold'))
        self.output_label.grid(row=3, column=0, sticky=tk.W, pady=(0, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_dir, font=default_font)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(output_frame, text="浏览", command=self.browse_output_dir).grid(row=0, column=1)
        
        # 处理模式选择
        mode_frame = ttk.LabelFrame(main_frame, text="处理模式", padding="5")
        mode_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))

        ttk.Radiobutton(mode_frame, text="国韩报税Excel处理（数据提取到模板 + 汇总文件）",
                       variable=self.processing_mode, value="normal",
                       command=self.on_mode_changed).grid(row=0, column=0, sticky=tk.W, padx=(10, 0))

        ttk.Radiobutton(mode_frame, text="业绩分组处理（拆分多Sheet）",
                       variable=self.processing_mode, value="teacher",
                       command=self.on_mode_changed).grid(row=1, column=0, sticky=tk.W, padx=(10, 0))

        ttk.Radiobutton(mode_frame, text="工资Excel处理（生成汇总表 + 各员工工资条）",
                       variable=self.processing_mode, value="salary",
                       command=self.on_mode_changed).grid(row=2, column=0, sticky=tk.W, padx=(10, 0))

        # 工资配置按钮 (仅在工资模式下显示)
        self.salary_config_button = ttk.Button(mode_frame, text="工资配置", 
                                              command=self.open_salary_config)
        self.salary_config_button.grid(row=2, column=1, padx=(20, 0), pady=2)
        
        # 工资输出模式固定为单文件模式（包含汇总功能）
        self.salary_output_mode = tk.StringVar(value="single_file")

        # 文件信息显示
        info_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="5")
        info_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        info_frame.columnconfigure(0, weight=1)
        
        self.info_label = ttk.Label(info_frame, text="请选择源文件夹", 
                                   foreground="gray", font=default_font)
        self.info_label.grid(row=0, column=0, sticky=tk.W)
        
        # 控制按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(10, 5))
        
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)
        
        self.process_button = ttk.Button(button_frame, text="开始处理", command=self.start_processing)
        self.process_button.grid(row=0, column=0, padx=(0, 3), pady=5, sticky='ew')
        
        self.stop_button = ttk.Button(button_frame, text="停止处理", command=self.stop_processing, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=3, pady=5, sticky='ew')

        self.open_output_button = ttk.Button(button_frame, text="打开输出目录", command=self.open_output_directory)
        self.open_output_button.grid(row=0, column=2, padx=(3, 0), pady=5, sticky='ew')
        
        # 进度条
        progress_frame = ttk.LabelFrame(main_frame, text="处理进度", padding="5")
        progress_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        
        # 进度条
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 进度文本
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_text, 
                                       font=default_font)
        self.progress_label.grid(row=0, column=1, padx=(10, 0))
        
        # 日志输出
        log_frame = ttk.Frame(progress_frame)
        log_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, font=('Consolas', 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 绑定事件
        self.source_dir.trace('w', self.on_source_dir_changed)
        
    def setup_logging(self):
        """设置日志记录"""
        # 创建自定义日志处理器，将日志输出到UI
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                
            def emit(self, record):
                msg = self.format(record)
                def append():
                    self.text_widget.configure(state='normal')
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.configure(state='disabled')
                    self.text_widget.see(tk.END)
                
                # 使用after方法确保在主线程中更新UI
                self.text_widget.after(0, append)
        
        # 获取根日志器并清除已有的处理器，避免重复输出
        logger = logging.getLogger()
        logger.handlers.clear()  # 清除所有已有的处理器
        logger.setLevel(logging.INFO)
        
        # 配置UI日志处理器
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)
        
        # 添加控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(console_handler)
        
    def browse_source_dir(self):
        """浏览源目录"""
        directory = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        if directory:
            self.source_dir.set(directory)
            
    def browse_output_dir(self):
        """浏览输出目录"""
        directory = filedialog.askdirectory(title="选择输出文件夹")
        if directory:
            self.output_dir.set(directory)
    
    def browse_template_file(self):
        """浏览模板文件"""
        file_path = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.template_file.set(file_path)
            
    def browse_operation_table(self):
        """浏览手工费操作表文件"""
        file_path = filedialog.askopenfilename(
            title="选择手工费操作表",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.operation_table_file.set(file_path)
            self.log_message(f"已选择手工费操作表: {file_path}", "INFO")
            
            # 立即尝试读取和验证操作表
            self.log_message("开始读取手工费操作表内容...", "INFO")
            try:
                self.salary_processor.set_operation_table_path(file_path)
                # 检查是否成功读取数据
                if hasattr(self.salary_processor, 'operation_data') and self.salary_processor.operation_data:
                    self.log_message("✅ 手工费操作表读取成功!", "INFO")
                    
                    # 保存操作表路径到用户配置
                    try:
                        current_config = self.salary_processor.get_user_config()
                        current_config['last_operation_table_path'] = file_path
                        if self.salary_processor.save_user_config(current_config):
                            self.log_message("📝 操作表路径已保存到配置", "INFO")
                        else:
                            self.log_message("⚠️  操作表路径保存失败", "WARNING")
                    except Exception as e:
                        self.log_message(f"保存操作表路径配置失败: {str(e)}", "WARNING")
                else:
                    raise Exception("未能读取到有效的操作表数据")
            except Exception as e:
                self.log_message(f"❌ 读取手工费操作表失败: {str(e)}", "ERROR")
                messagebox.showerror("文件错误", f"读取手工费操作表失败:\n{str(e)}")
            
    def open_salary_config(self):
        """打开工资配置对话框"""
        try:
            current_config = self.salary_processor.get_user_config()
            # 获取当前的模板路径
            existing_templates = getattr(self.salary_processor, 'template_paths', {})
            dialog = SalaryConfigDialog(self.root, current_config, existing_templates)
            
            result = dialog.show()
            
            if result:
                # 分离模板路径和配置数据
                template_paths = result.pop('template_paths', {})
                config = result
                
                # 保存用户配置
                if self.salary_processor.save_user_config(config):
                    self.log_message("✅ 工资配置已保存", "INFO")
                else:
                    self.log_message("❌ 工资配置保存失败", "ERROR")
                    
                # 设置模板路径
                if template_paths:
                    self.salary_processor.set_template_paths(template_paths)
                    self.log_message(f"📁 已设置 {len(template_paths)} 个工资模板", "INFO")
                    
                    # 验证模板
                    validation_results = self.salary_processor.validate_templates()
                    invalid_templates = [k for k, v in validation_results.items() if not v]
                    if invalid_templates:
                        messagebox.showwarning("模板验证", 
                                             f"以下模板验证失败: {', '.join(invalid_templates)}")
                    else:
                        self.log_message("✅ 所有工资模板验证通过", "INFO")
                else:
                    messagebox.showwarning("警告", "请至少设置一个工资模板文件")
                    
                # 重新验证源目录
                if self.source_dir.get():
                    self.on_source_dir_changed()
                    
        except Exception as e:
            self.log_message(f"打开工资配置失败: {str(e)}", "ERROR")
            messagebox.showerror("错误", f"打开工资配置失败: {str(e)}")
            
    def check_saved_salary_config(self):
        """检查已保存的工资配置状态"""
        try:
            # 检查是否有保存的模板路径
            if hasattr(self.salary_processor, 'template_paths') and self.salary_processor.template_paths:
                template_count = len(self.salary_processor.template_paths)
                self.log_message(f"✅ 发现已保存的工资配置: {template_count} 个模板", "INFO")
                
                # 验证模板文件是否仍然存在
                missing_templates = []
                valid_templates = []
                for job_type, template_path in self.salary_processor.template_paths.items():
                    if os.path.exists(template_path):
                        valid_templates.append(job_type)
                    else:
                        missing_templates.append(f"{job_type}: {template_path}")
                
                if valid_templates:
                    self.log_message(f"📋 有效模板: {', '.join(valid_templates)}", "INFO")
                    
                if missing_templates:
                    self.log_message("⚠️  部分模板文件缺失:", "WARNING")
                    for missing in missing_templates:
                        self.log_message(f"   - {missing}", "WARNING")
                    self.log_message("建议重新设置工资配置", "WARNING")
                    
            else:
                self.log_message("ℹ️  未找到已保存的工资配置，使用工资模式时请先设置", "INFO")
                
            # 检查是否有保存的手工费操作表路径
            user_config = self.salary_processor.get_user_config()
            if 'last_operation_table_path' in user_config:
                operation_path = user_config['last_operation_table_path']
                if os.path.exists(operation_path):
                    self.operation_table_file.set(operation_path)
                    self.log_message(f"📋 自动加载上次使用的手工费操作表: {os.path.basename(operation_path)}", "INFO")
                    try:
                        self.salary_processor.set_operation_table_path(operation_path)
                    except Exception as e:
                        self.log_message(f"⚠️  加载操作表失败: {str(e)}", "WARNING")
                else:
                    self.log_message(f"⚠️  上次使用的手工费操作表不存在: {operation_path}", "WARNING")
                        
        except Exception as e:
            self.log_message(f"检查配置状态时出错: {str(e)}", "ERROR")

    def on_mode_changed(self):
        """处理模式改变时的回调"""
        mode = self.processing_mode.get()
        
        # 根据模式显示或隐藏界面元素
        if mode == "normal":
            # 国韩报税模式 - 显示模板文件选择
            self.template_label.grid()
            self.template_frame.grid()
            self.salary_config_button.grid_remove()
            self.operation_table_label.grid_remove()
            self.operation_table_frame.grid_remove()

        elif mode == "teacher":
            # 老师分组模式 - 隐藏模板文件选择和工资相关界面
            self.template_label.grid_remove()
            self.template_frame.grid_remove()
            self.salary_config_button.grid_remove()
            self.operation_table_label.grid_remove()
            self.operation_table_frame.grid_remove()

        elif mode == "salary":
            # 工资处理模式 - 隐藏模板文件选择，显示工资相关界面
            self.template_label.grid_remove()
            self.template_frame.grid_remove()
            self.salary_config_button.grid()
            self.operation_table_label.grid()
            self.operation_table_frame.grid()

            
        # 重新验证目录以更新文件信息
        if self.source_dir.get():
            self.on_source_dir_changed()

    def on_source_dir_changed(self, *args):
        """源目录改变时的处理"""
        directory = self.source_dir.get()
        if directory and os.path.exists(directory):
            mode = self.processing_mode.get()
            
            # 根据处理模式选择不同的验证器
            if mode == "teacher":
                result = self.teacher_processor.validate_teacher_source_directory(directory)
            elif mode == "salary":
                info = self.salary_processor.get_processing_info(directory)
                result = {
                    'valid': info['valid_files'] > 0,
                    'message': f"找到 {info['valid_files']} 个有效的工资文件" if info['valid_files'] > 0 else "没有找到有效的工资文件",
                    'files': [emp['file_name'] for emp in info['employee_info']],
                    'file_count': info['valid_files']
                }
                if not info['ready_to_process'] and 'error_message' in info:
                    result['message'] = info['error_message']
                    result['valid'] = False
            else:
                result = self.processor.validate_source_directory(directory)

            if result['valid']:
                self.info_label.config(text=f"✓ {result['message']}", foreground="green")
                if result.get('files'):
                    file_list = ', '.join(result['files'][:5])
                    if result.get('file_count', 0) > 5:
                        file_list += f" ... (共{result['file_count']}个文件)"
                    self.info_label.config(text=f"✓ {result['message']}\n示例文件: {file_list}")
            else:
                self.info_label.config(text=f"✗ {result['message']}", foreground="red")
        else:
            self.info_label.config(text="请选择有效的源文件夹", foreground="gray")
            
    def start_processing(self):
        """开始处理"""
        source = self.source_dir.get()
        output = self.output_dir.get()
        
        # 验证输入
        if not source:
            messagebox.showerror("错误", "请选择源文件夹")
            return
            
        if not output:
            messagebox.showerror("错误", "请选择输出文件夹")
            return
            
        if not os.path.exists(source):
            messagebox.showerror("错误", "源文件夹不存在")
            return
            
        mode = self.processing_mode.get()
        
        # 根据处理模式验证源目录和准备输出目录
        if mode == "teacher":
            # 老师分组模式
            validation = self.teacher_processor.validate_teacher_source_directory(source)
            if not validation['valid']:
                messagebox.showerror("错误", validation['message'])
                return

            # 准备老师分组输出目录
            actual_output_dir = self.teacher_processor.prepare_teacher_output_directory(output)
        elif mode == "salary":
            # 工资处理模式 - 全面验证
            try:
                # 1. 检查模板配置
                if not hasattr(self.salary_processor, 'template_paths') or not self.salary_processor.template_paths:
                    messagebox.showerror("错误", "请先点击'工资配置'设置工资模板文件")
                    return
                    
                # 2. 验证模板文件存在性
                missing_templates = []
                for job_type, template_path in self.salary_processor.template_paths.items():
                    if not os.path.exists(template_path):
                        missing_templates.append(f"{job_type}: {template_path}")
                
                if missing_templates:
                    messagebox.showerror("错误", f"以下模板文件不存在:\n" + "\n".join(missing_templates))
                    return
                
                # 3. 检查操作表
                operation_table_path = self.operation_table_file.get().strip()
                if not operation_table_path:
                    messagebox.showerror("错误", "请选择手工费操作表文件")
                    return
                if not os.path.exists(operation_table_path):
                    messagebox.showerror("错误", "手工费操作表文件不存在")
                    return
                    
                # 4. 设置操作表路径（这里会验证文件格式）
                self.salary_processor.set_operation_table_path(operation_table_path)
                
                # 5. 检查源目录
                excel_files = self.salary_processor.scan_excel_files(source)
                if not excel_files:
                    messagebox.showerror("错误", "源目录中没有找到有效的Excel文件")
                    return
                
                self.log_message(f"找到 {len(excel_files)} 个待处理文件", "INFO")
                
                # 6. 准备输出目录
                actual_output_dir = output
                os.makedirs(actual_output_dir, exist_ok=True)
                
            except Exception as e:
                self.log_message(f"工资处理验证失败: {str(e)}", "ERROR")
                messagebox.showerror("验证错误", f"工资处理验证失败:\n{str(e)}")
                return
        else:
            # 常规模式 (国韩报税模式)
            validation = self.processor.validate_source_directory(source)
            if not validation['valid']:
                messagebox.showerror("错误", validation['message'])
                return

            # 检查用户是否选择了模板文件
            template_path = self.template_file.get()
            if not template_path:
                messagebox.showerror("错误", "请选择模板文件")
                return
                
            if not os.path.exists(template_path):
                messagebox.showerror("错误", "模板文件不存在")
                return

            # 验证模板文件是否有效
            try:
                from openpyxl import load_workbook
                workbook = load_workbook(template_path)
                workbook.close()
            except Exception as e:
                self.log_message(f"模板文件无效: {template_path}, 错误: {e}", "ERROR")
                messagebox.showerror("模板错误", f"模板文件无效: {template_path}")
                return

            # 设置用户选择的模板文件路径到处理器
            self.processor.set_template_path(template_path)

            # 准备输出目录
            actual_output_dir = self.processor.prepare_output_directory(output)
        
        # 更新UI状态
        self.process_button.config(state='disabled')
        self.stop_button.config(state='normal')
        self.progress_var.set(0)
        self.progress_text.set("正在处理...")
        
        # 清空日志
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        
        # 根据处理模式开始处理
        if mode == "teacher":
            # 老师分组处理
            self.teacher_processor.process_teacher_batch(
                source_dir=source,
                output_dir=actual_output_dir,
                progress_callback=self.update_progress,
                complete_callback=self.on_processing_complete
            )
        elif mode == "salary":
            # 工资处理
            self.start_salary_processing(source, actual_output_dir)
        else:
            # 常规处理
            self.processor.process_batch(
                source_dir=source,
                output_dir=actual_output_dir,
                progress_callback=self.update_progress,
                complete_callback=self.on_processing_complete
            )
        
    def start_salary_processing(self, source_dir: str, output_dir: str):
        """开始工资处理"""
        
        try:
            self.log_message("🚀 开始工资处理...", "INFO")
            
            # 验证处理器状态
            if not hasattr(self.salary_processor, 'template_paths') or not self.salary_processor.template_paths:
                raise Exception("请先设置工资模板文件")
            
            self.log_message("✅ 处理器状态验证通过", "INFO")
            
            # 强制更新UI，让用户看到进度
            self.root.update()
            
            # 使用单文件多Sheet模式处理（包含汇总表）
            self.log_message("📋 使用单文件多Sheet模式处理（包含汇总表）...", "INFO")
            
            result = self.salary_processor.process_files_to_single_excel(
                source_dir=source_dir,
                output_dir=output_dir,
                progress_callback=self.update_salary_progress,
                log_callback=self.log_message
            )
            
            self.log_message("🔄 处理器执行完毕，准备完成回调", "INFO")
            
            # 完成处理
            self.on_salary_processing_complete(result)
            
        except Exception as e:
            import traceback
            error_msg = f"工资处理异常: {str(e)}"
            traceback_msg = traceback.format_exc()
            
            # 记录详细错误
            self.logger.error(f"{error_msg}\n{traceback_msg}")
            print(f"[处理异常] {error_msg}")
            print(f"[异常详情] {traceback_msg}")
            
            # 更新UI
            self.log_message(error_msg, "ERROR")
            messagebox.showerror("处理错误", 
                f"工资处理失败:\n\n{str(e)}\n\n请检查控制台了解详细信息。")
            self.reset_ui_state()
        
    def update_salary_progress(self, progress: float):
        """更新工资处理进度"""
        def update():
            try:
                self.progress_var.set(progress)
                self.progress_text.set(f"正在处理工资文件... {progress:.1f}%")
                self.root.update_idletasks()
            except Exception as e:
                print(f"进度更新失败: {e}")
        
        self.root.after_idle(update)
        
    def log_message_safe(self, message: str):
        """线程安全的日志消息"""
        self.root.after(0, lambda: self.log_message(message, "INFO"))
        
    def on_salary_processing_complete(self, result: dict):
        """工资处理完成回调"""
        try:
            self.reset_ui_state()
            
            # 单文件模式的结果处理（汇总表 + 各员工工资条）
            success = result.get('success', False)
            processed_employees = result.get('processed_employees', 0)
            total_employees = result.get('total_employees', 0)
            processed_files = result.get('processed_files', 0)
            output_file = result.get('output_file', '')
            
            self.log_message(f"工资处理完成: 处理员工 {processed_employees}人, 来源文件 {processed_files}个", "INFO")
            
            if result.get('errors'):
                for error in result['errors'][:5]:  # 只显示前5个错误
                    self.log_message(f"错误: {error}", "ERROR")
                    
            # 显示完成消息
            if success and processed_employees > 0:
                self.progress_text.set(f"完成: 汇总 {processed_employees} 人工资单")
                messagebox.showinfo("处理完成",
                    f"工资汇总处理完成！\n\n处理员工: {processed_employees} 人\n来源文件: {processed_files} 个\n\n📊 已生成汇总表和各员工工资条\n\n汇总文件已保存:\n{output_file}")
            else:
                self.progress_text.set("处理失败")
                messagebox.showerror("处理失败", 
                    f"工资汇总处理失败！\n\n详细错误请查看日志。")
                    
        except Exception as e:
            self.log_message(f"处理完成回调出错: {str(e)}", "ERROR")
            
    def reset_ui_state(self):
        """重置UI状态"""
        self.process_button.config(state='normal')
        self.stop_button.config(state='disabled')
        self.progress_var.set(100)
        
    def stop_processing(self):
        """停止处理"""
        self.log_message("用户请求停止处理...", "WARNING")

        # 根据当前处理模式停止对应的处理器
        mode = self.processing_mode.get()
        if mode == "teacher":
            self.teacher_processor.stop_processing()
        elif mode == "salary":
            self.salary_processor.stop_processing()
        else:
            self.processor.stop_processing()

        self.progress_text.set("正在停止...")
        
    def update_progress(self, current: int, total: int, current_file: str):
        """更新进度条和日志"""
        def update():
            try:
                progress = (current / total) * 100 if total > 0 else 0
                self.progress_var.set(progress)
                self.progress_text.set(f"正在处理: {current_file} ({current}/{total})")

                # 强制更新UI
                self.root.update_idletasks()

            except Exception as e:
                print(f"进度更新失败: {e}")  # 使用print避免日志循环
            
        # 使用after_idle确保UI响应
        self.root.after_idle(update)
        
    def on_processing_complete(self, success_files: list, failed_files: list):
        """处理完成回调"""
        def update():
            try:
                self.process_button.config(state='normal')
                self.stop_button.config(state='disabled')
                self.progress_var.set(100)

                # 记录日志
                self.log_message(f"批量处理完成: 成功 {len(success_files)}, 失败 {len(failed_files)}", "INFO")

                # 显示单一的完成提示
                if failed_files:
                    self.progress_text.set(f"完成: 成功 {len(success_files)}, 失败 {len(failed_files)}")
                    messagebox.showwarning("处理完成",
                        f"批量处理完成！\n\n成功: {len(success_files)} 个文件\n失败: {len(failed_files)} 个文件\n\n详细信息请查看日志")
                else:
                    self.progress_text.set(f"全部完成: {len(success_files)} 个文件")
                    messagebox.showinfo("处理完成",
                        f"批量处理完成！\n\n成功处理 {len(success_files)} 个文件\n\n详情请查看日志。")

                # 强制更新UI
                self.root.update_idletasks()
                
            except Exception as e:
                print(f"完成回调失败: {e}")  # 使用print避免日志循环
                
        # 使用after确保在主线程中执行
        self.root.after(100, update)  # 稍微延迟确保所有处理完成
        
    def open_output_directory(self):
        """打开输出目录"""
        output_dir = self.output_dir.get()

        if not output_dir:
            messagebox.showwarning("提示", "请先选择输出文件夹")
            return
            
        # 检查目录是否存在
        if not os.path.exists(output_dir):
            messagebox.showerror("错误", f"输出目录不存在: {output_dir}")
            return
            
        try:
            # 根据操作系统使用不同的命令打开文件夹
            if os.name == 'nt':  # Windows
                os.startfile(output_dir)
            elif os.name == 'posix':  # macOS and Linux
                if os.uname().sysname == 'Darwin':  # macOS
                    subprocess.run(['open', output_dir])
                else:  # Linux
                    subprocess.run(['xdg-open', output_dir])

            self.log_message(f"已打开输出目录: {output_dir}", "INFO")

        except Exception as e:
            self.log_message(f"打开输出目录失败: {str(e)}", "ERROR")
            messagebox.showerror("错误", f"无法打开输出目录:\n{str(e)}")

    def log_message(self, message: str, level: str = "INFO"):
        """记录日志消息"""
        # 只通过logger记录，让TextHandler处理UI显示，避免重复
        self.logger.log(getattr(logging, level.upper(), logging.INFO), message)
        
    def _center_window(self):
        """居中显示窗口"""
        # 从配置文件获取窗口尺寸
        window_size = UI_CONFIG['window_size']  # 格式: "800x600"
        width, height = map(int, window_size.split('x'))
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 计算居中位置
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # 设置窗口位置和大小
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def run(self):
        """运行应用"""
        # 启动主循环
        self.root.mainloop()


def main():
    """主函数"""
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main() 