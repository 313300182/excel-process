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
from config.settings import UI_CONFIG, APP_NAME, APP_VERSION, TEMPLATE_FILENAME


class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.processor = ProcessorController()
        self.teacher_processor = TeacherProcessorController()
        self.source_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.progress_text = tk.StringVar(value="就绪")
        
        # 处理模式选择
        self.processing_mode = tk.StringVar(value="normal")  # normal 或 teacher

        self.setup_ui()
        self.setup_logging()
        
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
        main_frame.rowconfigure(5, weight=1)
        
        # 源目录选择
        ttk.Label(main_frame, text="源文件夹:", font=('微软雅黑', 10, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        source_frame.columnconfigure(0, weight=1)
        
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_dir, font=default_font)
        self.source_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(source_frame, text="浏览", command=self.browse_source_dir).grid(row=0, column=1)
        
        # 输出目录选择
        ttk.Label(main_frame, text="输出文件夹:", font=('微软雅黑', 10, 'bold')).grid(
            row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_dir, font=default_font)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(output_frame, text="浏览", command=self.browse_output_dir).grid(row=0, column=1)
        
        # 处理模式选择
        mode_frame = ttk.LabelFrame(main_frame, text="处理模式", padding="5")
        mode_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))

        ttk.Radiobutton(mode_frame, text="国韩报税Excel处理（数据提取到模板）",
                       variable=self.processing_mode, value="normal",
                       command=self.on_mode_changed).grid(row=0, column=0, sticky=tk.W, padx=(10, 0))

        ttk.Radiobutton(mode_frame, text="业绩分组处理（拆分多Sheet）",
                       variable=self.processing_mode, value="teacher",
                       command=self.on_mode_changed).grid(row=1, column=0, sticky=tk.W, padx=(10, 0))

        # 文件信息显示
        info_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="5")
        info_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        info_frame.columnconfigure(0, weight=1)
        
        self.info_label = ttk.Label(info_frame, text="请选择源文件夹", 
                                   foreground="gray", font=default_font)
        self.info_label.grid(row=0, column=0, sticky=tk.W)
        
        # 控制按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=(10, 5))
        
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
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
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
        
        # 配置日志
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
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
            
    def on_mode_changed(self):
        """处理模式改变时的回调"""
        # 重新验证目录以更新文件信息
        if self.source_dir.get():
            self.on_source_dir_changed()

    def on_source_dir_changed(self, *args):
        """源目录改变时的处理"""
        directory = self.source_dir.get()
        if directory and os.path.exists(directory):
            # 根据处理模式选择不同的验证器
            if self.processing_mode.get() == "teacher":
                result = self.teacher_processor.validate_teacher_source_directory(directory)
            else:
                result = self.processor.validate_source_directory(directory)

            if result['valid']:
                self.info_label.config(text=f"✓ {result['message']}", foreground="green")
                if result['files']:
                    file_list = ', '.join(result['files'][:5])
                    if result['file_count'] > 5:
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
            
        # 根据处理模式验证源目录和准备输出目录
        if self.processing_mode.get() == "teacher":
            # 老师分组模式
            validation = self.teacher_processor.validate_teacher_source_directory(source)
            if not validation['valid']:
                messagebox.showerror("错误", validation['message'])
                return

            # 准备老师分组输出目录
            actual_output_dir = self.teacher_processor.prepare_teacher_output_directory(output)
        else:
            # 常规模式
            validation = self.processor.validate_source_directory(source)
            if not validation['valid']:
                messagebox.showerror("错误", validation['message'])
                return

            # 检查模板文件是否存在
            if not self.processor.writer.validate_template():
                self.log_message(f"模板文件不存在或无效: {self.processor.writer.template_path}", "ERROR")
                messagebox.showerror("模板错误", f"模板文件不存在或无效: {self.processor.writer.template_path}")
                return

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
        if self.processing_mode.get() == "teacher":
            # 老师分组处理
            self.teacher_processor.process_teacher_batch(
                source_dir=source,
                output_dir=actual_output_dir,
                progress_callback=self.update_progress,
                complete_callback=self.on_processing_complete
            )
        else:
            # 常规处理
            self.processor.process_batch(
                source_dir=source,
                output_dir=actual_output_dir,
                progress_callback=self.update_progress,
                complete_callback=self.on_processing_complete
            )
        
    def stop_processing(self):
        """停止处理"""
        self.log_message("用户请求停止处理...", "WARNING")

        # 根据当前处理模式停止对应的处理器
        if self.processing_mode.get() == "teacher":
            self.teacher_processor.stop_processing()
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