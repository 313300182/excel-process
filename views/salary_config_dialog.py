# -*- coding: utf-8 -*-
"""
工资配置对话框
提供工资计算参数的配置界面
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from typing import Dict, Any, Optional

from config.salary_settings import DEFAULT_SALARY_CONFIG, SALARY_CONFIG


class SalaryConfigDialog:
    """工资配置对话框"""
    
    def __init__(self, parent, current_config: Dict[str, Any], existing_template_paths: Dict[str, str] = None):
        self.parent = parent
        self.config = current_config.copy()
        self.result = None
        self.template_paths = {}
        self.existing_template_paths = existing_template_paths or {}
        
        # 创建对话框窗口
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("工资配置设置")
        self.dialog.geometry("600x700")
        self.dialog.resizable(True, True)
        
        # 设置对话框为模态
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self._center_dialog()
        
        self.setup_ui()
        
    def _center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"600x700+{x}+{y}")
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建滚动框架
        canvas = tk.Canvas(self.dialog)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 主框架
        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建各个配置区域
        self._create_template_section(main_frame)
        self._create_base_salary_section(main_frame)
        self._create_floating_salary_section(main_frame)
        self._create_commission_section(main_frame)
        self._create_manual_fee_section(main_frame)
        self._create_other_config_section(main_frame)
        self._create_buttons(main_frame)
        
        # 配置滚动
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
    def _create_template_section(self, parent):
        """创建模板文件配置区域"""
        template_frame = ttk.LabelFrame(parent, text="工资模板文件", padding="10")
        template_frame.pack(fill=tk.X, pady=(0, 10))
        
        job_types = SALARY_CONFIG['job_types']
        self.template_vars = {}
        
        for i, job_type in enumerate(job_types):
            # 职业类型标签
            ttk.Label(template_frame, text=f"{job_type}模板:").grid(
                row=i, column=0, sticky=tk.W, pady=2)
            
            # 文件路径输入框
            self.template_vars[job_type] = tk.StringVar()
            # 如果有已保存的模板路径，加载它
            if hasattr(self, 'existing_template_paths') and job_type in self.existing_template_paths:
                self.template_vars[job_type].set(self.existing_template_paths[job_type])
                
            entry = ttk.Entry(template_frame, textvariable=self.template_vars[job_type], width=40)
            entry.grid(row=i, column=1, padx=(5, 5), pady=2, sticky=tk.EW)
            
            # 浏览按钮
            browse_btn = ttk.Button(template_frame, text="浏览",
                                  command=lambda jt=job_type: self._browse_template(jt))
            browse_btn.grid(row=i, column=2, padx=(0, 5), pady=2)
            
        template_frame.columnconfigure(1, weight=1)
        
    def _create_base_salary_section(self, parent):
        """创建基本底薪配置区域"""
        base_frame = ttk.LabelFrame(parent, text="基本底薪配置", padding="10")
        base_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 默认基本底薪
        ttk.Label(base_frame, text="默认基本底薪:").grid(row=0, column=0, sticky=tk.W)
        self.base_salary_var = tk.StringVar(value=str(self.config.get('base_salary', {}).get('default', 5000)))
        ttk.Entry(base_frame, textvariable=self.base_salary_var, width=20).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(base_frame, text="元").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 特殊人员底薪
        ttk.Label(base_frame, text="特殊人员底薪:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        
        special_frame = ttk.Frame(base_frame)
        special_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        special_frame.columnconfigure(0, weight=1)
        
        # 特殊底薪列表
        self.special_base_tree = ttk.Treeview(special_frame, columns=('name', 'salary'), 
                                            show='headings', height=4)
        self.special_base_tree.heading('name', text='姓名')
        self.special_base_tree.heading('salary', text='底薪')
        self.special_base_tree.column('name', width=150)
        self.special_base_tree.column('salary', width=100)
        self.special_base_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        special_scroll = ttk.Scrollbar(special_frame, orient=tk.VERTICAL, 
                                     command=self.special_base_tree.yview)
        special_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.special_base_tree.configure(yscrollcommand=special_scroll.set)
        
        # 添加/删除按钮
        btn_frame = ttk.Frame(base_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        ttk.Button(btn_frame, text="添加", command=self._add_special_base_salary).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="删除", command=self._remove_special_base_salary).pack(side=tk.LEFT)
        
        # 加载现有数据
        self._load_special_base_salary()
        
    def _create_floating_salary_section(self, parent):
        """创建浮动底薪配置区域"""
        floating_frame = ttk.LabelFrame(parent, text="浮动底薪配置", padding="10")
        floating_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 默认浮动底薪
        ttk.Label(floating_frame, text="默认浮动底薪:").grid(row=0, column=0, sticky=tk.W)
        self.floating_salary_var = tk.StringVar(value=str(self.config.get('floating_salary', {}).get('default', 0)))
        ttk.Entry(floating_frame, textvariable=self.floating_salary_var, width=20).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(floating_frame, text="元").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
    def _create_commission_section(self, parent):
        """创建提成配置区域"""
        commission_frame = ttk.LabelFrame(parent, text="提成比例配置", padding="10")
        commission_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 服务提成比例
        ttk.Label(commission_frame, text="服务提成比例:").grid(row=0, column=0, sticky=tk.W)
        self.service_rate_var = tk.StringVar(value=str(self.config.get('commission_rates', {}).get('service_rate', 1.50)))
        ttk.Entry(commission_frame, textvariable=self.service_rate_var, width=20).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(commission_frame, text="%").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 操作提成比例
        ttk.Label(commission_frame, text="操作提成比例:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.operation_rate_var = tk.StringVar(value=str(self.config.get('commission_rates', {}).get('operation_rate', 0.80)))
        ttk.Entry(commission_frame, textvariable=self.operation_rate_var, width=20).grid(
            row=1, column=1, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        ttk.Label(commission_frame, text="%").grid(row=1, column=2, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        
    def _create_manual_fee_section(self, parent):
        """创建手工费配置区域"""
        manual_frame = ttk.LabelFrame(parent, text="手工费配置", padding="10")
        manual_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 身体部位手工费
        ttk.Label(manual_frame, text="身体部位手工费:").grid(row=0, column=0, sticky=tk.W)
        self.body_rate_var = tk.StringVar(value=str(self.config.get('manual_fees', {}).get('body_rate', 60)))
        ttk.Entry(manual_frame, textvariable=self.body_rate_var, width=20).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(manual_frame, text="元/次").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 面部手工费
        ttk.Label(manual_frame, text="面部手工费:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.face_rate_var = tk.StringVar(value=str(self.config.get('manual_fees', {}).get('face_rate', 80)))
        ttk.Entry(manual_frame, textvariable=self.face_rate_var, width=20).grid(
            row=1, column=1, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        ttk.Label(manual_frame, text="元/次").grid(row=1, column=2, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        
    def _create_other_config_section(self, parent):
        """创建其他配置区域"""
        other_frame = ttk.LabelFrame(parent, text="其他配置", padding="10")
        other_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 培训补贴
        ttk.Label(other_frame, text="培训补贴:").grid(row=0, column=0, sticky=tk.W)
        self.training_allowance_var = tk.StringVar(value=str(self.config.get('other_config', {}).get('training_allowance', 200)))
        ttk.Entry(other_frame, textvariable=self.training_allowance_var, width=20).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(other_frame, text="元").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 社保扣除比例
        ttk.Label(other_frame, text="社保扣除比例:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.social_rate_var = tk.StringVar(value=str(self.config.get('other_config', {}).get('social_security_rate', 8.0)))
        ttk.Entry(other_frame, textvariable=self.social_rate_var, width=20).grid(
            row=1, column=1, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        ttk.Label(other_frame, text="%").grid(row=1, column=2, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        
        # 个人所得税比例
        ttk.Label(other_frame, text="个人所得税比例:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        self.tax_rate_var = tk.StringVar(value=str(self.config.get('other_config', {}).get('personal_tax_rate', 3.0)))
        ttk.Entry(other_frame, textvariable=self.tax_rate_var, width=20).grid(
            row=2, column=1, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        ttk.Label(other_frame, text="%").grid(row=2, column=2, padx=(5, 0), sticky=tk.W, pady=(10, 0))
        
    def _create_buttons(self, parent):
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # 重置按钮
        ttk.Button(button_frame, text="重置为默认", command=self._reset_to_default).pack(side=tk.LEFT)
        
        # 右侧按钮
        right_frame = ttk.Frame(button_frame)
        right_frame.pack(side=tk.RIGHT)
        
        ttk.Button(right_frame, text="取消", command=self._cancel).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(right_frame, text="确定", command=self._confirm).pack(side=tk.LEFT)
        
    def _browse_template(self, job_type: str):
        """浏览模板文件"""
        file_path = filedialog.askopenfilename(
            title=f"选择{job_type}工资模板",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.template_vars[job_type].set(file_path)
            
    def _load_special_base_salary(self):
        """加载特殊底薪数据"""
        special_rates = self.config.get('base_salary', {}).get('special_rates', {})
        
        for name, salary in special_rates.items():
            self.special_base_tree.insert('', 'end', values=(name, salary))
            
    def _add_special_base_salary(self):
        """添加特殊底薪"""
        dialog = SpecialSalaryDialog(self.dialog, "添加特殊底薪")
        if dialog.result:
            name, salary = dialog.result
            # 检查是否已存在
            for item in self.special_base_tree.get_children():
                values = self.special_base_tree.item(item)['values']
                if values[0] == name:
                    messagebox.showwarning("警告", f"员工 {name} 已存在！")
                    return
            
            self.special_base_tree.insert('', 'end', values=(name, salary))
            
    def _remove_special_base_salary(self):
        """删除特殊底薪"""
        selection = self.special_base_tree.selection()
        if selection:
            for item in selection:
                self.special_base_tree.delete(item)
        else:
            messagebox.showwarning("警告", "请选择要删除的项目！")
            
    def _reset_to_default(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置为默认配置吗？这将丢失所有自定义设置。"):
            self.config = DEFAULT_SALARY_CONFIG.copy()
            self._refresh_ui()
            
    def _refresh_ui(self):
        """刷新界面"""
        # 刷新基本配置
        self.base_salary_var.set(str(self.config.get('base_salary', {}).get('default', 5000)))
        self.floating_salary_var.set(str(self.config.get('floating_salary', {}).get('default', 0)))
        self.service_rate_var.set(str(self.config.get('commission_rates', {}).get('service_rate', 1.50)))
        self.operation_rate_var.set(str(self.config.get('commission_rates', {}).get('operation_rate', 0.80)))
        self.body_rate_var.set(str(self.config.get('manual_fees', {}).get('body_rate', 60)))
        self.face_rate_var.set(str(self.config.get('manual_fees', {}).get('face_rate', 80)))
        self.training_allowance_var.set(str(self.config.get('other_config', {}).get('training_allowance', 200)))
        self.social_rate_var.set(str(self.config.get('other_config', {}).get('social_security_rate', 8.0)))
        self.tax_rate_var.set(str(self.config.get('other_config', {}).get('personal_tax_rate', 3.0)))
        
        # 清空并重新加载特殊底薪
        for item in self.special_base_tree.get_children():
            self.special_base_tree.delete(item)
        self._load_special_base_salary()
        
    def _validate_config(self) -> bool:
        """验证配置"""
        try:
            # 验证数值字段
            float(self.base_salary_var.get())
            float(self.floating_salary_var.get())
            float(self.service_rate_var.get())
            float(self.operation_rate_var.get())
            float(self.body_rate_var.get())
            float(self.face_rate_var.get())
            float(self.training_allowance_var.get())
            float(self.social_rate_var.get())
            float(self.tax_rate_var.get())
            
            # 验证特殊底薪
            for item in self.special_base_tree.get_children():
                values = self.special_base_tree.item(item)['values']
                float(values[1])  # 验证薪资是数字
                
            return True
            
        except ValueError as e:
            messagebox.showerror("错误", "请输入有效的数值！")
            return False
            
    def _collect_config(self) -> Dict[str, Any]:
        """收集配置数据"""
        # 收集特殊底薪
        special_base_rates = {}
        for item in self.special_base_tree.get_children():
            values = self.special_base_tree.item(item)['values']
            special_base_rates[values[0]] = float(values[1])
            
        # 收集模板路径
        template_paths = {}
        for job_type, var in self.template_vars.items():
            path = var.get().strip()
            if path:
                template_paths[job_type] = path
                
        config = {
            'base_salary': {
                'default': float(self.base_salary_var.get()),
                'special_rates': special_base_rates
            },
            'floating_salary': {
                'default': float(self.floating_salary_var.get()),
                'special_rates': {}
            },
            'commission_rates': {
                'service_rate': float(self.service_rate_var.get()),
                'operation_rate': float(self.operation_rate_var.get())
            },
            'manual_fees': {
                'body_rate': float(self.body_rate_var.get()),
                'face_rate': float(self.face_rate_var.get())
            },
            'other_config': {
                'training_allowance': float(self.training_allowance_var.get()),
                'social_security_rate': float(self.social_rate_var.get()),
                'personal_tax_rate': float(self.tax_rate_var.get())
            }
        }
        
        return config, template_paths
        
    def _confirm(self):
        """确认配置"""
        if self._validate_config():
            config, template_paths = self._collect_config()
            self.result = (config, template_paths)
            self.dialog.destroy()
            
    def _cancel(self):
        """取消配置"""
        self.result = None
        self.dialog.destroy()
        

class SpecialSalaryDialog:
    """特殊底薪输入对话框"""
    
    def __init__(self, parent, title: str):
        self.parent = parent
        self.result = None
        
        # 创建对话框
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("300x150")
        self.dialog.resizable(False, False)
        
        # 设置为模态
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._center_dialog()
        self.setup_ui()
        
    def _center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (150 // 2)
        self.dialog.geometry(f"300x150+{x}+{y}")
        
    def setup_ui(self):
        """设置用户界面"""
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 姓名输入
        ttk.Label(main_frame, text="员工姓名:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=self.name_var, width=20)
        name_entry.grid(row=0, column=1, padx=(10, 0), pady=(0, 10), sticky=tk.EW)
        name_entry.focus()
        
        # 底薪输入
        ttk.Label(main_frame, text="底薪金额:").grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
        self.salary_var = tk.StringVar()
        salary_entry = ttk.Entry(main_frame, textvariable=self.salary_var, width=20)
        salary_entry.grid(row=1, column=1, padx=(10, 0), pady=(0, 10), sticky=tk.EW)
        
        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(20, 0))
        
        ttk.Button(button_frame, text="取消", command=self._cancel).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="确定", command=self._confirm).pack(side=tk.LEFT)
        
        main_frame.columnconfigure(1, weight=1)
        
        # 绑定回车键
        self.dialog.bind('<Return>', lambda e: self._confirm())
        
    def _validate_input(self) -> bool:
        """验证输入"""
        name = self.name_var.get().strip()
        salary = self.salary_var.get().strip()
        
        if not name:
            messagebox.showerror("错误", "请输入员工姓名！")
            return False
            
        if not salary:
            messagebox.showerror("错误", "请输入底薪金额！")
            return False
            
        try:
            float(salary)
        except ValueError:
            messagebox.showerror("错误", "请输入有效的底薪金额！")
            return False
            
        return True
        
    def _confirm(self):
        """确认输入"""
        if self._validate_input():
            name = self.name_var.get().strip()
            salary = float(self.salary_var.get())
            self.result = (name, salary)
            self.dialog.destroy()
            
    def _cancel(self):
        """取消输入"""
        self.result = None
        self.dialog.destroy() 