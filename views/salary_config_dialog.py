# -*- coding: utf-8 -*-
"""
工资配置对话框
提供工资计算参数的配置界面 - 支持三个职业类型的独立配置
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from typing import Dict, Any, Optional

from config.salary_settings import DEFAULT_SALARY_CONFIG, SALARY_CONFIG, JOB_SPECIFIC_CONFIG


class SalaryConfigDialog:
    """工资配置对话框"""
    
    def __init__(self, parent, current_config: Dict[str, Any], existing_template_paths: Dict[str, str] = None):
        self.parent = parent
        self.config = current_config.copy()
        self.result = None
        self.existing_template_paths = existing_template_paths or {}
        
        # 职业类型列表
        self.job_types = ['服务总监', '服务老师', '操作老师']
        
        # 存储各职业的配置变量
        self.job_configs = {}
        for job_type in self.job_types:
            self.job_configs[job_type] = {
                'template_var': tk.StringVar(),
                'base_salary_var': tk.StringVar(),
                'floating_salary_var': tk.StringVar(),
                'commission_rate_var': tk.StringVar(),
                'body_rate_var': tk.StringVar(),
                'face_rate_var': tk.StringVar(),
                'training_allowance_var': tk.StringVar(),
                'social_rate_var': tk.StringVar(),
                'tax_rate_var': tk.StringVar(),
                'late_deduction_rate_var': tk.StringVar()
            }
        
        # 创建对话框窗口
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("工资配置设置")
        self.dialog.geometry("800x700")
        self.dialog.resizable(True, True)
        
        # 设置对话框为模态
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self._center_dialog()
        
        # 绑定关闭事件
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        self._setup_ui()
        self._load_current_config()
        
    def _center_dialog(self):
        """居中显示对话框"""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"800x700+{x}+{y}")
        
    def _setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="工资配置设置", 
                               font=("", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 创建职业类型选项卡
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 为每个职业类型创建选项卡
        for job_type in self.job_types:
            self._create_job_tab(job_type)
        
        # 创建按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 添加按钮
        ttk.Button(button_frame, text="重置为默认", 
                  command=self._reset_to_default).pack(side=tk.LEFT)
        
        ttk.Button(button_frame, text="取消", 
                  command=self._on_cancel).pack(side=tk.RIGHT, padx=(5, 0))
        
        ttk.Button(button_frame, text="确定", 
                  command=self._on_ok).pack(side=tk.RIGHT)
        
    def _create_job_tab(self, job_type: str):
        """创建职业类型选项卡"""
        # 创建选项卡框架
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=job_type)
        
        # 创建滚动框架
        canvas = tk.Canvas(tab_frame)
        scrollbar = ttk.Scrollbar(tab_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 打包滚动组件
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 创建配置内容
        self._create_template_section(scrollable_frame, job_type)
        self._create_salary_section(scrollable_frame, job_type)
        self._create_commission_section(scrollable_frame, job_type)
        self._create_manual_fee_section(scrollable_frame, job_type)
        self._create_other_config_section(scrollable_frame, job_type)
        
    def _create_template_section(self, parent, job_type: str):
        """创建模板文件配置区域"""
        frame = ttk.LabelFrame(parent, text="模板文件", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(frame, text="工资条模板:").grid(row=0, column=0, sticky=tk.W)
        
        template_var = self.job_configs[job_type]['template_var']
        ttk.Entry(frame, textvariable=template_var, width=50).grid(
            row=0, column=1, padx=(5, 5), sticky=tk.W)
        
        ttk.Button(frame, text="浏览", 
                  command=lambda: self._browse_template(job_type)).grid(
            row=0, column=2, padx=(5, 0))
        
    def _create_salary_section(self, parent, job_type: str):
        """创建底薪配置区域"""
        frame = ttk.LabelFrame(parent, text="底薪配置", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))
        
        # 基本底薪
        ttk.Label(frame, text="基本底薪:").grid(row=0, column=0, sticky=tk.W)
        base_var = self.job_configs[job_type]['base_salary_var']
        ttk.Entry(frame, textvariable=base_var, width=15).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(frame, text="元").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 浮动底薪
        ttk.Label(frame, text="浮动底薪:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        floating_var = self.job_configs[job_type]['floating_salary_var']
        ttk.Entry(frame, textvariable=floating_var, width=15).grid(
            row=1, column=1, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        ttk.Label(frame, text="元").grid(row=1, column=2, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        
    def _create_commission_section(self, parent, job_type: str):
        """创建提成配置区域"""
        frame = ttk.LabelFrame(parent, text="提成配置", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))
        
        # 根据职业类型显示对应的提成类型
        if job_type == '服务总监':
            commission_label = "专家提成比例:"
        elif job_type == '服务老师':
            commission_label = "服务提成比例:"
        else:  # 操作老师
            commission_label = "操作提成比例:"
        
        ttk.Label(frame, text=commission_label).grid(row=0, column=0, sticky=tk.W)
        commission_var = self.job_configs[job_type]['commission_rate_var']
        ttk.Entry(frame, textvariable=commission_var, width=15).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(frame, text="%").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
    def _create_manual_fee_section(self, parent, job_type: str):
        """创建手工费配置区域"""
        frame = ttk.LabelFrame(parent, text="手工费配置", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))
        
        # 身体部位手工费
        ttk.Label(frame, text="身体部位单价:").grid(row=0, column=0, sticky=tk.W)
        body_var = self.job_configs[job_type]['body_rate_var']
        ttk.Entry(frame, textvariable=body_var, width=15).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(frame, text="元/次").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 面部手工费
        ttk.Label(frame, text="面部单价:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        face_var = self.job_configs[job_type]['face_rate_var']
        ttk.Entry(frame, textvariable=face_var, width=15).grid(
            row=1, column=1, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        ttk.Label(frame, text="元/次").grid(row=1, column=2, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        
    def _create_other_config_section(self, parent, job_type: str):
        """创建其他配置区域"""
        frame = ttk.LabelFrame(parent, text="其他配置", padding="10")
        frame.pack(fill=tk.X, pady=(0, 10))
        
        # 培训补贴
        ttk.Label(frame, text="培训补贴:").grid(row=0, column=0, sticky=tk.W)
        training_var = self.job_configs[job_type]['training_allowance_var']
        ttk.Entry(frame, textvariable=training_var, width=15).grid(
            row=0, column=1, padx=(5, 0), sticky=tk.W)
        ttk.Label(frame, text="元/天").grid(row=0, column=2, padx=(5, 0), sticky=tk.W)
        
        # 社保扣除比例
        ttk.Label(frame, text="社保扣除比例:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        social_var = self.job_configs[job_type]['social_rate_var']
        ttk.Entry(frame, textvariable=social_var, width=15).grid(
            row=1, column=1, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        ttk.Label(frame, text="%").grid(row=1, column=2, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        
        # 个人所得税比例
        ttk.Label(frame, text="个人所得税比例:").grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        tax_var = self.job_configs[job_type]['tax_rate_var']
        ttk.Entry(frame, textvariable=tax_var, width=15).grid(
            row=2, column=1, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        ttk.Label(frame, text="%").grid(row=2, column=2, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        
        # 迟到扣款单价
        ttk.Label(frame, text="迟到扣款单价:").grid(row=3, column=0, sticky=tk.W, pady=(5, 0))
        late_var = self.job_configs[job_type]['late_deduction_rate_var']
        ttk.Entry(frame, textvariable=late_var, width=15).grid(
            row=3, column=1, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        ttk.Label(frame, text="元/次").grid(row=3, column=2, padx=(5, 0), sticky=tk.W, pady=(5, 0))
        
    def _browse_template(self, job_type: str):
        """浏览模板文件"""
        file_path = filedialog.askopenfilename(
            title=f"选择{job_type}工资模板文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.job_configs[job_type]['template_var'].set(file_path)
    
    def _load_current_config(self):
        """加载当前配置"""
        try:
            # 获取基础配置
            base_config = self.config.get('base_salary', {})
            floating_config = self.config.get('floating_salary', {})
            commission_config = self.config.get('commission_config', {})
            manual_config = self.config.get('manual_fees', {})
            other_config = self.config.get('other_config', {})
            
            # 获取职业特定配置
            job_specific_config_data = self.config.get('job_specific_config', {})
            
            # 为每个职业加载配置
            for job_type in self.job_types:
                job_config = self.job_configs[job_type]
                
                # 模板文件
                if job_type in self.existing_template_paths:
                    job_config['template_var'].set(self.existing_template_paths[job_type])
                
                # 基本配置（优先使用用户保存的职业特定配置）
                job_specific_data = job_specific_config_data.get(job_type, {})
                if job_specific_data:
                    # 使用用户保存的职业特定配置
                    job_config['base_salary_var'].set(str(job_specific_data.get('base_salary', 5000)))
                    job_config['floating_salary_var'].set(str(job_specific_data.get('floating_salary', 0)))
                else:
                    # 使用默认配置（根据职业类型设置不同的默认值）
                    job_default_config = JOB_SPECIFIC_CONFIG.get(job_type, {})
                    default_base_salary = job_default_config.get('default_base_salary', 5000)
                    job_config['base_salary_var'].set(str(base_config.get('default', default_base_salary)))
                    job_config['floating_salary_var'].set(str(floating_config.get('default', 0)))
                job_config['body_rate_var'].set(str(manual_config.get('body_rate', 60)))
                job_config['face_rate_var'].set(str(manual_config.get('face_rate', 80)))
                job_config['training_allowance_var'].set(str(other_config.get('training_allowance', 200)))
                job_config['social_rate_var'].set(str(other_config.get('social_security_rate', 8.0)))
                job_config['tax_rate_var'].set(str(other_config.get('personal_tax_rate', 3.0)))
                job_config['late_deduction_rate_var'].set(str(other_config.get('late_deduction_rate', 20.0)))
                
                # 提成比例（根据职业类型）
                if job_type == '服务总监':
                    rate = commission_config.get('expert_commission', {}).get('default_rate', 1.2)
                elif job_type == '服务老师':
                    rate = commission_config.get('service_commission', {}).get('default_rate', 1.5)
                else:  # 操作老师
                    rate = commission_config.get('operation_commission', {}).get('default_rate', 0.8)
                
                job_config['commission_rate_var'].set(str(rate))
                
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败: {str(e)}")
    
    def _reset_to_default(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置为默认配置吗？这将丢失所有自定义设置。"):
            self.config = DEFAULT_SALARY_CONFIG.copy()
            self._load_current_config()
    
    def _validate_config(self) -> bool:
        """验证配置"""
        try:
            for job_type in self.job_types:
                job_config = self.job_configs[job_type]
                
                # 验证数值字段
                float(job_config['base_salary_var'].get())
                float(job_config['floating_salary_var'].get())
                float(job_config['commission_rate_var'].get())
                float(job_config['body_rate_var'].get())
                float(job_config['face_rate_var'].get())
                float(job_config['training_allowance_var'].get())
                float(job_config['social_rate_var'].get())
                float(job_config['tax_rate_var'].get())
                float(job_config['late_deduction_rate_var'].get())
                
                # 验证模板文件路径
                template_path = job_config['template_var'].get().strip()
                if template_path and not os.path.exists(template_path):
                    messagebox.showerror("错误", f"{job_type}的模板文件路径不存在:\n{template_path}")
                    return False
            
            return True
            
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数值！")
            return False
    
    def _collect_config(self) -> Dict[str, Any]:
        """收集配置数据"""
        # 收集模板路径
        template_paths = {}
        for job_type in self.job_types:
            path = self.job_configs[job_type]['template_var'].get().strip()
            if path:
                template_paths[job_type] = path
        
        # 收集各职业的基础底薪配置
        # 注意：这里的default实际上不会被使用，因为计算逻辑会根据职业类型从JOB_SPECIFIC_CONFIG获取
        first_job = self.job_configs[self.job_types[0]]
        
        # 收集职业特定的基础底薪配置
        job_specific_salaries = {}
        for job_type in self.job_types:
            job_config = self.job_configs[job_type]
            job_specific_salaries[job_type] = {
                'base_salary': float(job_config['base_salary_var'].get()),
                'floating_salary': float(job_config['floating_salary_var'].get())
            }
        
        config = {
            'base_salary': {
                'default': float(first_job['base_salary_var'].get()),  # 保持向后兼容
                'special_rates': {}  # 已删除特殊人员底薪
            },
            'job_specific_config': job_specific_salaries,  # 新增：职业特定配置
            'floating_salary': {
                'default': float(first_job['floating_salary_var'].get()),
                'special_rates': {}
            },
            'commission_config': {
                'expert_commission': {
                    'default_rate': float(self.job_configs['服务总监']['commission_rate_var'].get()),
                    'default_quantity': 1,
                    'special_rates': {},
                    'special_quantities': {}
                },
                'service_commission': {
                    'default_rate': float(self.job_configs['服务老师']['commission_rate_var'].get()),
                    'default_quantity': 1,
                    'special_rates': {},
                    'special_quantities': {}
                },
                'operation_commission': {
                    'default_rate': float(self.job_configs['操作老师']['commission_rate_var'].get()),
                    'default_quantity': 1,
                    'special_rates': {},
                    'special_quantities': {}
                }
            },
            'commission_rates': {
                'expert_rate': float(self.job_configs['服务总监']['commission_rate_var'].get()),
                'service_rate': float(self.job_configs['服务老师']['commission_rate_var'].get()),
                'operation_rate': float(self.job_configs['操作老师']['commission_rate_var'].get())
            },
            'manual_fees': {
                'body_rate': float(first_job['body_rate_var'].get()),
                'face_rate': float(first_job['face_rate_var'].get())
            },
            'other_config': {
                'training_allowance': float(first_job['training_allowance_var'].get()),
                'social_security_rate': float(first_job['social_rate_var'].get()),
                'personal_tax_rate': float(first_job['tax_rate_var'].get()),
                'base_monthly_rest_days': 4,
                'current_month_holiday_days': 0,
                'late_deduction_rate': float(first_job['late_deduction_rate_var'].get())
            },
            'template_paths': template_paths
        }
        
        return config
    
    def _on_ok(self):
        """确定按钮事件"""
        if self._validate_config():
            self.result = self._collect_config()
            self.dialog.destroy()
    
    def _on_cancel(self):
        """取消按钮事件"""
        self.result = None
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict[str, Any]]:
        """显示对话框并返回结果"""
        self.dialog.wait_window()
        return self.result 