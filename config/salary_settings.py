# -*- coding: utf-8 -*-
"""
工资Excel处理配置文件
定义工资数据提取和填充的规则
"""

import os

# 工资处理配置
SALARY_CONFIG = {
    # 支持的职业类型
    'job_types': ['服务总监', '服务老师', '操作老师'],
    
    # 模板文件配置
    'templates': {
        '服务总监': 'salary_template_supervisor.xlsx',
        '服务老师': 'salary_template_service.xlsx', 
        '操作老师': 'salary_template_operation.xlsx'
    },
    
    # 源文件数据提取配置
    'source_extraction': {
        # 员工信息提取位置
        'employee_info': {
            'name_cell': 'B2',  # 姓名位置
            'month_cell': 'D2',  # 月份位置
        },
        
        # 数据提取区域
        'data_start_row': 5,  # 数据开始行
        'data_start_col': 1,  # 数据开始列
        
        # 字段映射
        'fields': {
            'category': 1,      # A列 - 类别
            'project': 2,       # B列 - 项目
            'quantity': 3,      # C列 - 数量/基数
            'rate': 4,          # D列 - 比率/单价
            'amount': 5,        # E列 - 金额
            'note': 6,          # F列 - 备注
        },
        
        # 识别关键字段
        'key_fields': {
            'base_salary': '基本底薪',
            'floating_salary': '浮动底薪', 
            'service_commission': '服务提成',
            'operation_commission': '操作提成',
            'body_manual_fee': '身体部位',
            'face_manual_fee': '面部',
            'training_allowance': '培训补贴'
        },
        
        # 结束标识
        'end_markers': ['应发合计', '扣减项目', '实发工资'],
        
        # 最大读取行数
        'max_rows': 100
    },
    
    # 目标模板填充配置
    'template_mapping': {
        # 基本信息映射
        'employee_name': 'B2',  # 姓名
        'month': 'D2',          # 月份
        
        # 应发项目映射
        'salary_items': {
            'base_salary': 'E6',        # 基本底薪
            'floating_salary': 'E7',    # 浮动底薪
            'service_commission': 'E8', # 服务提成
            'operation_commission': 'E9', # 操作提成
            'training_allowance': 'E10', # 培训补贴
            'body_manual_fee': 'E11',   # 身体部位手工费
            'face_manual_fee': 'E12',   # 面部手工费
            'total_salary': 'E13'       # 应发合计
        },
        
        # 手工费数量和单价映射
        'manual_fee_details': {
            'body_quantity': 'C11',     # 身体部位数量
            'body_rate': 'D11',         # 身体部位单价
            'face_quantity': 'C12',     # 面部数量
            'face_rate': 'D12',         # 面部单价
        },
        
        # 扣减项目映射  
        'deduction_items': {
            'late_deduction': 'E16',    # 考勤扣款
            'absent_deduction': 'E17',  # 迟到扣款
            'social_security': 'E18',   # 社保
            'personal_tax': 'E19',      # 个人所得税
            'total_deduction': 'E20'    # 扣减小计
        },
        
        # 实发工资
        'net_salary': 'E23'
    }
}

# 默认配置值
DEFAULT_SALARY_CONFIG = {
    # 基本底薪配置
    'base_salary': {
        'default': 5000,  # 默认基本底薪
        'special_rates': {}  # 特殊人员底薪 {姓名: 底薪}
    },
    
    # 浮动底薪配置
    'floating_salary': {
        'default': 0,
        'special_rates': {}  # 特殊人员浮动底薪
    },
    
    # 提成比例配置
    'commission_rates': {
        'service_rate': 1.50,    # 服务提成比例 (%)
        'operation_rate': 0.80   # 操作提成比例 (%)
    },
    
    # 手工费配置
    'manual_fees': {
        'body_rate': 60,   # 身体部位手工费单价
        'face_rate': 80    # 面部手工费单价  
    },
    
    # 其他配置
    'other_config': {
        'training_allowance': 200,  # 培训补贴标准
        'social_security_rate': 8.0,  # 社保扣除比例 (%)
        'personal_tax_rate': 3.0      # 个人所得税比例 (%)
    }
}

# 职业特殊配置
JOB_SPECIFIC_CONFIG = {
    '服务总监': {
        'base_multiplier': 1.5,     # 基础倍率
        'commission_bonus': 0.2,    # 提成奖金
        'special_allowance': 1000   # 特殊补贴
    },
    '服务老师': {
        'base_multiplier': 1.2,
        'commission_bonus': 0.1,
        'special_allowance': 500
    },
    '操作老师': {
        'base_multiplier': 1.0,
        'commission_bonus': 0.0,
        'special_allowance': 0
    }
}

# 配置文件保存路径
SALARY_USER_CONFIG_FILE = 'salary_user_config.json' 