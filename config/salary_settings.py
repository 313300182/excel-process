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
        
        # 应发项目映射（包含数量、单价、金额）
        'salary_items': {
            # 基本底薪
            'base_salary_quantity': 'C6',     # 基本底薪数量
            'base_salary_rate': 'D6',         # 基本底薪单价
            'base_salary': 'E6',              # 基本底薪金额
            
            # 浮动底薪
            'floating_salary_quantity': 'C7', # 浮动底薪数量
            'floating_salary_rate': 'D7',     # 浮动底薪单价
            'floating_salary': 'E7',          # 浮动底薪金额
            
            # 专家提成（服务总监使用）
            'expert_commission_quantity': 'C8', # 专家提成数量
            'expert_commission_rate': 'D8',     # 专家提成比例
            'expert_commission': 'E8',          # 专家提成金额
            
            # 服务提成（服务老师使用）
            'service_commission_quantity': 'C8', # 服务提成数量
            'service_commission_rate': 'D8',     # 服务提成比例
            'service_commission': 'E8',          # 服务提成金额

            # 操作提成（操作老师使用）
            'operation_commission_quantity': 'C8', # 操作提成数量
            'operation_commission_rate': 'D8',     # 操作提成比例
            'operation_commission': 'E8',          # 操作提成金额

            # 培训补贴
            'training_allowance_quantity': 'C10', # 培训补贴数量
            'training_allowance_rate': 'D10',     # 培训补贴单价
            'training_allowance': 'E10',          # 培训补贴金额
            
            # 身体部位手工费
            'body_manual_fee_quantity': 'C11', # 身体部位数量
            'body_manual_fee_rate': 'D11',     # 身体部位单价
            'body_manual_fee': 'E11',          # 身体部位手工费金额
            
            # 面部手工费
            'face_manual_fee_quantity': 'C12', # 面部数量
            'face_manual_fee_rate': 'D12',     # 面部单价
            'face_manual_fee': 'E12',          # 面部手工费金额
            
            # 应发合计
            'total_salary': 'E13'             # 应发合计
        },
        
        # 考勤信息映射
        # 'attendance_info': {
        #     'work_days': 'G2',      # 上班天数
        #     'rest_days': 'G3',      # 休息天数
        #     'late_count': 'G4',     # 迟到次数
        #     'training_days': 'G5',  # 培训天数
        # },
        
        # 扣减项目映射  
        'deduction_items': {
            'absent_deduction_quantity': 'C16',  # 缺勤天数
            'absent_deduction_rate': 'D16',      # 缺勤单价
            'absent_deduction': 'E16',           # 缺勤扣款金额
            'late_deduction_quantity': 'C17',    # 迟到次数
            'late_deduction_rate': 'D17',        # 迟到单价
            'late_deduction': 'E17',             # 迟到扣款金额
            'social_security_quantity': 'C18',   # 社保数量
            'social_security_rate': 'D18',       # 社保单价
            'social_security': 'E18',            # 社保金额
            'personal_tax_quantity': 'C19',      # 个税数量
            'personal_tax_rate': 'D19',          # 个税单价
            'personal_tax': 'E19',               # 个税金额
            'total_deduction': 'E20'             # 扣减小计
        },
        
        # 实发工资
        'net_salary': 'E22'
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
    
    # 提成配置
    'commission_config': {
        # 专家提成配置（服务总监）
        'expert_commission': {
            'default_rate': 1.20,        # 默认专家提成比例 (%)
            'default_quantity': 1,       # 默认专家提成数量
            'special_rates': {},         # 特殊人员专家提成比例
            'special_quantities': {}     # 特殊人员专家提成数量
        },
        # 服务提成配置（服务老师）
        'service_commission': {
            'default_rate': 1.50,        # 默认服务提成比例 (%)
            'default_quantity': 1,       # 默认服务提成数量
            'special_rates': {},         # 特殊人员服务提成比例
            'special_quantities': {}     # 特殊人员服务提成数量
        },
        # 操作提成配置（操作老师）
        'operation_commission': {
            'default_rate': 0.80,        # 默认操作提成比例 (%)
            'default_quantity': 1,       # 默认操作提成数量
            'special_rates': {},         # 特殊人员操作提成比例
            'special_quantities': {}     # 特殊人员操作提成数量
        }
    },
    
    # 提成比例配置（保持向后兼容）
    'commission_rates': {
        'expert_rate': 1.20,     # 专家提成比例 (%)
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
        'social_security_rate': 505.26,  # 社保单价（元）
        'personal_tax_rate': 3.0,      # 个人所得税比例 (%)
        'base_monthly_rest_days': 4,   # 基础月休天数
        'current_month_holiday_days': 0,  # 当月节日休息天数（可配置）
        'late_deduction_rate': 20,     # 迟到扣款单价（正数）
    }
}

# 职业特殊配置
JOB_SPECIFIC_CONFIG = {
    '服务总监': {
        'default_base_salary': 8000,   # 默认基础底薪
        'special_allowance': 1000      # 特殊补贴
    },
    '服务老师': {
        'default_base_salary': 5000,   # 默认基础底薪
        'special_allowance': 500
    },
    '操作老师': {
        'default_base_salary': 5000,   # 默认基础底薪
        'special_allowance': 0
    }
}

# 配置文件保存路径
SALARY_USER_CONFIG_FILE = 'salary_user_config.json' 