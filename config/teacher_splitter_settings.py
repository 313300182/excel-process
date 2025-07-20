# -*- coding: utf-8 -*-
"""
老师分组Excel处理配置文件
定义按老师分组提取和输出数据的规则
"""

import os

# 老师分组源数据配置
TEACHER_SOURCE_CONFIG = {
    # 数据开始的行号 (1-based indexing) 
    'start_row': 3,
    
    # 字段映射 - 定义要提取的数据列
    'fields': {
        'date': 1,              # A列 - 日期
        'customer': 2,          # B列 - 客户
        'service_director': 3,  # C列 - 服务总监
        'service_teacher': 4,   # D列 - 服务老师
        'operation_teacher': 5, # E列 - 操作老师
        'store_name': 6,        # F列 - 店名
        'commission': 7,        # G列 - 实收业绩 
        'experience_card': 8,   # H列 - 体验卡
        'notes': 9,            # I列 - 开单明细
        'public_revenue': 10,   # J列 - 公收
        'private_revenue': 11,  # K列 - 应收
    },
    
    # 老师角色列 - 用于分组
    'teacher_columns': {
        'service_director': 3,  # 服务总监
        'service_teacher': 4,   # 服务老师  
        'operation_teacher': 5, # 操作老师
    },
    
    # 结束标识 - 遇到这些内容时停止读取
    'end_markers': ['合计', '合  计', '小计', '小 计', '总计', '总 计'],
    
    # 跳过空行配置
    'skip_empty_rows': True,
    
    # 必填字段 - 这些字段都为空时认为是空行
    'required_fields': ['customer'],  # 至少客户不能为空
    
    # 最大读取行数
    'max_rows': 2000,
}

# 老师分组输出配置
TEACHER_OUTPUT_CONFIG = {
    # 输出模板中的列映射
    'output_columns': {
        'date': 1,              # A列 - 日期
        'customer': 2,          # B列 - 客户
        'service_director': 3,  # C列 - 服务总监
        'service_teacher': 4,   # D列 - 服务老师
        'operation_teacher': 5, # E列 - 操作老师
        'store_name': 6,        # F列 - 店名
        'commission': 7,        # G列 - 实收业绩
        'experience_card': 8,   # H列 - 体验卡
    },
    
    # 数据起始行（从第2行开始，保留表头）
    'data_start_row': 2,
    
    # 表头
    'headers': {
        (1, 1): '日期',
        (1, 2): '客户', 
        (1, 3): '服务总监',
        (1, 4): '服务老师',
        (1, 5): '操作老师',
        (1, 6): '店名',
        (1, 7): '实收业绩',
        (1, 8): '体验卡',
    },
    
    # 是否添加合计行
    'add_total_row': True,
    'total_label': '合计',
    'total_label_column': 2,  # 合计标签放在B列
    'total_amount_column': 7, # 合计金额放在G列
    'total_card_column': 8,   # 体验卡合计放在H列
}

# 输出文件配置
TEACHER_FILE_CONFIG = {
    # 输出文件命名格式
    'filename_format': '{original_name}-老师分组-{timestamp}.xlsx',
    
    # Sheet命名格式
    'sheet_name_format': '{teacher_name}({role})',
    
    # 角色中文名称映射
    'role_names': {
        'service_director': '服务总监',
        'service_teacher': '服务老师', 
        'operation_teacher': '操作老师',
    },
    
    # 时间戳格式
    'timestamp_format': '%Y-%m-%d',
    
    # 空值处理
    'empty_teacher_name': '未分配',
}

# 注意：老师分组功能不使用模板文件，直接基于源文件创建新的Excel文件 