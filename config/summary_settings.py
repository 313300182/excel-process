# -*- coding: utf-8 -*-
"""
汇总功能配置文件
定义汇总数据的提取和输出规则
"""

# 汇总数据配置
SUMMARY_CONFIG = {
    # 品名及规格字段映射策略
    'item_name_strategy': 'notes_first',  # 'notes_first': 优先使用开单明细, 'customer_service': 使用客户+服务
    
    # 默认单位
    'default_unit': '个',
    
    # 数量计算方式
    'quantity_method': 'count',  # 'count': 按记录数量, 'amount_based': 基于金额计算
    
    # 金额字段选择
    'amount_field': 'commission',  # 'commission': 实收业绩, 'order_amount': 开单金额
    
    # 单价计算方式
    'unit_price_method': 'average',  # 'average': 平均值, 'first': 第一个记录的值
    
    # 最小金额阈值（低于此金额的项目不包含在汇总中）
    'min_amount_threshold': 0.01,
    
    # 排序方式
    'sort_by': 'amount_desc',  # 'amount_desc': 按金额降序, 'amount_asc': 按金额升序, 'name': 按名称
}

# 汇总输出配置
SUMMARY_OUTPUT_CONFIG = {
    # 文件命名格式
    'filename_format': '{original_name}_汇总_{timestamp}.xlsx',
    
    # 时间戳格式
    'timestamp_format': '%Y-%m-%d',
    
    # 工作表名称
    'worksheet_name': '合同开票信息提取',
    
    # 标题
    'title': '合同开票信息提取',
    
    # 表头配置
    'headers': {
        'A2': '序号',
        'B2': '品名及规格',
        'C2': '单位', 
        'D2': '数量',
        'E2': '单价',
        'F2': '金额'
    },
    
    # 列宽配置
    'column_widths': {
        'A': 8,   # 序号
        'B': 25,  # 品名及规格
        'C': 8,   # 单位
        'D': 12,  # 数量
        'E': 15,  # 单价
        'F': 15,  # 金额
    },
    
    # 样式配置
    'styles': {
        'title': {
            'font_size': 16,
            'font_bold': True,
            'alignment': 'center'
        },
        'header': {
            'font_size': 12,
            'font_bold': True,
            'font_color': 'FFFFFF',
            'background_color': '366092',
            'alignment': 'center'
        },
        'data': {
            'font_size': 11,
            'alignment_number': 'right',
            'alignment_text': 'left',
            'alignment_center': 'center'
        },
        'total': {
            'font_size': 12,
            'font_bold': True,
            'background_color': 'FFE4B5',
            'alignment': 'center'
        }
    }
}

# 品名及规格提取规则
ITEM_NAME_RULES = {
    # 优先级规则 (数字越小优先级越高)
    'rules': [
        {
            'priority': 1,
            'source_field': 'notes',
            'condition': 'not_empty',
            'process': 'clean_text'
        },
        {
            'priority': 2, 
            'source_field': 'customer',
            'condition': 'not_empty',
            'process': 'add_service_suffix'  # 添加"服务"后缀
        },
        {
            'priority': 3,
            'source_field': None,
            'condition': 'fallback',
            'process': 'default_name',
            'default_value': '其他服务'
        }
    ],
    
    # 文本处理规则
    'text_processing': {
        'max_length': 50,  # 最大长度
        'remove_chars': ['\n', '\r', '\t'],  # 需要移除的字符
        'replace_patterns': {
            '  ': ' ',  # 多个空格替换为单个空格
        }
    }
} 