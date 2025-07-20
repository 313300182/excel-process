# -*- coding: utf-8 -*-
"""
Excel处理工具配置文件
定义数据提取和填充的规则
"""

import os
import sys

# 应用配置
APP_NAME = "Excel批量处理工具"
APP_VERSION = "1.0.0"

# 文件配置
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls']
TEMPLATE_FILENAME = 'template.xlsx'

# 源Excel数据提取配置
SOURCE_CONFIG = {
    # 数据开始的行列位置 (1-based indexing)
    'start_row': 11,
    'start_col': 1,
    
    # 字段映射 - 定义要提取的数据列
    'fields': {
        'product_name': 2,   # B列 - 品名及规格
        'unit': 4,           # D列 - 单位
        'quantity': 5,       # E列 - 数量
        'amount': 7,         # G列 - 金额
    },
    
    # 结束标识 - 遇到这些内容时停止读取
    'end_markers': ['合计', '合  计', '小计', '小 计', '总计', '总 计'],
    
    # 跳过空行配置
    'skip_empty_rows': True,
    
    # 必填字段 - 这些字段都为空时认为是空行
    'required_fields': ['product_name'],  # 至少品名及规格不能为空
    
    # 最大读取行数 (防止无限循环)
    'max_rows': 1000,
    
    # 数字格式处理
    'number_format': {
        'remove_comma': True,  # 移除逗号分隔符
        'decimal_places': 2,   # 保留小数位数
    }
}

# 目标模板填充配置
TARGET_CONFIG = {
    # 模板中的填充位置 - 支持多行数据，从第4行开始写入明细数据（保留前3行）
    'fill_positions': {
        'product_name': 1,   # A列 - 项目名称
        'unit': 4,           # D列 - 单位
        'quantity': 5,       # E列 - 商品数量
        'amount': 7,         # G列 - 金额
        'code': 2,       # B列 - 商品和服务税收分类编码
        'tax_rate': 8, # H列 - 折扣金额
    },
    
    # 数据起始行（从第4行开始，保留前3行表头和说明）
    'data_start_row': 4,
    
    # 表头信息（不写入，保持模板原有表头）
    'headers': {},
    
    # 默认值配置 - 固定填充的字段（不从源文件提取）
    'default_values': {
        'code': '1060105010000000000',
        'tax_rate': 0.13,
    },

    # 是否在末尾添加合计行
    'add_total_row': False,
    'total_label': '合计',
    'total_label_column': 1,  # 合计标签放在A列
    'total_amount_column': 7, # 合计金额放在G列
}

# 输出配置
OUTPUT_CONFIG = {
    # 输出文件命名格式
    'filename_format': '{original_name}-输出-{timestamp}.xlsx',
    
    # 时间戳格式
    'timestamp_format': '%Y-%m-%d',
    
    # 是否创建输出目录
    'create_output_dir': True,
    
    # 默认输出目录名
    'default_output_dir': '输出模板-{timestamp}',
    
    # 输出目录时间戳格式
    'dir_timestamp_format': '%Y-%m-%d'
}

# UI配置
UI_CONFIG = {
    'window_title': APP_NAME,
    'window_size': '800x600',
    'font_family': '微软雅黑',
    'font_size': 10
}

# 日志配置（已移除文件日志，仅UI显示）
LOG_CONFIG = {
    'level': 'INFO',
    'format': '%(asctime)s - %(levelname)s - %(message)s'
}

def get_template_path():
    """获取模板文件路径（支持打包后的环境）"""
    # 检查是否在PyInstaller打包环境中
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # 打包后的环境，模板文件在临时目录中
        base_dir = sys._MEIPASS
    else:
        # 开发环境
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    return os.path.join(base_dir, 'templates', TEMPLATE_FILENAME)

