#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller hook for openpyxl
解决openpyxl模块缺失问题
"""

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# 收集openpyxl的所有子模块
hiddenimports = collect_submodules('openpyxl')

# 收集openpyxl的数据文件
datas = collect_data_files('openpyxl')

# 手动添加经常缺失的模块
hiddenimports += [
    'openpyxl.cell_writer',
    'openpyxl.workbook.workbook',
    'openpyxl.worksheet.worksheet', 
    'openpyxl.styles.styles',
    'openpyxl.styles.numbers',
    'openpyxl.styles.borders',
    'openpyxl.styles.fills',
    'openpyxl.styles.fonts',
    'openpyxl.styles.alignment',
    'openpyxl.styles.protection',
    'openpyxl.chart.chart',
    'openpyxl.comments.comments',
    'openpyxl.drawing.drawing',
    'openpyxl.packaging.relationship',
    'openpyxl.packaging.manifest',
    'openpyxl.packaging.core',
    'openpyxl.packaging.extended',
    'openpyxl.packaging.custom',
    'openpyxl.utils.exceptions',
    'openpyxl.utils.indexed_list',
    'openpyxl.utils.datetime',
    'openpyxl.utils.units',
    'openpyxl.utils.dataframe',
    'openpyxl.xml.functions',
    'openpyxl.xml.constants',
    'openpyxl.reader.excel',
    'openpyxl.writer.excel',
    'et_xmlfile',
    'et_xmlfile.xmlfile',
] 