# -*- coding: utf-8 -*-
"""
老师分组Excel写入器模型
负责将老师分组数据写入多个sheet的Excel文件
"""

import os
import logging
import shutil
import stat
from datetime import datetime
from typing import List, Dict, Any, Optional
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.styles import Font, PatternFill, Alignment

from config.teacher_splitter_settings import (
    TEACHER_OUTPUT_CONFIG, 
    TEACHER_FILE_CONFIG
)


class TeacherExcelWriter:
    """老师分组Excel文件写入器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def create_teacher_grouped_file(self, 
                                   all_data: List[Dict[str, Any]], 
                                   output_dir: str, 
                                   original_filename: str,
                                   source_file_path: str) -> Optional[str]:
        """
        创建按老师分组的多sheet Excel文件
        
        Args:
            all_data: 所有数据列表
            output_dir: 输出目录
            original_filename: 原始文件名
            source_file_path: 源文件路径
            
        Returns:
            str: 生成的文件路径，失败返回None
        """
        try:
            if not all_data:
                self.logger.warning("没有数据需要写入")
                return None
                
            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)
            
            # 生成输出文件名
            output_filename = self._generate_output_filename(original_filename)
            output_path = os.path.join(output_dir, output_filename)
            
            # 按老师分组数据
            grouped_data = self._group_data_by_teachers(all_data)
            
            # 复制源文件作为基础
            self.logger.info(f"复制源文件: {source_file_path} -> {output_path}")
            shutil.copy2(source_file_path, output_path)
            
            # 打开复制的文件
            workbook = load_workbook(output_path)
            
            # 重命名原始sheet为"原始数据"
            if workbook.worksheets:
                original_sheet = workbook.worksheets[0]
                original_sheet.title = "原始数据"
                self.logger.info(f"重命名原始sheet为: 原始数据")
            
            sheet_count = 0
            # 为每个老师创建sheet
            for teacher_info, teacher_data in grouped_data.items():
                sheet_name = self._generate_sheet_name(teacher_info)
                worksheet = workbook.create_sheet(sheet_name)
                
                # 写入表头
                self._write_headers(worksheet)
                
                # 写入数据
                self._write_teacher_data(worksheet, teacher_data)
                
                # 添加合计行
                if TEACHER_OUTPUT_CONFIG.get('add_total_row', False):
                    self._write_total_row(worksheet, teacher_data)
                
                # 设置列宽
                self._adjust_column_width(worksheet)
                
                sheet_count += 1
                self.logger.info(f"创建sheet: {sheet_name}, 数据行数: {len(teacher_data)}")
            
            # 保存文件
            workbook.save(output_path)
            workbook.close()
            
            self.logger.info(f"成功创建老师分组文件: {output_path}, 原始数据1个sheet + {sheet_count}个老师分组sheet")
            return output_path
            
        except Exception as e:
            self.logger.error(f"创建老师分组文件失败: {e}")
            return None
    
    def _group_data_by_teachers(self, all_data: List[Dict[str, Any]]) -> Dict[tuple, List[Dict[str, Any]]]:
        """
        按老师分组数据
        
        Args:
            all_data: 所有数据
            
        Returns:
            Dict: 分组后的数据，key为(老师姓名, 角色)元组
        """
        grouped_data = {}
        teacher_columns = ['service_director', 'service_teacher', 'operation_teacher']
        role_names = TEACHER_FILE_CONFIG['role_names']
        empty_name = TEACHER_FILE_CONFIG['empty_teacher_name']
        
        for row in all_data:
            # 为每个老师角色创建分组
            for role in teacher_columns:
                teacher_name = row.get(role)
                
                # 处理空值
                if not teacher_name or str(teacher_name).strip() == '':
                    teacher_name = empty_name
                else:
                    teacher_name = str(teacher_name).strip()
                
                # 创建分组key
                teacher_key = (teacher_name, role_names[role])
                
                if teacher_key not in grouped_data:
                    grouped_data[teacher_key] = []
                
                # 添加数据到对应分组
                grouped_data[teacher_key].append(row)
        
        # 按老师姓名排序
        sorted_groups = dict(sorted(grouped_data.items(), key=lambda x: (x[0][0], x[0][1])))
        
        return sorted_groups
    
    def _generate_sheet_name(self, teacher_info: tuple) -> str:
        """
        生成sheet名称
        
        Args:
            teacher_info: (老师姓名, 角色)元组
            
        Returns:
            str: sheet名称
        """
        teacher_name, role = teacher_info
        sheet_name = TEACHER_FILE_CONFIG['sheet_name_format'].format(
            teacher_name=teacher_name, 
            role=role
        )
        
        # Excel sheet名称限制
        # 不能超过31个字符，不能包含特殊字符
        invalid_chars = ['/', '\\', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            sheet_name = sheet_name.replace(char, '_')
        
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:28] + '...'
        
        return sheet_name
    
    def _generate_output_filename(self, original_filename: str) -> str:
        """
        生成输出文件名
        
        Args:
            original_filename: 原始文件名
            
        Returns:
            str: 生成的文件名
        """
        name_without_ext = os.path.splitext(original_filename)[0]
        timestamp = datetime.now().strftime(TEACHER_FILE_CONFIG['timestamp_format'])
        
        filename = TEACHER_FILE_CONFIG['filename_format'].format(
            original_name=name_without_ext,
            timestamp=timestamp
        )
        
        return filename
    
    def _write_headers(self, worksheet) -> None:
        """
        写入表头
        
        Args:
            worksheet: 工作表对象
        """
        try:
            headers = TEACHER_OUTPUT_CONFIG['headers']
            
            # 创建表头样式
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center")
            
            for (row, col), header_text in headers.items():
                cell = worksheet.cell(row=row, column=col, value=header_text)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                
        except Exception as e:
            self.logger.error(f"写入表头失败: {e}")
    
    def _write_teacher_data(self, worksheet, teacher_data: List[Dict[str, Any]]) -> int:
        """
        写入老师数据
        
        Args:
            worksheet: 工作表对象
            teacher_data: 老师数据列表
            
        Returns:
            int: 最后写入的行号
        """
        try:
            output_columns = TEACHER_OUTPUT_CONFIG['output_columns']
            data_start_row = TEACHER_OUTPUT_CONFIG['data_start_row']
            
            for row_index, data_row in enumerate(teacher_data):
                current_row = data_start_row + row_index
                
                for field_name, col_index in output_columns.items():
                    value = data_row.get(field_name, "")
                    
                    # 处理数值字段
                    if field_name in ['commission', 'experience_card'] and value:
                        try:
                            if isinstance(value, str):
                                value = value.replace(',', '').replace('，', '')
                            value = float(value) if value else 0
                        except (ValueError, TypeError):
                            value = 0
                    
                    worksheet.cell(row=current_row, column=col_index, value=value)
                    
            return data_start_row + len(teacher_data) - 1
                        
        except Exception as e:
            self.logger.error(f"写入老师数据失败: {e}")
            return data_start_row
    
    def _write_total_row(self, worksheet, teacher_data: List[Dict[str, Any]]) -> None:
        """
        写入合计行
        
        Args:
            worksheet: 工作表对象
            teacher_data: 老师数据列表
        """
        try:
            data_start_row = TEACHER_OUTPUT_CONFIG['data_start_row']
            total_row = data_start_row + len(teacher_data)
            
            total_label = TEACHER_OUTPUT_CONFIG['total_label']
            total_label_column = TEACHER_OUTPUT_CONFIG['total_label_column']
            total_amount_column = TEACHER_OUTPUT_CONFIG['total_amount_column']
            total_card_column = TEACHER_OUTPUT_CONFIG['total_card_column']
            
            # 写入合计标签
            worksheet.cell(row=total_row, column=total_label_column, value=total_label)
            
            # 计算合计金额
            total_commission = 0
            total_cards = 0
            
            for data_row in teacher_data:
                # 计算实收业绩合计
                commission = data_row.get('commission', 0)
                if commission:
                    try:
                        if isinstance(commission, str):
                            commission = commission.replace(',', '').replace('，', '')
                        total_commission += float(commission)
                    except (ValueError, TypeError):
                        pass
                
                # 计算体验卡合计
                cards = data_row.get('experience_card', 0)
                if cards:
                    try:
                        if isinstance(cards, str):
                            cards = cards.replace(',', '').replace('，', '')
                        total_cards += float(cards)
                    except (ValueError, TypeError):
                        pass
            
            # 写入合计数据
            worksheet.cell(row=total_row, column=total_amount_column, value=total_commission)
            worksheet.cell(row=total_row, column=total_card_column, value=total_cards)
            
            # 设置合计行样式
            for col in range(1, 9):  # A-H列
                cell = worksheet.cell(row=total_row, column=col)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            
        except Exception as e:
            self.logger.error(f"写入合计行失败: {e}")
    
    def _adjust_column_width(self, worksheet) -> None:
        """
        调整列宽
        
        Args:
            worksheet: 工作表对象
        """
        try:
            # 设置各列宽度
            column_widths = {
                'A': 12,  # 日期
                'B': 15,  # 客户
                'C': 12,  # 服务总监
                'D': 12,  # 服务老师
                'E': 12,  # 操作老师
                'F': 20,  # 店名
                'G': 15,  # 实收业绩
                'H': 12,  # 体验卡
            }
            
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
                
        except Exception as e:
            self.logger.error(f"调整列宽失败: {e}") 