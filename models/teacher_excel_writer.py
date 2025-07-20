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
        创建按老师角色分类的多sheet Excel文件
        
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
            
            # 按角色分组数据
            grouped_data = self._group_data_by_role(all_data)
            
            # 创建新的工作簿
            workbook = Workbook()
            
            # 删除默认的sheet
            if workbook.worksheets:
                workbook.remove(workbook.worksheets[0])
            
            sheet_count = 0
            role_names = TEACHER_FILE_CONFIG['role_names']
            
            # 为每个角色创建sheet
            for role, role_data in grouped_data.items():
                if not role_data:  # 跳过空数据的角色
                    continue
                    
                sheet_name = role_names[role]
                worksheet = workbook.create_sheet(sheet_name)
                
                # 写入表头
                self._write_headers(worksheet)
                
                # 写入数据
                self._write_teacher_data(worksheet, role_data)
                
                # 添加合计行
                if TEACHER_OUTPUT_CONFIG.get('add_total_row', False):
                    self._write_total_row(worksheet, role_data)
                
                # 设置列宽
                self._adjust_column_width(worksheet)
                
                sheet_count += 1
                self.logger.info(f"创建sheet: {sheet_name}, 数据行数: {len(role_data)}")
            
            # 保存文件
            workbook.save(output_path)
            workbook.close()
            
            self.logger.info(f"成功创建老师角色分类文件: {output_path}, 共{sheet_count}个角色分类sheet")
            return output_path
            
        except Exception as e:
            self.logger.error(f"创建老师角色分类文件失败: {e}")
            return None

    
    def _group_data_by_role(self, all_data: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """
        按角色分组数据
        
        Args:
            all_data: 所有数据
            
        Returns:
            Dict: 分组后的数据，key为角色名称
        """
        grouped_data = {
            '服务总监': [],
            '服务老师': [],
            '操作老师': [],
            '店家分类': []
        }
        
        role_mapping = {
            'service_director': '服务总监',
            'service_teacher': '服务老师',
            'operation_teacher': '操作老师'
        }
        
        empty_name = TEACHER_FILE_CONFIG['empty_teacher_name']
        
        for row in all_data:
            # 为每个角色检查是否有值，有值则加入对应的sheet
            for role_field, role_name in role_mapping.items():
                teacher_name = row.get(role_field)
                
                # 只有该角色字段有值时才加入对应的sheet
                if teacher_name and str(teacher_name).strip() != '':
                    teacher_name = str(teacher_name).strip()
                    # 不修改原数据，直接添加到对应角色的数据列表
                    grouped_data[role_name].append(row)
            
            # 店家分类：所有数据都加入店家分类sheet
            grouped_data['店家分类'].append(row)
        
        return grouped_data
    

    
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
            
            # 添加实收业绩和体验卡合计行
            grand_total_row = total_row + 1
            grand_total = total_commission + total_cards
            
            # 写入总合计标签和数值
            worksheet.cell(row=grand_total_row, column=total_label_column, value="实收业绩和体验卡合计")
            worksheet.cell(row=grand_total_row, column=total_amount_column, value=grand_total)
            
            # 设置总合计行样式
            for col in range(1, 9):  # A-H列
                cell = worksheet.cell(row=grand_total_row, column=col)
                cell.font = Font(bold=True, color="FF0000")  # 红色字体以示区别
                cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # 黄色背景
            
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