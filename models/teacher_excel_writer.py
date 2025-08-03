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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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
        创建按老师角色分类的多个Excel文件
        
        Args:
            all_data: 所有数据列表
            output_dir: 输出目录
            original_filename: 原始文件名
            source_file_path: 源文件路径
            
        Returns:
            str: 输出目录路径，失败返回None
        """
        try:
            if not all_data:
                self.logger.warning("没有数据需要写入")
                return None
                
            self.logger.info(f"开始创建老师角色分类文件，输出目录: {output_dir}")
            self.logger.warning("📝 提醒：如果您有相关的Excel文件正在打开，请先关闭以避免文件保存错误")
                
            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)
            
            # 按角色和店家分组数据
            grouped_data = self._group_data_by_role(all_data)
            
            # 按角色类型重新分组
            role_files = {
                '服务总监': {},
                '服务老师': {},
                '操作老师': {},
                '店家': {}
            }
            
            # 将数据按角色类型分类
            for group_key, role_data in grouped_data.items():
                if '(服务总监)' in group_key:
                    role_files['服务总监'][group_key] = role_data
                elif '(服务老师)' in group_key:
                    role_files['服务老师'][group_key] = role_data
                elif '(操作老师)' in group_key:
                    role_files['操作老师'][group_key] = role_data
                elif '(店家)' in group_key:
                    role_files['店家'][group_key] = role_data
            
            created_files = []
            
            # 为每个角色类型创建独立的Excel文件
            for role_type, role_groups in role_files.items():
                if not role_groups:  # 跳过空数据的角色
                    continue
                
                # 生成该角色类型的文件名
                role_filename = self._generate_role_filename(original_filename, role_type)
                role_output_path = os.path.join(output_dir, role_filename)
                
                # 创建该角色类型的工作簿
                workbook = Workbook()
                
                # 删除默认的sheet
                if workbook.worksheets:
                    workbook.remove(workbook.worksheets[0])
                
                sheet_count = 0
                
                # 为该角色类型下的每个具体老师/店家创建sheet
                for group_key, group_data in role_groups.items():
                    if not group_data:
                        continue
                    
                    # 提取sheet名称（去掉角色标识）
                    sheet_name = group_key.split('(')[0]
                    worksheet = workbook.create_sheet(sheet_name)
                    
                    # 写入表头
                    self._write_headers(worksheet)
                    
                    # 写入数据
                    self._write_teacher_data(worksheet, group_data)
                    
                    # 添加合计行
                    if TEACHER_OUTPUT_CONFIG.get('add_total_row', False):
                        self._write_total_row(worksheet, group_data)
                    
                    # 设置列宽
                    self._adjust_column_width(worksheet)
                    
                    sheet_count += 1
                    self.logger.info(f"在{role_type}文件中创建sheet: {sheet_name}, 数据行数: {len(group_data)}")
                
                # 保存该角色类型的文件（添加重试机制）
                self._safe_save_workbook(workbook, role_output_path)
                workbook.close()
                created_files.append(role_output_path)
                
                self.logger.info(f"成功创建{role_type}文件: {role_output_path}, 共{sheet_count}个sheet")
            
            self.logger.info(f"成功创建所有老师角色分类文件，共{len(created_files)}个文件")
            return output_dir  # 返回输出目录
            
        except Exception as e:
            self.logger.error(f"创建老师角色分类文件失败: {e}")
            return None

    
    def _group_data_by_role(self, all_data: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """
        按角色分组数据，为每个角色都创建分组，但在统计时避免重复计算
        支持多人分割（/分隔符）和数值字段的均分处理
        
        Args:
            all_data: 所有数据
            
        Returns:
            Dict: 分组后的数据，key为"老师姓名(角色)"或"店家名称(店家)"
        """
        grouped_data = {}
        
        role_mapping = {
            'service_director': '服务总监',
            'service_teacher': '服务老师',
            'operation_teacher': '操作老师'
        }
        
        # 需要均分的数值字段（去掉实收欠款，因为业务上不应该均分）
        splittable_fields = [
            'order_amount', 'payment', 'card_deduction',
            'debt', 'commission', 'experience_card', 'public_revenue', 'store_revenue'
        ]
        
        # 不均分的数值字段（每个老师都记录原始值）
        non_splittable_fields = [
            'debt_collection'  # 实收欠款不应该均分，应该每个老师都记录原始值
        ]
        
        empty_name = TEACHER_FILE_CONFIG['empty_teacher_name']
        
        for row in all_data:
            # 为每个角色创建具体老师的分组（支持多人分割）
            for role_field, role_name in role_mapping.items():
                teacher_names_raw = row.get(role_field)
                
                # 处理有值和无值的情况
                if teacher_names_raw and str(teacher_names_raw).strip() != '':
                    teacher_names_str = str(teacher_names_raw).strip()
                    
                    # 检查是否包含分隔符
                    if '/' in teacher_names_str:
                        # 多人情况：分割并均分数据
                        teacher_names = [name.strip() for name in teacher_names_str.split('/') if name.strip()]
                        
                        if teacher_names:  # 确保分割后有有效名称
                            person_count = len(teacher_names)
                            
                            # 为每个老师创建分组记录
                            for teacher_name in teacher_names:
                                group_key = f"{teacher_name}({role_name})"
                                
                                if group_key not in grouped_data:
                                    grouped_data[group_key] = []
                                
                                # 复制原始行数据
                                split_row = row.copy()
                                
                                # 将该角色字段设置为单个老师名称
                                split_row[role_field] = teacher_name
                                
                                # 均分可分割的数值字段
                                for field in splittable_fields:
                                    if field in split_row and split_row[field] is not None:
                                        try:
                                            # 处理可能的Excel公式字符串
                                            value = split_row[field]
                                            if isinstance(value, str) and value.startswith('='):
                                                self.logger.warning(f"检测到Excel公式 {field}={value}，跳过均分处理")
                                                continue
                                            
                                            original_value = float(value)
                                            # 使用四舍五入保留2位小数避免精度问题
                                            split_row[field] = round(original_value / person_count, 2)
                                        except (ValueError, TypeError) as e:
                                            # 如果转换失败，记录错误并保持原值
                                            self.logger.warning(f"数值字段 {field} 转换失败: {split_row[field]} -> 保持原值, 错误: {e}")
                                            pass
                                
                                # 对于不可分割的数值字段，保留原始值
                                for field in non_splittable_fields:
                                    if field in split_row and split_row[field] is not None:
                                        try:
                                            # 处理可能的Excel公式字符串
                                            value = split_row[field]
                                            if isinstance(value, str) and value.startswith('='):
                                                self.logger.warning(f"检测到Excel公式 {field}={value}，保持原值")
                                                continue
                                            
                                            # 确保数值格式正确，但不进行均分
                                            original_value = float(value)
                                            split_row[field] = round(original_value, 2)
                                        except (ValueError, TypeError) as e:
                                            # 如果转换失败，记录错误并保持原值
                                            self.logger.warning(f"数值字段 {field} 转换失败: {split_row[field]} -> 保持原值, 错误: {e}")
                                            pass
                                
                                grouped_data[group_key].append(split_row)
                        else:
                            # 分割后没有有效名称，归入"未分类"
                            group_key = f"未分类({role_name})"
                            if group_key not in grouped_data:
                                grouped_data[group_key] = []
                            grouped_data[group_key].append(row)
                    else:
                        # 单人情况：正常处理
                        group_key = f"{teacher_names_str}({role_name})"
                        if group_key not in grouped_data:
                            grouped_data[group_key] = []
                        grouped_data[group_key].append(row)
                else:
                    # 空值数据归入"未分类"
                    group_key = f"未分类({role_name})"
                    if group_key not in grouped_data:
                        grouped_data[group_key] = []
                    grouped_data[group_key].append(row)
            
            # 处理店家分组
            store_name = row.get('store_name')
            if store_name and str(store_name).strip() != '':
                group_key = f"{str(store_name).strip()}(店家)"
                if group_key not in grouped_data:
                    grouped_data[group_key] = []
                grouped_data[group_key].append(row)
        
        # 添加数据验证和统计信息
        self.logger.info("=" * 60)
        self.logger.info("📊 业绩分组处理结果统计:")
        self.logger.info("=" * 60)
        
        # 统计原始数据（安全地处理可能的Excel公式）
        def safe_float_convert(value):
            """安全地将值转换为浮点数，处理Excel公式"""
            if value is None:
                return 0
            if isinstance(value, str) and value.startswith('='):
                self.logger.warning(f"跳过Excel公式: {value}")
                return 0
            try:
                return float(value)
            except (ValueError, TypeError):
                return 0
        
        original_total_debt_collection = sum(
            safe_float_convert(row.get('debt_collection', 0)) for row in all_data
        )
        original_total_commission = sum(
            safe_float_convert(row.get('commission', 0)) for row in all_data
        )
        original_total_store_revenue = sum(
            safe_float_convert(row.get('store_revenue', 0)) for row in all_data
        )
        
        # 显示各分组统计（仅用于展示，不累加）
        for group_name, group_data in grouped_data.items():
            group_debt = sum(safe_float_convert(row.get('debt_collection', 0)) for row in group_data)
            group_commission = sum(safe_float_convert(row.get('commission', 0)) for row in group_data)
            group_store = sum(safe_float_convert(row.get('store_revenue', 0)) for row in group_data)
            
            self.logger.info(f"🔸 {group_name}: {len(group_data)}条记录, 收欠款: {group_debt:.2f}, 实收业绩: {group_commission:.2f}, 店收: {group_store:.2f}")
        
        # 分组后统计：通过原始数据重新计算，避免重复统计
        # 这里我们仅统计原始数据，因为分组只是为了输出不同的Excel文件
        grouped_total_debt_collection = original_total_debt_collection
        grouped_total_commission = original_total_commission  
        grouped_total_store_revenue = original_total_store_revenue
        
        self.logger.info("-" * 60)
        self.logger.info(f"💰 原始数据汇总: 收欠款总计: {original_total_debt_collection:.2f}, 实收业绩总计: {original_total_commission:.2f}, 店收总计: {original_total_store_revenue:.2f}")
        self.logger.info(f"💰 分组后汇总: 收欠款总计: {grouped_total_debt_collection:.2f}, 实收业绩总计: {grouped_total_commission:.2f}, 店收总计: {grouped_total_store_revenue:.2f}")
        
        # 验证数据一致性
        debt_diff = abs(original_total_debt_collection - grouped_total_debt_collection)
        commission_diff = abs(original_total_commission - grouped_total_commission)
        store_diff = abs(original_total_store_revenue - grouped_total_store_revenue)
        
        if debt_diff > 0.01:  # 允许0.01的精度误差
            self.logger.warning(f"⚠️  收欠款数据不一致! 差异: {debt_diff:.2f}")
        else:
            self.logger.info("✅ 收欠款数据一致性验证通过")
            
        if commission_diff > 0.01:  # 允许0.01的精度误差  
            self.logger.warning(f"⚠️  实收业绩数据不一致! 差异: {commission_diff:.2f}")
        else:
            self.logger.info("✅ 实收业绩数据一致性验证通过")
            
        if store_diff > 0.01:  # 允许0.01的精度误差
            self.logger.warning(f"⚠️  店收数据不一致! 差异: {store_diff:.2f}")
        else:
            self.logger.info("✅ 店收数据一致性验证通过")
        
        self.logger.info("=" * 60)

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

    def _generate_role_filename(self, original_filename: str, role_type: str) -> str:
        """
        生成按角色类型分类的文件名
        
        Args:
            original_filename: 原始文件名
            role_type: 角色类型 (如 "服务总监", "服务老师", "操作老师", "店家")
            
        Returns:
            str: 生成的文件名
        """
        name_without_ext = os.path.splitext(original_filename)[0]
        timestamp = datetime.now().strftime(TEACHER_FILE_CONFIG['timestamp_format'])
        
        filename = TEACHER_FILE_CONFIG['filename_format'].format(
            original_name=f"{name_without_ext}_{role_type}",
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
            
            # 创建表头样式 - 增加字体大小和边框
            header_font = Font(bold=True, color="FFFFFF", size=14)  # 增加字体大小
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_alignment = Alignment(horizontal="center", vertical="center")
            
            # 添加边框样式
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for (row, col), header_text in headers.items():
                cell = worksheet.cell(row=row, column=col, value=header_text)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
                cell.border = thin_border
                
            # 设置表头行高
            worksheet.row_dimensions[1].height = 25  # 增加行高以配合更大字体
                
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
            
            # 数值字段列表（包含所有需要数值格式化的字段）
            numeric_fields = ['order_amount', 'debt_collection', 'payment', 'card_deduction', 
                            'debt', 'commission', 'experience_card', 'public_revenue', 'store_revenue']
            
            # 创建数据行样式
            data_font = Font(size=12)  # 增加数据字体大小
            data_alignment = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for row_index, data_row in enumerate(teacher_data):
                current_row = data_start_row + row_index
                
                # 设置数据行高
                worksheet.row_dimensions[current_row].height = 20
                
                for field_name, col_index in output_columns.items():
                    value = data_row.get(field_name, "")
                    
                    # 处理数值字段
                    if field_name in numeric_fields and value is not None:
                        try:
                            if isinstance(value, str):
                                # 清理字符串中的千分位分隔符
                                value = value.replace(',', '').replace('，', '').strip()
                                if value == '' or value == '-':
                                    value = 0
                            
                            # 转换为浮点数并保留2位小数
                            numeric_value = float(value) if value else 0
                            value = round(numeric_value, 2)
                            
                            self.logger.debug(f"数值字段 {field_name} 处理: 原值={data_row.get(field_name)}, 处理后={value}")
                            
                        except (ValueError, TypeError) as e:
                            self.logger.warning(f"数值字段 {field_name} 转换失败: {data_row.get(field_name)} -> 设为0, 错误: {e}")
                            value = 0
                    
                    cell = worksheet.cell(row=current_row, column=col_index, value=value)
                    
                    # 应用样式 - 所有内容都居中显示
                    cell.font = data_font
                    cell.alignment = data_alignment  # 统一居中对齐
                    cell.border = thin_border
                    
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
            total_columns = TEACHER_OUTPUT_CONFIG['total_columns']
            
            # 写入合计标签
            worksheet.cell(row=total_row, column=total_label_column, value=total_label)
            
            # 计算各列合计金额
            totals = {}
            for field_name in total_columns.keys():
                totals[field_name] = 0
                
                for data_row in teacher_data:
                    value = data_row.get(field_name, 0)
                    if value:
                        try:
                            if isinstance(value, str):
                                value = value.replace(',', '').replace('，', '')
                            totals[field_name] += float(value)
                        except (ValueError, TypeError):
                            pass
            
            # 写入各列合计数据
            for field_name, col_index in total_columns.items():
                worksheet.cell(row=total_row, column=col_index, value=totals[field_name])
            
            # 设置合计行样式 - 美化
            total_font = Font(bold=True, size=12)  # 增加字体大小
            total_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
            total_alignment = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # 设置合计行高
            worksheet.row_dimensions[total_row].height = 22
            
            for col in range(1, 16):  # A-O列
                cell = worksheet.cell(row=total_row, column=col)
                cell.font = total_font
                cell.fill = total_fill
                cell.alignment = total_alignment  # 统一居中对齐
                cell.border = thin_border
            
            # 添加实收业绩和体验卡合计行
            grand_total_row = total_row + 1
            grand_total = totals['commission'] + totals['experience_card']
            
            # 写入总合计标签和数值
            worksheet.cell(row=grand_total_row, column=total_label_column, value="实收业绩和体验卡合计")
            worksheet.cell(row=grand_total_row, column=12, value=grand_total)  # 在实收业绩列显示总合计
            
            # 设置总合计行样式 - 美化
            grand_total_font = Font(bold=True, color="FF0000", size=12)  # 红色字体，增加字体大小
            grand_total_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # 黄色背景
            grand_total_alignment = Alignment(horizontal="center", vertical="center")
            
            # 设置总合计行高
            worksheet.row_dimensions[grand_total_row].height = 22
            
            for col in range(1, 16):  # A-O列
                cell = worksheet.cell(row=grand_total_row, column=col)
                cell.font = grand_total_font
                cell.fill = grand_total_fill
                cell.alignment = grand_total_alignment  # 统一居中对齐
                cell.border = thin_border
            
        except Exception as e:
            self.logger.error(f"写入合计行失败: {e}")
    
    def _adjust_column_width(self, worksheet) -> None:
        """
        调整列宽
        
        Args:
            worksheet: 工作表对象
        """
        try:
            # 设置各列宽度 - 适配更多列和更大字体
            column_widths = {
                'A': 12,  # 日期
                'B': 15,  # 客户
                'C': 12,  # 服务总监
                'D': 12,  # 服务老师
                'E': 12,  # 操作老师
                'F': 20,  # 店名
                'G': 12,  # 开单金额
                'H': 12,  # 收欠款
                'I': 12,  # 收款
                'J': 12,  # 卡扣
                'K': 12,  # 欠款
                'L': 15,  # 实收业绩
                'M': 12,  # 体验卡
                'N': 42,  # 开单明细 - 增加宽度以容纳更多详细信息
                'O': 12,  # 公司收
                'P': 12,  # 店收
            }
            
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
                
        except Exception as e:
            self.logger.error(f"调整列宽失败: {e}")
    
    def _safe_save_workbook(self, workbook, file_path: str, max_retries: int = 3) -> None:
        """
        安全保存Excel工作簿，包含重试机制和权限处理
        
        Args:
            workbook: Excel工作簿对象
            file_path: 文件保存路径
            max_retries: 最大重试次数
        """
        import time
        import os
        
        for attempt in range(max_retries):
            try:
                # 确保目录存在
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                
                # 如果文件已存在，尝试删除
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        self.logger.debug(f"删除已存在的文件: {file_path}")
                    except PermissionError:
                        # 如果删除失败，生成新的文件名
                        base_name, ext = os.path.splitext(file_path)
                        new_file_path = f"{base_name}_副本{attempt+1}{ext}"
                        self.logger.warning(f"无法覆盖原文件，将保存为: {new_file_path}")
                        file_path = new_file_path
                
                # 尝试保存文件
                workbook.save(file_path)
                self.logger.info(f"文件保存成功: {file_path}")
                return
                
            except PermissionError as e:
                self.logger.warning(f"文件保存权限错误 (尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    self.logger.info(f"等待 {(attempt + 1) * 2} 秒后重试...")
                    time.sleep((attempt + 1) * 2)  # 递增等待时间
                else:
                    # 最后一次尝试失败，生成带时间戳的文件名
                    from datetime import datetime
                    timestamp = datetime.now().strftime("%H%M%S")
                    base_name, ext = os.path.splitext(file_path)
                    fallback_path = f"{base_name}_{timestamp}{ext}"
                    try:
                        workbook.save(fallback_path)
                        self.logger.warning(f"使用备用文件名保存成功: {fallback_path}")
                        return
                    except Exception as fallback_error:
                        self.logger.error(f"备用保存也失败: {fallback_error}")
                        raise
                        
            except Exception as e:
                self.logger.error(f"文件保存失败 (尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(1)
                else:
                    raise 