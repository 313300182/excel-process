# -*- coding: utf-8 -*-
"""
工资Excel写入器
将工资数据填充到工资模板文件中
"""

import logging
import os
import json
from typing import Dict, Any, Optional, List
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from config.salary_settings import (
    SALARY_CONFIG, DEFAULT_SALARY_CONFIG, 
    JOB_SPECIFIC_CONFIG, SALARY_USER_CONFIG_FILE
)


class SalaryExcelWriter:
    """工资Excel写入器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.template_mapping = SALARY_CONFIG['template_mapping']
        self.templates = SALARY_CONFIG['templates']
        self.user_config = self._load_user_config()
        
    def _load_user_config(self) -> Dict[str, Any]:
        """
        加载用户配置
        
        Returns:
            Dict[str, Any]: 用户配置
        """
        try:
            if os.path.exists(SALARY_USER_CONFIG_FILE):
                with open(SALARY_USER_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    
                # 分离模板路径和配置数据
                if 'template_paths' in saved_config:
                    self.template_paths = saved_config.pop('template_paths')
                    self.logger.info(f"已加载模板路径: {list(self.template_paths.keys())}")
                    
                self.logger.info("已加载用户配置")
                return saved_config
        except Exception as e:
            self.logger.warning(f"加载用户配置失败: {str(e)}")
            
        return DEFAULT_SALARY_CONFIG.copy()
        
    def save_user_config(self, config: Dict[str, Any], template_paths: Dict[str, str] = None) -> bool:
        """
        保存用户配置（包括模板路径）
        
        Args:
            config: 配置数据
            template_paths: 模板文件路径
            
        Returns:
            bool: 是否保存成功
        """
        try:
            # 合并配置和模板路径
            save_config = config.copy()
            if template_paths:
                save_config['template_paths'] = template_paths
            elif hasattr(self, 'template_paths'):
                save_config['template_paths'] = self.template_paths
                
            with open(SALARY_USER_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(save_config, f, ensure_ascii=False, indent=2)
            self.user_config = config
            self.logger.info("用户配置已保存")
            return True
        except Exception as e:
            self.logger.error(f"保存用户配置失败: {str(e)}")
            return False
            
    def set_template_paths(self, template_paths: Dict[str, str]):
        """
        设置模板文件路径
        
        Args:
            template_paths: 职业类型到模板路径的映射
        """
        self.template_paths = template_paths
        self.logger.info(f"设置模板路径: {template_paths}")
        
    def process_salary_file(self, salary_data: Dict[str, Any], 
                           job_type: str, output_dir: str) -> str:
        """
        处理工资文件，生成工资条
        
        Args:
            salary_data: 工资数据
            job_type: 职业类型
            output_dir: 输出目录
            
        Returns:
            str: 输出文件路径
        """
        workbook = None
        try:
            # 获取模板文件路径
            template_path = self._get_template_path(job_type)
            if not template_path or not os.path.exists(template_path):
                raise Exception(f"找不到 {job_type} 的模板文件: {template_path}")
            
            # 加载模板文件
            self.logger.debug(f"加载模板文件: {template_path}")
            workbook = openpyxl.load_workbook(template_path)
            worksheet = workbook.active
            
            # 验证模板格式
            self._validate_template_structure(worksheet)
            
            # 填充基本信息
            self._fill_employee_info(worksheet, salary_data['employee_info'])
            
            # 计算和填充工资数据
            calculated_data = self._calculate_salary_data(
                salary_data, job_type)
            self._fill_salary_data(worksheet, calculated_data)
            
            # 生成输出文件
            output_path = self._generate_output_path(
                salary_data['employee_info'], job_type, output_dir)
            
            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # 保存文件
            workbook.save(output_path)
            self.logger.info(f"工资条已生成: {output_path}")
            
            return output_path
            
        except Exception as e:
            self.logger.error(f"处理工资文件失败: {str(e)}")
            raise Exception(f"处理工资文件失败: {str(e)}")
        finally:
            # 确保workbook被正确关闭，释放内存
            if workbook:
                try:
                    workbook.close()
                    workbook = None
                except:
                    pass
            
    def process_multiple_salary_to_single_file(self, employees_data: List[Dict[str, Any]], 
                                              output_path: str) -> str:
        """
        处理多个员工的工资数据到单个Excel文件，每个员工一个sheet
        
        Args:
            employees_data: 员工工资数据列表，每个元素包含 salary_data 和 job_type
            output_path: 输出文件路径
            
        Returns:
            str: 输出文件路径
        """
        output_workbook = None
        try:
            self.logger.info(f"开始处理 {len(employees_data)} 个员工到单个文件: {output_path}")
            
            # 创建新的工作簿
            output_workbook = openpyxl.Workbook()
            
            # 删除默认的工作表
            if 'Sheet' in [ws.title for ws in output_workbook.worksheets]:
                default_sheet = output_workbook['Sheet']
                output_workbook.remove(default_sheet)
            
            processed_count = 0
            
            for emp_data in employees_data:
                try:
                    salary_data = emp_data['salary_data']
                    job_type = emp_data['job_type']
                    employee_name = salary_data['employee_info'].get('name', '未知员工')
                    
                    self.logger.debug(f"处理员工: {employee_name}, 职业类型: {job_type}")
                    
                    # 获取模板文件路径
                    template_path = self._get_template_path(job_type)
                    if not template_path or not os.path.exists(template_path):
                        self.logger.error(f"找不到 {job_type} 的模板文件: {template_path}")
                        continue
                    
                    # 加载模板文件
                    template_workbook = openpyxl.load_workbook(template_path)
                    template_worksheet = template_workbook.active
                    
                    # 验证模板格式
                    self._validate_template_structure(template_worksheet)
                    
                    # 创建新的工作表，使用员工姓名作为sheet名
                    safe_name = self._sanitize_sheet_name(employee_name)
                    # 确保sheet名唯一
                    unique_name = self._get_unique_sheet_name(output_workbook, safe_name)
                    new_worksheet = output_workbook.create_sheet(title=unique_name)
                    
                    # 复制模板格式和内容
                    self._copy_worksheet_content(template_worksheet, new_worksheet)
                    
                    # 填充员工信息
                    self._fill_employee_info(new_worksheet, salary_data['employee_info'])
                    
                    # 计算和填充工资数据
                    calculated_data = self._calculate_salary_data(salary_data, job_type)
                    self._fill_salary_data(new_worksheet, calculated_data)
                    
                    processed_count += 1
                    self.logger.debug(f"员工 {employee_name} 处理完成")
                    
                    # 关闭模板工作簿
                    template_workbook.close()
                    
                except Exception as e:
                    employee_name = emp_data.get('salary_data', {}).get('employee_info', {}).get('name', '未知员工')
                    self.logger.error(f"处理员工 {employee_name} 失败: {str(e)}")
                    continue
            
            if processed_count == 0:
                raise Exception("没有成功处理任何员工数据")
            
            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # 保存文件
            output_workbook.save(output_path)
            self.logger.info(f"批量工资条已生成: {output_path}，包含 {processed_count} 个员工")
            
            return output_path
            
        except Exception as e:
            self.logger.error(f"批量处理工资文件失败: {str(e)}")
            raise Exception(f"批量处理工资文件失败: {str(e)}")
        finally:
            # 确保workbook被正确关闭，释放内存
            if output_workbook:
                try:
                    output_workbook.close()
                    output_workbook = None
                except:
                    pass
    
    def _sanitize_sheet_name(self, name: str) -> str:
        """
        清理工作表名称，移除不安全字符
        
        Args:
            name: 原名称
            
        Returns:
            str: 安全的工作表名称
        """
        # Excel工作表名称限制：不能包含 [ ] : * ? / \
        unsafe_chars = ['[', ']', ':', '*', '?', '/', '\\']
        safe_name = name
        
        for char in unsafe_chars:
            safe_name = safe_name.replace(char, '_')
        
        # 限制长度（Excel工作表名最大31个字符）
        if len(safe_name) > 31:
            safe_name = safe_name[:31]
            
        return safe_name.strip()
    
    def _get_unique_sheet_name(self, workbook, preferred_name: str) -> str:
        """
        获取唯一的工作表名称
        
        Args:
            workbook: 工作簿
            preferred_name: 首选名称
            
        Returns:
            str: 唯一的工作表名称
        """
        existing_names = [ws.title for ws in workbook.worksheets]
        
        if preferred_name not in existing_names:
            return preferred_name
        
        # 如果名称已存在，添加数字后缀
        counter = 1
        while True:
            candidate_name = f"{preferred_name}_{counter}"
            if len(candidate_name) > 31:
                # 如果加数字后超过31个字符，截短原名称
                base_length = 31 - len(f"_{counter}")
                candidate_name = f"{preferred_name[:base_length]}_{counter}"
            
            if candidate_name not in existing_names:
                return candidate_name
            counter += 1
    
    def _copy_worksheet_content(self, source_ws, target_ws):
        """
        复制工作表内容（包括格式、公式、样式等）
        
        Args:
            source_ws: 源工作表
            target_ws: 目标工作表
        """
        try:
            # 复制单元格数据和格式
            for row in source_ws.iter_rows():
                for cell in row:
                    target_cell = target_ws[cell.coordinate]
                    
                    # 复制值
                    if cell.value is not None:
                        target_cell.value = cell.value
                    
                    # 复制格式
                    if cell.has_style:
                        target_cell.font = cell.font.copy()
                        target_cell.border = cell.border.copy()
                        target_cell.fill = cell.fill.copy()
                        target_cell.number_format = cell.number_format
                        target_cell.protection = cell.protection.copy()
                        target_cell.alignment = cell.alignment.copy()
            
            # 复制合并单元格
            for merged_range in source_ws.merged_cells.ranges:
                target_ws.merge_cells(str(merged_range))
            
            # 复制行高
            for row_num in range(1, source_ws.max_row + 1):
                if row_num in source_ws.row_dimensions:
                    source_height = source_ws.row_dimensions[row_num].height
                    if source_height is not None:
                        target_ws.row_dimensions[row_num].height = source_height
                        
            # 复制列宽
            for col_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                if col_letter in source_ws.column_dimensions:
                    source_width = source_ws.column_dimensions[col_letter].width
                    if source_width is not None:
                        target_ws.column_dimensions[col_letter].width = source_width
                        self.logger.debug(f"复制列 {col_letter} 宽度: {source_width}")
                        
            # 复制更多列（如果需要）
            for col_num in range(1, source_ws.max_column + 1):
                col_letter = get_column_letter(col_num)
                if col_letter in source_ws.column_dimensions:
                    source_width = source_ws.column_dimensions[col_letter].width
                    if source_width is not None and col_letter not in target_ws.column_dimensions:
                        target_ws.column_dimensions[col_letter].width = source_width
                        self.logger.debug(f"复制列 {col_letter} 宽度: {source_width}")
                        
            self.logger.debug("工作表内容复制完成")
            
        except Exception as e:
            self.logger.warning(f"复制工作表内容时出现警告: {str(e)}")
            # 即使复制格式失败，也不影响数据填充
            
    def _get_template_path(self, job_type: str) -> Optional[str]:
        """
        获取模板文件路径
        
        Args:
            job_type: 职业类型
            
        Returns:
            Optional[str]: 模板文件路径
        """
        if hasattr(self, 'template_paths') and job_type in self.template_paths:
            return self.template_paths[job_type]
        return None
        
    def _fill_employee_info(self, worksheet, employee_info: Dict[str, Any]):
        """
        填充员工基本信息
        
        Args:
            worksheet: 工作表
            employee_info: 员工信息
        """
        try:
            # 填充姓名
            if 'name' in employee_info and employee_info['name']:
                name_cell = self.template_mapping['employee_name']
                self.logger.debug(f"写入姓名到单元格 {name_cell}: {employee_info['name']}")
                worksheet[name_cell] = str(employee_info['name'])
                
            # 填充月份
            if 'month' in employee_info and employee_info['month']:
                month_cell = self.template_mapping['month']
                self.logger.debug(f"写入月份到单元格 {month_cell}: {employee_info['month']}")
                worksheet[month_cell] = str(employee_info['month'])
                
            self.logger.debug("员工基本信息填充完成")
            
        except Exception as e:
            self.logger.error(f"填充员工信息失败: {str(e)}")
            raise Exception(f"填充员工信息失败: {str(e)}")
            
    def _calculate_salary_data(self, salary_data: Dict[str, Any], 
                              job_type: str) -> Dict[str, float]:
        """
        计算工资数据（基于业绩数据和手工费数据）
        
        Args:
            salary_data: 包含员工信息、业绩数据和操作数据的字典
            job_type: 职业类型
            
        Returns:
            Dict[str, float]: 计算后的工资数据
        """
        calculated = {}
        
        try:
            employee_name = salary_data['employee_info'].get('name', '')
            performance_data = salary_data.get('performance_data', {})
            operation_data = salary_data.get('operation_data', {})
            
            # 获取配置
            base_config = self.user_config.get('base_salary', {})
            floating_config = self.user_config.get('floating_salary', {})
            commission_config = self.user_config.get('commission_rates', {})
            manual_config = self.user_config.get('manual_fees', {})
            other_config = self.user_config.get('other_config', {})
            job_config = JOB_SPECIFIC_CONFIG.get(job_type, {})
            
            # 计算基本底薪
            base_salary = base_config.get('special_rates', {}).get(
                employee_name, base_config.get('default', 0))
            calculated['base_salary'] = base_salary * job_config.get('base_multiplier', 1.0)
            
            # 计算浮动底薪
            floating_salary = floating_config.get('special_rates', {}).get(
                employee_name, floating_config.get('default', 0))
            calculated['floating_salary'] = floating_salary
            
            # 计算提成（基于业绩数据）
            performance_value = performance_data.get('total_performance_value', 0)
            
            # 服务提成：基于总业绩值
            service_rate = commission_config.get('service_rate', 0) / 100
            calculated['service_commission'] = performance_value * service_rate * (
                1 + job_config.get('commission_bonus', 0))
            
            # 操作提成：基于总业绩值的一定比例
            operation_rate = commission_config.get('operation_rate', 0) / 100
            calculated['operation_commission'] = performance_value * operation_rate * (
                1 + job_config.get('commission_bonus', 0))
            
            # 计算手工费（基于操作表数据）
            body_quantity = operation_data.get('body_count', 0)
            body_rate = manual_config.get('body_rate', 0)
            calculated['body_manual_fee'] = body_quantity * body_rate
            
            face_quantity = operation_data.get('face_count', 0)
            face_rate = manual_config.get('face_rate', 0)
            calculated['face_manual_fee'] = face_quantity * face_rate
            
            # 保存原始数量和单价数据（用于填充到模板）
            calculated['body_count'] = body_quantity
            calculated['face_count'] = face_quantity
            calculated['body_rate_config'] = body_rate
            calculated['face_rate_config'] = face_rate
            
            # 培训补贴
            calculated['training_allowance'] = other_config.get('training_allowance', 0)
            
            # 特殊补贴（职业特有）
            calculated['special_allowance'] = job_config.get('special_allowance', 0)
            
            # 计算应发合计
            salary_total = (
                calculated['base_salary'] +
                calculated['floating_salary'] +
                calculated['service_commission'] +
                calculated['operation_commission'] +
                calculated['training_allowance'] +
                calculated['body_manual_fee'] +
                calculated['face_manual_fee'] +
                calculated['special_allowance']
            )
            calculated['total_salary'] = salary_total
            
            # 计算扣减项目（暂时设为0，可以根据需要扩展）
            calculated['late_deduction'] = 0  # 考勤扣款
            calculated['absent_deduction'] = 0  # 迟到扣款
            
            # 社保
            social_rate = other_config.get('social_security_rate', 8.0) / 100
            calculated['social_security'] = salary_total * social_rate
            
            # 个人所得税（简化计算）
            tax_rate = other_config.get('personal_tax_rate', 3.0) / 100
            calculated['personal_tax'] = max(0, (salary_total - 5000) * tax_rate)
            
            # 扣减小计
            deduction_total = (
                calculated['late_deduction'] +
                calculated['absent_deduction'] +
                calculated['social_security'] +
                calculated['personal_tax']
            )
            calculated['total_deduction'] = deduction_total
            
            # 实发工资
            calculated['net_salary'] = salary_total - deduction_total
            
            self.logger.info(f"员工 {employee_name} 工资计算完成，应发: {salary_total:.2f}, 实发: {calculated['net_salary']:.2f}")
            
        except Exception as e:
            self.logger.error(f"计算工资数据失败: {str(e)}")
            
        return calculated
        
    def _extract_quantity_from_details(self, details: List[Dict[str, Any]], 
                                     keyword: str) -> float:
        """
        从明细中提取数量
        
        Args:
            details: 工资明细
            keyword: 关键词
            
        Returns:
            float: 数量
        """
        total_quantity = 0.0
        
        try:
            for detail in details:
                project = str(detail.get('project', ''))
                if keyword in project:
                    quantity = detail.get('quantity', 0) or 0
                    total_quantity += float(quantity)
                    
        except Exception as e:
            self.logger.warning(f"提取 {keyword} 数量时出错: {str(e)}")
            
        return total_quantity
        
    def _fill_salary_data(self, worksheet, calculated_data: Dict[str, float]):
        """
        填充工资数据到模板
        
        Args:
            worksheet: 工作表
            calculated_data: 计算后的工资数据
        """
        try:
            self.logger.debug(f"开始填充工资数据，共 {len(calculated_data)} 项")
            
            # 填充应发项目
            salary_mapping = self.template_mapping.get('salary_items', {})
            self.logger.debug(f"填充应发项目，映射: {salary_mapping}")
            
            for key, cell in salary_mapping.items():
                if key in calculated_data:
                    try:
                        value = calculated_data[key]
                        rounded_value = round(float(value), 2) if value != 0 else 0
                        self.logger.debug(f"写入应发项目 {key} 到单元格 {cell}: {rounded_value}")
                        worksheet[cell] = rounded_value
                    except Exception as e:
                        self.logger.error(f"写入应发项目 {key} 到 {cell} 失败: {str(e)}")
                        raise
            
            # 填充手工费数量和单价
            manual_fee_mapping = self.template_mapping.get('manual_fee_details', {})
            if manual_fee_mapping:
                self.logger.debug(f"填充手工费明细，映射: {manual_fee_mapping}")
                
                # 填充身体部位数量和单价
                if 'body_quantity' in manual_fee_mapping and 'body_count' in calculated_data:
                    body_quantity = calculated_data.get('body_count', 0)
                    cell = manual_fee_mapping['body_quantity']
                    worksheet[cell] = int(body_quantity) if body_quantity else 0
                    self.logger.debug(f"写入身体部位数量到 {cell}: {body_quantity}")
                
                if 'body_rate' in manual_fee_mapping and 'body_rate_config' in calculated_data:
                    body_rate = calculated_data.get('body_rate_config', 0)
                    cell = manual_fee_mapping['body_rate']
                    worksheet[cell] = float(body_rate) if body_rate else 0
                    self.logger.debug(f"写入身体部位单价到 {cell}: {body_rate}")
                
                # 填充面部数量和单价
                if 'face_quantity' in manual_fee_mapping and 'face_count' in calculated_data:
                    face_quantity = calculated_data.get('face_count', 0)
                    cell = manual_fee_mapping['face_quantity']
                    worksheet[cell] = int(face_quantity) if face_quantity else 0
                    self.logger.debug(f"写入面部数量到 {cell}: {face_quantity}")
                
                if 'face_rate' in manual_fee_mapping and 'face_rate_config' in calculated_data:
                    face_rate = calculated_data.get('face_rate_config', 0)
                    cell = manual_fee_mapping['face_rate']
                    worksheet[cell] = float(face_rate) if face_rate else 0
                    self.logger.debug(f"写入面部单价到 {cell}: {face_rate}")
                        
            # 填充扣减项目
            deduction_mapping = self.template_mapping.get('deduction_items', {})
            self.logger.debug(f"填充扣减项目，映射: {deduction_mapping}")
            
            for key, cell in deduction_mapping.items():
                if key in calculated_data:
                    try:
                        value = calculated_data[key]
                        rounded_value = round(float(value), 2) if value != 0 else 0
                        self.logger.debug(f"写入扣减项目 {key} 到单元格 {cell}: {rounded_value}")
                        worksheet[cell] = rounded_value
                    except Exception as e:
                        self.logger.error(f"写入扣减项目 {key} 到 {cell} 失败: {str(e)}")
                        raise
                        
            # 填充实发工资
            net_cell = self.template_mapping.get('net_salary')
            if net_cell:
                try:
                    net_salary = round(float(calculated_data.get('net_salary', 0)), 2)
                    self.logger.debug(f"写入实发工资到单元格 {net_cell}: {net_salary}")
                    worksheet[net_cell] = net_salary
                except Exception as e:
                    self.logger.error(f"写入实发工资到 {net_cell} 失败: {str(e)}")
                    raise
            
            self.logger.debug("工资数据填充完成")
            
        except Exception as e:
            self.logger.error(f"填充工资数据失败: {str(e)}")
            raise Exception(f"填充工资数据失败: {str(e)}")
            
    def _generate_output_path(self, employee_info: Dict[str, Any], 
                            job_type: str, output_dir: str) -> str:
        """
        生成输出文件路径
        
        Args:
            employee_info: 员工信息
            job_type: 职业类型
            output_dir: 输出目录
            
        Returns:
            str: 输出文件路径
        """
        name = employee_info.get('name', '未知员工')
        month = employee_info.get('month', '未知月份')
        
        # 清理文件名中的特殊字符
        safe_name = self._sanitize_filename(name)
        safe_month = self._sanitize_filename(month)
        safe_job_type = self._sanitize_filename(job_type)
        
        filename = f"{safe_name}_{safe_month}_{safe_job_type}_工资条.xlsx"
        return os.path.join(output_dir, filename)
        
    def get_user_config(self) -> Dict[str, Any]:
        """
        获取当前用户配置
        
        Returns:
            Dict[str, Any]: 用户配置
        """
        return self.user_config.copy()
        
    def validate_templates(self, template_paths: Dict[str, str]) -> Dict[str, bool]:
        """
        验证模板文件
        
        Args:
            template_paths: 模板文件路径
            
        Returns:
            Dict[str, bool]: 验证结果
        """
        results = {}
        
        for job_type, template_path in template_paths.items():
            try:
                if not os.path.exists(template_path):
                    results[job_type] = False
                    continue
                    
                # 尝试打开文件
                workbook = openpyxl.load_workbook(template_path)
                worksheet = workbook.active
                
                # 检查关键单元格是否存在
                required_cells = [
                    self.template_mapping['employee_name'],
                    self.template_mapping['month'],
                    self.template_mapping['net_salary']
                ]
                
                valid = True
                for cell in required_cells:
                    try:
                        _ = worksheet[cell]
                    except:
                        valid = False
                        break
                        
                results[job_type] = valid
                
            except Exception as e:
                self.logger.error(f"验证模板 {job_type} 失败: {str(e)}")
                results[job_type] = False
                
        return results 
        
    def _validate_template_structure(self, worksheet):
        """
        验证模板结构是否正确
        
        Args:
            worksheet: 工作表
        """
        try:
            self.logger.debug("开始验证模板结构")
            
            # 验证关键单元格是否存在
            critical_cells = [
                self.template_mapping.get('employee_name'),
                self.template_mapping.get('month'),
                self.template_mapping.get('net_salary')
            ]
            
            for cell in critical_cells:
                if cell:
                    try:
                        # 尝试访问单元格
                        _ = worksheet[cell]
                        self.logger.debug(f"验证单元格 {cell}: OK")
                    except Exception as e:
                        raise Exception(f"无法访问关键单元格 {cell}: {str(e)}")
            
            # 验证工资项目单元格
            salary_items = self.template_mapping.get('salary_items', {})
            for key, cell in salary_items.items():
                try:
                    _ = worksheet[cell]
                    self.logger.debug(f"验证工资项目单元格 {key}({cell}): OK")
                except Exception as e:
                    self.logger.warning(f"工资项目单元格 {key}({cell}) 访问异常: {str(e)}")
            
            self.logger.debug("模板结构验证完成")
            
        except Exception as e:
            self.logger.error(f"模板结构验证失败: {str(e)}")
            raise Exception(f"模板结构验证失败: {str(e)}") 