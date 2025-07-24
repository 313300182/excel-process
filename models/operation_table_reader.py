# -*- coding: utf-8 -*-
"""
手工费操作表读取器
从手工费汇总表中提取员工的部位数量和面部数量
"""

import logging
from typing import Dict, Any, Optional
import openpyxl
from openpyxl.utils import column_index_from_string


class OperationTableReader:
    """手工费操作表读取器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def read_operation_data(self, file_path: str) -> Dict[str, Any]:
        """
        读取手工费操作表数据
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            Dict[str, Any]: 员工手工费数据 {员工姓名: {部位数量: x, 面部数量: y}}
        """
        workbook = None
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            operation_data = {}
            
            # 查找表头位置
            header_info = self._find_headers(worksheet)
            if not header_info:
                raise Exception("未找到有效的表头信息")
                
            start_row = header_info['header_row'] + 1
            name_col = header_info['name_col']
            body_count_col = header_info.get('body_count_col')
            face_count_col = header_info.get('face_count_col')
            
            self.logger.info(f"表头行: {header_info['header_row']}, 姓名列: {name_col}, 部位数量列: {body_count_col}, 面部数量列: {face_count_col}")
            
            self.logger.info("=" * 50)
            self.logger.info("开始读取手工费操作表数据:")
            self.logger.info("=" * 50)
            
            # 读取数据行
            for row_idx in range(start_row, worksheet.max_row + 1):
                try:
                    # 获取员工姓名
                    name_cell = worksheet.cell(row=row_idx, column=name_col)
                    if not name_cell.value:
                        continue
                        
                    employee_name = str(name_cell.value).strip()
                    if not employee_name or employee_name in ['合计', '小计', '总计']:
                        continue
                    
                    # 获取部位数量
                    body_count = 0
                    if body_count_col:
                        body_cell = worksheet.cell(row=row_idx, column=body_count_col)
                        body_count = self._convert_to_number(body_cell.value)
                    
                    # 获取面部数量
                    face_count = 0
                    if face_count_col:
                        face_cell = worksheet.cell(row=row_idx, column=face_count_col)
                        face_count = self._convert_to_number(face_cell.value)
                    
                    # 存储数据
                    operation_data[employee_name] = {
                        'body_count': body_count,
                        'face_count': face_count
                    }
                    
                    # 打印员工数据（更清楚的格式）
                    self.logger.info(f"📋 员工: {employee_name:8} | 部位数量: {body_count:3} | 面部数量: {face_count:3}")
                    
                except Exception as e:
                    self.logger.warning(f"读取第 {row_idx} 行数据失败: {str(e)}")
                    continue
            
            # 输出汇总信息
            self.logger.info("=" * 50)
            self.logger.info(f"✅ 手工费操作表读取完成! 共读取 {len(operation_data)} 个员工数据")
            
            if operation_data:
                total_body = sum(data['body_count'] for data in operation_data.values())
                total_face = sum(data['face_count'] for data in operation_data.values())
                self.logger.info(f"📊 汇总: 总部位数量={total_body}, 总面部数量={total_face}")
            
            self.logger.info("=" * 50)
            
            return operation_data
            
        except Exception as e:
            self.logger.error(f"读取手工费操作表失败 {file_path}: {str(e)}")
            raise Exception(f"读取手工费操作表失败: {str(e)}")
        finally:
            # 确保workbook被正确关闭，释放内存
            if workbook:
                try:
                    workbook.close()
                    workbook = None
                except:
                    pass
            
    def _find_headers(self, worksheet) -> Optional[Dict[str, Any]]:
        """
        查找表头信息
        
        Args:
            worksheet: Excel工作表
            
        Returns:
            Optional[Dict[str, Any]]: 表头信息
        """
        try:
            header_info = {
                'header_row': None,
                'name_col': None,
                'body_count_col': None,
                'face_count_col': None
            }
            
            # 在前10行中查找表头
            for row_idx in range(1, min(11, worksheet.max_row + 1)):
                row_data = []
                for col_idx in range(1, min(21, worksheet.max_column + 1)):  # 检查前20列
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    row_data.append(cell.value)
                
                # 检查是否为表头行
                name_found = False
                body_found = False
                face_found = False
                
                for col_idx, cell_value in enumerate(row_data, 1):
                    if cell_value and isinstance(cell_value, str):
                        cell_str = str(cell_value).strip()
                        
                        # 查找姓名列
                        if any(keyword in cell_str for keyword in ['姓名', '员工', '操作老师', '老师']):
                            header_info['name_col'] = col_idx
                            name_found = True
                            
                        # 查找部位数量列（更精确匹配，避免匹配到手工费列）
                        elif any(keyword in cell_str for keyword in ['部位数量']) and '手工' not in cell_str and '元' not in cell_str:
                            header_info['body_count_col'] = col_idx
                            body_found = True
                            
                        # 查找面部数量列
                        elif any(keyword in cell_str for keyword in ['面部数量', '面部']):
                            header_info['face_count_col'] = col_idx
                            face_found = True
                
                # 如果找到了姓名列，认为找到了表头行
                if name_found:
                    header_info['header_row'] = row_idx
                    self.logger.info(f"✅ 找到表头行: 第{row_idx}行")
                    self.logger.info(f"  - 姓名列: 第{header_info['name_col']}列")
                    self.logger.info(f"  - 部位数量列: 第{header_info['body_count_col']}列" if header_info['body_count_col'] else "  - ⚠️ 未找到部位数量列")
                    self.logger.info(f"  - 面部数量列: 第{header_info['face_count_col']}列" if header_info['face_count_col'] else "  - ⚠️ 未找到面部数量列")
                    break
            
            # 验证是否找到了必要的列
            if header_info['header_row'] and header_info['name_col']:
                return header_info
            else:
                self.logger.error("未找到有效的表头信息")
                return None
                
        except Exception as e:
            self.logger.error(f"查找表头失败: {str(e)}")
            return None
            
    def _convert_to_number(self, value) -> float:
        """
        将值转换为数字
        
        Args:
            value: 要转换的值
            
        Returns:
            float: 转换后的数字，如果转换失败返回0
        """
        if value is None:
            return 0.0
            
        try:
            if isinstance(value, (int, float)):
                return float(value)
            elif isinstance(value, str):
                # 移除逗号和空格
                cleaned = str(value).replace(',', '').replace(' ', '').strip()
                if cleaned:
                    return float(cleaned)
            return 0.0
        except (ValueError, TypeError):
            return 0.0
            
    def validate_operation_table(self, file_path: str) -> bool:
        """
        验证手工费操作表文件是否有效
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 是否有效
        """
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            # 检查是否能找到表头
            header_info = self._find_headers(worksheet)
            
            workbook.close()
            return header_info is not None
            
        except Exception as e:
            self.logger.error(f"验证手工费操作表失败 {file_path}: {str(e)}")
            return False
            
    def get_employee_operation_data(self, file_path: str, employee_name: str) -> Dict[str, float]:
        """
        获取特定员工的手工费数据
        
        Args:
            file_path: 文件路径
            employee_name: 员工姓名
            
        Returns:
            Dict[str, float]: 员工手工费数据
        """
        try:
            operation_data = self.read_operation_data(file_path)
            return operation_data.get(employee_name, {'body_count': 0, 'face_count': 0})
        except Exception as e:
            self.logger.error(f"获取员工 {employee_name} 手工费数据失败: {str(e)}")
            return {'body_count': 0, 'face_count': 0} 