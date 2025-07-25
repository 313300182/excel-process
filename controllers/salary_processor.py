# -*- coding: utf-8 -*-
"""
工资处理器控制器
协调工资Excel读取器和写入器，实现工资批量处理业务逻辑
"""

import os
import logging
import threading
from typing import List, Callable, Optional, Dict, Any
from concurrent.futures import ThreadPoolExecutor, as_completed

from models.salary_excel_reader import SalaryExcelReader
from models.salary_excel_writer import SalaryExcelWriter
from models.operation_table_reader import OperationTableReader
from config.settings import SUPPORTED_EXTENSIONS
from config.salary_settings import SALARY_CONFIG


class SalaryProcessorController:
    """工资处理器控制器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.reader = SalaryExcelReader()
        self.writer = SalaryExcelWriter()
        self.operation_reader = OperationTableReader()
        self.is_processing = False
        self.should_stop = False
        self.operation_table_path = None
        self.operation_data = {}
        
        # 从writer中获取已加载的模板路径
        self.template_paths = getattr(self.writer, 'template_paths', {})
        
        # 如果writer中有模板路径，同步到当前控制器
        if self.template_paths:
            self.logger.info(f"从配置文件加载了模板路径: {list(self.template_paths.keys())}")
        else:
            self.logger.info("未找到保存的模板路径，需要用户重新设置")
        
    def set_template_paths(self, template_paths: Dict[str, str]):
        """
        设置模板文件路径
        
        Args:
            template_paths: 职业类型到模板路径的映射
        """
        self.template_paths = template_paths
        self.writer.set_template_paths(template_paths)
        # 同步更新writer中的模板路径
        self.writer.template_paths = template_paths
        self.logger.info(f"已设置模板路径: {list(template_paths.keys())}")
        
    def set_operation_table_path(self, operation_table_path: str):
        """
        设置手工费操作表路径
        
        Args:
            operation_table_path: 操作表文件路径
        """
        self.operation_table_path = operation_table_path
        try:
            # 读取操作表数据
            self.operation_data = self.operation_reader.read_operation_data(operation_table_path)
            self.logger.info(f"已加载手工费操作表: {len(self.operation_data)} 个员工")
            
            # 输出手工费数据概览
            if self.operation_data:
                self.logger.info("💰 手工费数据概览:")
                for name, data in self.operation_data.items():
                    self.logger.info(f"  - {name}: 部位{data['body_count']}, 面部{data['face_count']}")
                    
        except Exception as e:
            self.logger.error(f"加载手工费操作表失败: {str(e)}")
            self.operation_data = {}
        
    def scan_excel_files(self, directory: str) -> List[str]:
        """
        扫描目录中的Excel文件
        
        Args:
            directory: 目录路径
            
        Returns:
            List[str]: Excel文件路径列表
        """
        excel_files = []
        
        try:
            if not os.path.exists(directory):
                self.logger.error(f"目录不存在: {directory}")
                return excel_files
                
            for root, dirs, files in os.walk(directory):
                for file in files:
                    _, ext = os.path.splitext(file)
                    if ext.lower() in SUPPORTED_EXTENSIONS:
                        file_path = os.path.join(root, file)
                        # 验证文件结构
                        if self.reader.validate_file_structure(file_path):
                            excel_files.append(file_path)
                        else:
                            self.logger.warning(f"文件结构不符合要求，跳过: {file}")
                            
        except Exception as e:
            self.logger.error(f"扫描Excel文件时出错: {str(e)}")
            
        return excel_files
        
    def get_processing_info(self, source_dir: str) -> Dict[str, Any]:
        """
        获取处理信息
        
        Args:
            source_dir: 源文件目录
            
        Returns:
            Dict[str, Any]: 处理信息
        """
        info = {
            'total_files': 0,
            'valid_files': 0,
            'invalid_files': [],
            'employee_info': [],
            'ready_to_process': False
        }
        
        try:
            excel_files = self.scan_excel_files(source_dir)
            info['total_files'] = len(excel_files)
            info['valid_files'] = len(excel_files)
            
            # 获取员工信息预览
            total_employees = 0
            for file_path in excel_files[:3]:  # 只预览前3个文件
                try:
                    salary_data = self.reader.read_salary_data(file_path)
                    employees = salary_data.get('employees', [])
                    total_employees += len(employees)
                    
                    # 添加前几个员工的信息到预览
                    for employee in employees[:5]:
                        employee_preview = {
                            'name': employee['employee_info']['name'],
                            'month': employee['employee_info']['month'],
                            'file_name': os.path.basename(file_path),
                            'performance_value': employee['performance_data']['total_performance_value']
                        }
                        info['employee_info'].append(employee_preview)
                        
                except Exception as e:
                    self.logger.warning(f"预览文件失败 {file_path}: {str(e)}")
            
            info['total_employees'] = total_employees
                    
            # 检查是否准备就绪
            templates_ready = len(self.template_paths) > 0
            files_ready = info['valid_files'] > 0
            operation_table_ready = self.operation_table_path is not None
            
            info['ready_to_process'] = templates_ready and files_ready
            
            if not templates_ready:
                info['error_message'] = "请先设置工资模板文件"
            elif not files_ready:
                info['error_message'] = "没有找到有效的工资数据文件"
            elif not operation_table_ready:
                info['error_message'] = "请先上传手工费操作表"
                info['ready_to_process'] = False
                
        except Exception as e:
            self.logger.error(f"获取处理信息时出错: {str(e)}")
            info['error_message'] = f"获取信息失败: {str(e)}"
            
        return info
        
    def determine_job_type_from_filename(self, filename: str) -> str:
        """
        从文件名确定职业类型
        
        Args:
            filename: 文件名
            
        Returns:
            str: 职业类型，如果无法识别返回None
        """
        job_types = SALARY_CONFIG['job_types']
        
        # 从文件名中提取职业类型
        for job_type in job_types:
            if job_type in filename:
                return job_type
                
        return None
        
    def determine_job_type(self, performance_value: float) -> str:
        """
        根据业绩数据确定职业类型（备用方法）
        
        Args:
            performance_value: 业绩价值（实收业绩+体验卡合计）*10000
            
        Returns:
            str: 职业类型
        """
        job_types = SALARY_CONFIG['job_types']
        
        try:
            # 根据业绩值判断职业类型
            if performance_value > 100000:  # 10万以上 -> 服务总监
                return job_types[0] if len(job_types) > 0 else job_types[0]  # 服务总监
            elif performance_value > 50000:  # 5万-10万 -> 服务老师
                return job_types[1] if len(job_types) > 1 else job_types[0]  # 服务老师
            else:  # 5万以下 -> 操作老师
                return job_types[2] if len(job_types) > 2 else job_types[0]  # 操作老师
                
        except Exception as e:
            self.logger.warning(f"确定职业类型时出错: {str(e)}")
            
        return job_types[0]  # 默认返回第一个职业类型
        
    def process_files(self, source_dir: str, output_dir: str, 
                     progress_callback: Optional[Callable] = None,
                     log_callback: Optional[Callable] = None,
                     max_workers: int = 1) -> Dict[str, Any]:
        """
        批量处理工资文件
        
        Args:
            source_dir: 源文件目录
            output_dir: 输出目录
            progress_callback: 进度回调函数
            log_callback: 日志回调函数
            max_workers: 最大工作线程数
            
        Returns:
            Dict[str, Any]: 处理结果
        """
        result = {
            'success': False,
            'processed_files': 0,
            'failed_files': 0,
            'total_files': 0,
            'output_files': [],
            'errors': []
        }
        
        def log_message(message: str, level: str = "INFO"):
            self.logger.info(message)
            if log_callback:
                log_callback(f"[{level}] {message}")
        
        log_message("🔧 初始化工资处理器状态")
        self.is_processing = True
        self.should_stop = False
        
        try:
            log_message("📁 开始扫描源文件目录")
            # 扫描Excel文件
            excel_files = self.scan_excel_files(source_dir)
            result['total_files'] = len(excel_files)
            log_message(f"📊 扫描完成，找到 {len(excel_files)} 个Excel文件")
            
            if not excel_files:
                raise Exception("没有找到有效的工资数据文件")
                
            log_message("🔍 检查模板路径配置")
            if not self.template_paths:
                raise Exception("请先设置工资模板文件")
            log_message(f"✅ 模板路径已配置: {list(self.template_paths.keys())}")
                
            # 确保输出目录存在
            log_message(f"📂 准备输出目录: {output_dir}")
            os.makedirs(output_dir, exist_ok=True)
                    
            log_message(f"🚀 开始处理 {len(excel_files)} 个工资文件")
            
            # 处理文件 - 使用单线程避免内存和并发问题
            log_message(f"📋 准备处理模式: max_workers={max_workers}")
            if max_workers <= 1:
                log_message("🔄 使用单线程处理模式")
                # 单线程处理 - 更稳定
                for i, file_path in enumerate(excel_files):
                    log_message(f"🔍 开始处理第 {i+1}/{len(excel_files)} 个文件")
                    if self.should_stop:
                        log_message("⏹️ 检测到停止信号，中断处理")
                        break
                    
                    try:
                        log_message(f"📄 正在处理文件: {os.path.basename(file_path)}", "INFO")
                        output_files = self._process_single_file(file_path, output_dir, log_message)
                        result['output_files'].extend(output_files)
                        result['processed_files'] += 1
                        log_message(f"✅ 文件处理完成: {os.path.basename(file_path)}")
                        
                        # 更新进度
                        if progress_callback:
                            progress_percent = int((i + 1) * 100 / len(excel_files))
                            log_message(f"📈 更新进度: {progress_percent}%")
                            progress_callback(progress_percent)
                            
                    except Exception as e:
                        result['failed_files'] += 1
                        error_msg = f"{os.path.basename(file_path)}: {str(e)}"
                        result['errors'].append(error_msg)
                        log_message(f"❌ 处理文件失败: {error_msg}", "ERROR")
                        
                        # 记录详细错误
                        import traceback
                        self.logger.error(f"文件处理异常详情:\n{traceback.format_exc()}")
                        
                log_message("🔄 单线程处理循环完成")
            else:
                # 多线程处理
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # 提交任务
                    future_to_file = {}
                    for file_path in excel_files:
                        if self.should_stop:
                            break
                        future = executor.submit(
                            self._process_single_file, 
                            file_path, output_dir, log_message
                        )
                        future_to_file[future] = file_path
                    
                    # 处理结果
                    completed = 0
                    for future in as_completed(future_to_file):
                        if self.should_stop:
                            break
                            
                        file_path = future_to_file[future]
                        completed += 1
                        
                        try:
                            output_files = future.result()
                            if output_files:
                                result['output_files'].extend(output_files)
                                result['processed_files'] += 1
                                log_message(f"✓ 已处理: {os.path.basename(file_path)} ({len(output_files)}个工资条)")
                            else:
                                result['failed_files'] += 1
                                
                        except Exception as e:
                            result['failed_files'] += 1
                            error_msg = f"处理文件失败 {os.path.basename(file_path)}: {str(e)}"
                            result['errors'].append(error_msg)
                            log_message(error_msg, "ERROR")
                            
                        # 更新进度
                        if progress_callback:
                            progress = (completed / len(excel_files)) * 100
                            progress_callback(progress)
                        
            if self.should_stop:
                log_message("处理已被用户停止", "WARNING")
            else:
                result['success'] = True
                log_message(f"处理完成！成功: {result['processed_files']}, 失败: {result['failed_files']}")
                
        except Exception as e:
            error_msg = f"批量处理失败: {str(e)}"
            result['errors'].append(error_msg)
            self.logger.error(error_msg)
            if log_callback:
                log_callback(f"[ERROR] {error_msg}")
        
        finally:
            self.is_processing = False
            # 强制垃圾回收以释放内存
            import gc
            gc.collect()
            
            # 记录处理结果
            if log_callback:
                try:
                    import psutil
                    memory_info = psutil.virtual_memory()
                    log_callback(f"内存使用率: {memory_info.percent:.1f}%", "INFO")
                except:
                    pass
            
        return result
        
    def process_files_to_single_excel(self, source_dir: str, output_dir: str, output_filename: str = None,
                                     progress_callback: Optional[Callable] = None,
                                     log_callback: Optional[Callable] = None) -> Dict[str, Any]:
        """
        批量处理工资文件到单个Excel文件，每个员工一个sheet
        
        Args:
            source_dir: 源文件目录
            output_dir: 输出目录
            output_filename: 输出文件名（可选）
            progress_callback: 进度回调函数
            log_callback: 日志回调函数
            
        Returns:
            Dict[str, Any]: 处理结果
        """
        result = {
            'success': False,
            'processed_employees': 0,
            'total_employees': 0,
            'processed_files': 0,
            'failed_files': 0,
            'output_file': '',
            'errors': []
        }
        
        def log_message(message: str, level: str = "INFO"):
            self.logger.info(message)
            if log_callback:
                log_callback(f"[{level}] {message}")
        
        log_message("🔧 初始化工资处理器状态")
        self.is_processing = True
        self.should_stop = False
        
        try:
            log_message("📁 开始扫描源文件目录")
            # 扫描Excel文件
            excel_files = self.scan_excel_files(source_dir)
            result['processed_files'] = len(excel_files)
            log_message(f"📊 扫描完成，找到 {len(excel_files)} 个Excel文件")
            
            if not excel_files:
                raise Exception("没有找到有效的工资数据文件")
                
            log_message("🔍 检查模板路径配置")
            if not self.template_paths:
                raise Exception("请先设置工资模板文件")
            log_message(f"✅ 模板路径已配置: {list(self.template_paths.keys())}")
                
            # 确保输出目录存在
            log_message(f"📂 准备输出目录: {output_dir}")
            os.makedirs(output_dir, exist_ok=True)
            
            # 生成输出文件名
            if not output_filename:
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"工资单汇总_{timestamp}.xlsx"
            
            if not output_filename.endswith('.xlsx'):
                output_filename += '.xlsx'
                
            output_path = os.path.join(output_dir, output_filename)
            result['output_file'] = output_path
            
            log_message(f"📄 输出文件: {output_filename}")
            log_message(f"🚀 开始处理 {len(excel_files)} 个工资文件")
            
            # 收集所有员工数据
            all_employees_data = []
            total_files = len(excel_files)
            
            for i, file_path in enumerate(excel_files):
                if self.should_stop:
                    log_message("⏹️ 检测到停止信号，中断处理")
                    break
                
                try:
                    log_message(f"🔍 处理第 {i+1}/{total_files} 个文件: {os.path.basename(file_path)}")
                    
                    # 读取工资数据
                    salary_data = self.reader.read_salary_data(file_path)
                    employees = salary_data.get('employees', [])
                    
                    log_message(f"👥 文件 {os.path.basename(file_path)} 包含 {len(employees)} 个员工")
                    
                    # 优先从文件名识别职业类型
                    filename = os.path.basename(file_path)
                    job_type_from_file = self.determine_job_type_from_filename(filename)
                    
                    if job_type_from_file:
                        log_message(f"📋 从文件名识别职业类型: {job_type_from_file}")
                        # 检查是否有对应的模板
                        if job_type_from_file not in self.template_paths:
                            log_message(f"❌ 职业类型 {job_type_from_file} 没有对应的模板文件", "ERROR")
                            result['failed_files'] += 1
                            continue
                    
                    # 处理文件中的每个员工
                    for employee in employees:
                        if self.should_stop:
                            break
                            
                        employee_name = employee['employee_info'].get('name', '未知员工')
                        performance_value = employee['performance_data'].get('total_performance_value', 0)
                        
                        try:
                            # 确定职业类型：优先使用文件名识别，否则根据业绩判断
                            if job_type_from_file:
                                job_type = job_type_from_file
                                log_message(f"👤 员工 {employee_name}: 使用文件名职业类型 - {job_type}")
                            else:
                                job_type = self.determine_job_type(performance_value)
                                log_message(f"👤 员工 {employee_name}: 根据业绩({performance_value})判断职业类型 - {job_type}")
                            
                            # 检查是否有对应的模板
                            if job_type not in self.template_paths:
                                log_message(f"❌ 员工 {employee_name}: 没有找到职业类型 {job_type} 的模板文件", "ERROR")
                                continue
                            
                            # 从操作表获取完整数据（包括个人所得税）
                            default_operation_data = {
                                'body_count': 0, 
                                'face_count': 0,
                                'rest_days': 0,
                                'actual_absent_days': 0,
                                'late_count': 0,
                                'training_days': 0,
                                'work_days': 0,
                                'personal_tax_amount': 0
                            }
                            operation_data = self.operation_data.get(employee_name, default_operation_data)
                            if employee_name in self.operation_data:
                                log_message(f"📋 员工 {employee_name}: 从操作表获取 - 部位{operation_data.get('body_count', 0)}, 面部{operation_data.get('face_count', 0)}, 个税{operation_data.get('personal_tax_amount', 0):.2f}元")
                            else:
                                log_message(f"⚠️  员工 {employee_name}: 操作表中未找到，使用默认值 - 部位0, 面部0, 个税0元", "WARNING")
                            
                            # 构建员工数据
                            employee_data = {
                                'salary_data': {
                                    'employee_info': employee['employee_info'],
                                    'performance_data': employee['performance_data'],
                                    'operation_data': operation_data,
                                    'file_path': file_path
                                },
                                'job_type': job_type
                            }
                            
                            all_employees_data.append(employee_data)
                            result['total_employees'] += 1
                            
                            log_message(f"✅ 员工 {employee_name} 数据收集完成")
                            
                        except Exception as e:
                            log_message(f"❌ 处理员工 {employee_name} 失败: {str(e)}", "ERROR")
                            continue
                    
                    # 更新进度（收集阶段）
                    if progress_callback:
                        progress_percent = int((i + 1) * 50 / total_files)  # 收集阶段占50%
                        progress_callback(progress_percent)
                
                except Exception as e:
                    result['failed_files'] += 1
                    error_msg = f"处理文件失败 {os.path.basename(file_path)}: {str(e)}"
                    result['errors'].append(error_msg)
                    log_message(f"❌ {error_msg}", "ERROR")
                    
                    # 记录详细错误
                    import traceback
                    self.logger.error(f"文件处理异常详情:\n{traceback.format_exc()}")
            
            if not all_employees_data:
                raise Exception("没有收集到任何有效的员工数据")
            
            log_message(f"📊 数据收集完成，共收集 {len(all_employees_data)} 个员工数据")
            
            # 批量处理到单个Excel文件
            log_message(f"📝 开始生成包含所有员工的工资单Excel文件")
            if progress_callback:
                progress_callback(60)  # 开始生成阶段
                
            output_file = self.writer.process_multiple_salary_to_single_file(
                all_employees_data, output_path)
            
            if output_file:
                result['processed_employees'] = len(all_employees_data)
                result['success'] = True
                log_message(f"🎉 工资单汇总文件生成完成!")
                log_message(f"📄 输出文件: {output_file}")
                log_message(f"👥 包含员工: {result['processed_employees']} 人")
                
                if progress_callback:
                    progress_callback(100)
            else:
                raise Exception("生成工资单汇总文件失败")
                
        except Exception as e:
            error_msg = f"批量处理失败: {str(e)}"
            result['errors'].append(error_msg)
            self.logger.error(error_msg)
            if log_callback:
                log_callback(f"[ERROR] {error_msg}")
        
        finally:
            self.is_processing = False
            # 强制垃圾回收以释放内存
            import gc
            gc.collect()
            
            # 记录处理结果
            if log_callback:
                try:
                    import psutil
                    memory_info = psutil.virtual_memory()
                    log_callback(f"内存使用率: {memory_info.percent:.1f}%", "INFO")
                except:
                    pass
            
        return result
        
    def _process_single_file(self, file_path: str, output_dir: str,
                           log_callback: Callable) -> List[str]:
        """
        处理单个工资文件（可能包含多个员工）
        
        Args:
            file_path: 文件路径
            output_dir: 输出目录
            log_callback: 日志回调
            
        Returns:
            List[str]: 输出文件路径列表
        """
        output_files = []
        
        salary_data = None
        employees = []
        try:
            log_callback(f"🔍 开始读取源文件: {os.path.basename(file_path)}", "INFO")
            
            # 检查文件是否存在和可访问
            if not os.path.exists(file_path):
                raise Exception(f"文件不存在: {file_path}")
            if not os.path.isfile(file_path):
                raise Exception(f"路径不是文件: {file_path}")
                
            log_callback(f"📂 文件验证通过，开始读取: {os.path.getsize(file_path)} 字节", "DEBUG")
            
            # 读取工资数据
            log_callback(f"📖 调用读取器读取数据...", "DEBUG")
            salary_data = self.reader.read_salary_data(file_path)
            log_callback(f"📊 读取器返回数据: {type(salary_data)}", "DEBUG")
            
            employees = salary_data.get('employees', [])
            log_callback(f"👥 提取员工数据: {len(employees)} 个员工", "DEBUG")
            
            log_callback(f"📊 文件 {os.path.basename(file_path)} 包含 {len(employees)} 个员工", "INFO")
            
            # 立即提取需要的数据，然后释放原始数据
            processed_employees = []
            for emp in employees:
                # 只提取必要的数据
                processed_emp = {
                    'name': emp['employee_info'].get('name', ''),
                    'month': emp['employee_info'].get('month', ''),
                    'performance_value': emp['performance_data'].get('total_performance_value', 0)
                }
                processed_employees.append(processed_emp)
            
            # 立即清理原始数据，释放内存
            salary_data = None
            employees = None
            
            log_callback(f"✅ 源文件数据已提取并释放，开始处理 {len(processed_employees)} 个员工", "INFO")
            
            # 优先从文件名识别职业类型
            filename = os.path.basename(file_path)
            job_type_from_file = self.determine_job_type_from_filename(filename)
            
            if job_type_from_file:
                log_callback(f"从文件名识别职业类型: {job_type_from_file}", "INFO")
                # 检查是否有对应的模板
                if job_type_from_file not in self.template_paths:
                    log_callback(f"职业类型 {job_type_from_file} 没有对应的模板文件", "ERROR")
                    return output_files
            
            for i, emp_data in enumerate(processed_employees):
                if self.should_stop:
                    log_callback("处理已被停止", "WARNING")
                    break
                    
                employee_name = emp_data['name'] or "未知员工"
                try:
                    performance_value = emp_data['performance_value']
                    
                    log_callback(f"正在处理员工 {i+1}/{len(processed_employees)}: {employee_name}", "INFO")
                    
                    # 确定职业类型：优先使用文件名识别，否则根据业绩判断
                    if job_type_from_file:
                        job_type = job_type_from_file
                        log_callback(f"员工 {employee_name}: 使用文件名职业类型 - {job_type}", "INFO")
                    else:
                        job_type = self.determine_job_type(performance_value)
                        log_callback(f"员工 {employee_name}: 根据业绩({performance_value})判断职业类型 - {job_type}", "INFO")
                    
                    # 检查是否有对应的模板
                    if job_type not in self.template_paths:
                        log_callback(f"员工 {employee_name}: 没有找到职业类型 {job_type} 的模板文件", "ERROR")
                        continue
                    
                    # 从操作表获取完整数据（包括个人所得税）
                    default_operation_data = {
                        'body_count': 0, 
                        'face_count': 0,
                        'rest_days': 0,
                        'actual_absent_days': 0,
                        'late_count': 0,
                        'training_days': 0,
                        'work_days': 0,
                        'personal_tax_amount': 0
                    }
                    operation_data = self.operation_data.get(employee_name, default_operation_data)
                    if employee_name in self.operation_data:
                        log_callback(f"📋 员工 {employee_name}: 从操作表获取 - 部位数量={operation_data.get('body_count', 0)}, 面部数量={operation_data.get('face_count', 0)}, 个税={operation_data.get('personal_tax_amount', 0):.2f}元", "INFO")
                    else:
                        log_callback(f"⚠️  员工 {employee_name}: 操作表中未找到，使用默认值 - 部位数量=0, 面部数量=0, 个税=0元", "WARNING")
                    
                    # 重新构建最小化的数据结构
                    combined_data = {
                        'employee_info': {
                            'name': employee_name,
                            'month': emp_data['month'] or '未知月份'
                        },
                        'performance_data': {
                            'total_performance_value': performance_value
                        },
                        'operation_data': operation_data,
                        'file_path': file_path
                    }
                    
                    # 生成工资条
                    try:
                        output_path = self.writer.process_salary_file(
                            combined_data, job_type, output_dir)
                            
                        if output_path:
                            output_files.append(output_path)
                            log_callback(f"✓ 员工 {employee_name} 工资条生成完成: {os.path.basename(output_path)}", "INFO")
                    except Exception as e:
                        log_callback(f"生成员工 {employee_name} 工资条失败: {str(e)}", "ERROR")
                        continue
                    
                    # 清理临时数据
                    combined_data = None
                    
                except Exception as e:
                    log_callback(f"处理员工 {employee_name} 失败: {str(e)}", "ERROR")
                    continue
                    
        except Exception as e:
            log_callback(f"处理文件失败 {os.path.basename(file_path)}: {str(e)}", "ERROR")
            import traceback
            self.logger.error(f"文件处理异常详情:\n{traceback.format_exc()}")
            raise Exception(f"处理文件失败: {str(e)}")
        finally:
            # 确保资源被清理
            try:
                salary_data = None
                employees = None
                if 'processed_employees' in locals():
                    processed_employees = None
                # 强制垃圾回收
                import gc
                gc.collect()
                log_callback(f"🧹 文件 {os.path.basename(file_path)} 资源清理完成", "DEBUG")
            except:
                pass
            
        return output_files
            
    def stop_processing(self):
        """停止处理"""
        self.should_stop = True
        self.logger.info("收到停止处理请求")
        
    def validate_templates(self) -> Dict[str, bool]:
        """
        验证所有模板文件
        
        Returns:
            Dict[str, bool]: 验证结果
        """
        return self.writer.validate_templates(self.template_paths)
        
    def get_user_config(self) -> Dict[str, Any]:
        """
        获取用户配置
        
        Returns:
            Dict[str, Any]: 用户配置
        """
        return self.writer.get_user_config()
        
    def save_user_config(self, config: Dict[str, Any]) -> bool:
        """
        保存用户配置
        
        Args:
            config: 配置数据
            
        Returns:
            bool: 是否保存成功
        """
        return self.writer.save_user_config(config, self.template_paths)
        
    def get_job_types(self) -> List[str]:
        """
        获取支持的职业类型列表
        
        Returns:
            List[str]: 职业类型列表
        """
        return SALARY_CONFIG['job_types']
        
    def get_template_requirements(self) -> Dict[str, str]:
        """
        获取模板文件要求说明
        
        Returns:
            Dict[str, str]: 职业类型到说明的映射
        """
        requirements = {}
        
        for job_type in SALARY_CONFIG['job_types']:
            requirements[job_type] = (
                f"{job_type}工资模板需要包含以下位置：\n"
                f"- 员工姓名: {self.writer.template_mapping['employee_name']}\n"
                f"- 月份: {self.writer.template_mapping['month']}\n"
                f"- 应发合计: {self.writer.template_mapping['salary_items']['total_salary']}\n"
                f"- 实发工资: {self.writer.template_mapping['net_salary']}\n"
                f"请确保模板文件格式正确"
            )
            
        return requirements
        
    def preview_processing(self, source_dir: str, max_files: int = 3) -> List[Dict[str, Any]]:
        """
        预览处理结果
        
        Args:
            source_dir: 源文件目录
            max_files: 最大预览文件数
            
        Returns:
            List[Dict[str, Any]]: 预览数据
        """
        preview_data = []
        
        try:
            excel_files = self.scan_excel_files(source_dir)
            
            for file_path in excel_files[:max_files]:
                try:
                    # 读取工资数据
                    salary_data = self.reader.read_salary_data(file_path)
                    
                    # 确定职业类型
                    job_type = self.determine_job_type(salary_data)
                    
                    preview_item = {
                        'file_name': os.path.basename(file_path),
                        'employee_name': salary_data['employee_info'].get('name', '未知'),
                        'month': salary_data['employee_info'].get('month', '未知'),
                        'job_type': job_type,
                        'total_records': salary_data['statistics']['total_records'],
                        'has_template': job_type in self.template_paths,
                        'template_path': self.template_paths.get(job_type, '')
                    }
                    
                    preview_data.append(preview_item)
                    
                except Exception as e:
                    preview_item = {
                        'file_name': os.path.basename(file_path),
                        'error': str(e)
                    }
                    preview_data.append(preview_item)
                    
        except Exception as e:
            self.logger.error(f"预览处理时出错: {str(e)}")
            
        return preview_data 