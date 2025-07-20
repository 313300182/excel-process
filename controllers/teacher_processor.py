# -*- coding: utf-8 -*-
"""
老师分组处理器控制器
协调老师分组Excel读取器和写入器，实现按老师分组的批量处理业务逻辑
"""

import os
import logging
import threading
from typing import List, Callable, Optional, Dict, Any
from concurrent.futures import ThreadPoolExecutor, as_completed

from models.teacher_excel_reader import TeacherExcelReader
from models.teacher_excel_writer import TeacherExcelWriter
from config.settings import SUPPORTED_EXTENSIONS
from config.teacher_splitter_settings import TEACHER_FILE_CONFIG


class TeacherProcessorController:
    """老师分组处理器控制器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.reader = TeacherExcelReader()
        self.writer = TeacherExcelWriter()
        self.is_processing = False
        self.should_stop = False
        
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
                        excel_files.append(file_path)
                        
            self.logger.info(f"在目录 {directory} 中找到 {len(excel_files)} 个Excel文件")
            return excel_files
            
        except Exception as e:
            self.logger.error(f"扫描目录失败: {e}")
            return excel_files
    
    def process_single_teacher_file(self, file_path: str, output_dir: str) -> Optional[str]:
        """
        处理单个Excel文件进行老师分组
        
        Args:
            file_path: 源文件路径
            output_dir: 输出目录
            
        Returns:
            str: 输出文件路径，失败返回None
        """
        try:
            if self.should_stop:
                return None
                
            self.logger.info(f"开始处理老师分组文件: {file_path}")
            
            # 读取数据
            data_list = self.reader.read_teacher_data(file_path)
            if not data_list:
                self.logger.warning(f"文件中没有有效数据: {file_path}")
                return None
                
            # 记录数据摘要
            teachers_summary = self.reader.get_teachers_summary(file_path)
            total_rows = teachers_summary['total_rows']
            teachers = teachers_summary['teachers']
            
            self.logger.info(f"文件 {os.path.basename(file_path)}: {total_rows} 行数据")
            for role, teacher_list in teachers.items():
                self.logger.info(f"  {role}: {len(teacher_list)} 人 - {', '.join(teacher_list[:5])}")
                
            # 生成输出文件
            original_filename = os.path.basename(file_path)
            output_path = self.writer.create_teacher_grouped_file(data_list, output_dir, original_filename, file_path)
            
            if output_path:
                self.logger.info(f"成功处理文件: {file_path} -> {output_path}")
            else:
                self.logger.error(f"处理文件失败: {file_path}")
                
            return output_path
            
        except Exception as e:
            self.logger.error(f"处理老师分组文件时发生错误 {file_path}: {e}")
            return None
    
    def process_teacher_batch(self, 
                             source_dir: str, 
                             output_dir: str,
                             progress_callback: Optional[Callable[[int, int, str], None]] = None,
                             complete_callback: Optional[Callable[[List[str], List[str]], None]] = None,
                             max_workers: int = 2) -> None:
        """
        批量处理Excel文件进行老师分组
        
        Args:
            source_dir: 源目录
            output_dir: 输出目录
            progress_callback: 进度回调函数 (current, total, current_file)
            complete_callback: 完成回调函数 (success_files, failed_files)
            max_workers: 最大线程数 (老师分组处理较复杂，使用较少线程)
        """
        def _process():
            try:
                self.is_processing = True
                self.should_stop = False
                
                # 扫描Excel文件
                excel_files = self.scan_excel_files(source_dir)
                if not excel_files:
                    self.logger.warning("没有找到需要处理的Excel文件")
                    if complete_callback:
                        complete_callback([], [])
                    return
                
                # 确保输出目录存在
                os.makedirs(output_dir, exist_ok=True)
                
                success_files = []
                failed_files = []
                total_files = len(excel_files)
                
                self.logger.info(f"开始批量老师分组处理 {total_files} 个文件，使用 {max_workers} 个线程")
                
                # 使用线程池并行处理
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # 提交所有任务
                    future_to_file = {
                        executor.submit(self.process_single_teacher_file, file_path, output_dir): file_path
                        for file_path in excel_files
                    }
                    
                    # 处理完成的任务
                    for i, future in enumerate(as_completed(future_to_file)):
                        if self.should_stop:
                            self.logger.info("用户请求停止处理")
                            break
                            
                        file_path = future_to_file[future]
                        
                        # 更新进度
                        if progress_callback:
                            progress_callback(i + 1, total_files, os.path.basename(file_path))
                        
                        try:
                            result = future.result()
                            if result:
                                success_files.append(result)
                            else:
                                failed_files.append(file_path)
                        except Exception as e:
                            self.logger.error(f"处理文件异常 {file_path}: {e}")
                            failed_files.append(file_path)
                
                # 处理完成回调
                if complete_callback:
                    complete_callback(success_files, failed_files)
                    
                self.logger.info(f"批量老师分组处理完成: 成功 {len(success_files)}, 失败 {len(failed_files)}")
                
            except Exception as e:
                self.logger.error(f"批量老师分组处理时发生错误: {e}")
                if complete_callback:
                    complete_callback([], excel_files if 'excel_files' in locals() else [])
            finally:
                self.is_processing = False
        
        # 在新线程中执行处理，避免阻塞UI
        processing_thread = threading.Thread(target=_process)
        processing_thread.daemon = True
        processing_thread.start()
    
    def stop_processing(self) -> None:
        """停止处理"""
        self.should_stop = True
        self.logger.info("已请求停止老师分组处理")
    
    def get_teacher_file_preview(self, file_path: str, rows: int = 5) -> List[Dict[str, Any]]:
        """
        获取老师分组文件预览
        
        Args:
            file_path: 文件路径
            rows: 预览行数
            
        Returns:
            List[Dict]: 预览数据
        """
        return self.reader.read_teacher_data(file_path)[:rows]
    
    def get_teacher_file_summary(self, file_path: str) -> Dict[str, Any]:
        """
        获取老师分组文件数据摘要
        
        Args:
            file_path: 文件路径
            
        Returns:
            Dict: 文件摘要信息
        """
        return self.reader.get_teachers_summary(file_path)
    
    def validate_teacher_source_directory(self, directory: str) -> Dict[str, Any]:
        """
        验证老师分组源目录
        
        Args:
            directory: 目录路径
            
        Returns:
            Dict: 验证结果
        """
        result = {
            'valid': False,
            'message': '',
            'file_count': 0,
            'files': [],
            'total_records': 0,
            'sample_summary': None
        }
        
        try:
            if not os.path.exists(directory):
                result['message'] = '目录不存在'
                return result
                
            if not os.path.isdir(directory):
                result['message'] = '路径不是目录'
                return result
                
            excel_files = self.scan_excel_files(directory)
            result['file_count'] = len(excel_files)
            result['files'] = [os.path.basename(f) for f in excel_files[:10]]
            
            if result['file_count'] == 0:
                result['message'] = '目录中没有Excel文件'
            else:
                result['valid'] = True
                result['message'] = f'找到 {result["file_count"]} 个Excel文件'
                
                # 获取第一个文件的摘要作为示例
                if excel_files:
                    try:
                        sample_summary = self.get_teacher_file_summary(excel_files[0])
                        result['sample_summary'] = sample_summary
                        result['total_records'] = sample_summary.get('total_rows', 0)
                        
                        if sample_summary.get('total_rows', 0) > 0:
                            result['message'] += f"\n示例文件: {os.path.basename(excel_files[0])} ({sample_summary['total_rows']} 行数据)"
                            
                            # 显示老师统计
                            teachers = sample_summary.get('teachers', {})
                            for role, teacher_list in teachers.items():
                                if teacher_list:
                                    result['message'] += f"\n{role}: {len(teacher_list)} 人"
                    except Exception as e:
                        self.logger.warning(f"获取示例文件摘要失败: {e}")
                
        except Exception as e:
            result['message'] = f'验证目录时发生错误: {e}'
            
        return result
    
    def prepare_teacher_output_directory(self, base_output_dir: str) -> str:
        """
        准备老师分组输出目录
        
        Args:
            base_output_dir: 基础输出目录
            
        Returns:
            str: 实际输出目录路径
        """
        try:
            # 创建带时间戳的子目录
            from datetime import datetime
            timestamp = datetime.now().strftime(TEACHER_FILE_CONFIG.get('timestamp_format', '%Y-%m-%d'))
            output_dir_name = f'老师分组输出-{timestamp}'
            output_dir = os.path.join(base_output_dir, output_dir_name)
                
            os.makedirs(output_dir, exist_ok=True)
            return output_dir
            
        except Exception as e:
            self.logger.error(f"准备输出目录失败: {e}")
            return base_output_dir
    
    def get_processing_status(self) -> Dict[str, Any]:
        """
        获取处理状态
        
        Returns:
            Dict: 状态信息
        """
        return {
            'is_processing': self.is_processing,
            'should_stop': self.should_stop,
        } 