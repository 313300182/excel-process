# -*- coding: utf-8 -*-
"""
Excel处理器控制器
协调Excel读取器和写入器，实现批量处理业务逻辑
"""

import os
import logging
import threading
from typing import List, Callable, Optional, Dict, Any
from concurrent.futures import ThreadPoolExecutor, as_completed

from models.excel_reader import ExcelReader
from models.excel_writer import ExcelWriter
from config.settings import SUPPORTED_EXTENSIONS, OUTPUT_CONFIG


class ProcessorController:
    """Excel处理器控制器"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.reader = ExcelReader()
        self.writer = ExcelWriter()
        self.is_processing = False
        self.should_stop = False
    
    def set_template_path(self, template_path: str):
        """
        设置模板文件路径
        
        Args:
            template_path: 模板文件路径
        """
        self.writer.set_template_path(template_path)
        
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
    
    def process_single_file(self, file_path: str, output_dir: str) -> Optional[str]:
        """
        处理单个Excel文件
        
        Args:
            file_path: 源文件路径
            output_dir: 输出目录
            
        Returns:
            str: 输出文件路径，失败返回None
        """
        try:
            if self.should_stop:
                return None
                
            self.logger.info(f"开始处理文件: {file_path}")
            
            # 读取数据
            data_list = self.reader.read_data(file_path)
            if not data_list:
                self.logger.warning(f"文件中没有有效数据: {file_path}")
                return None
                
            # 记录数据摘要
            total_amount = 0
            for row in data_list:
                amount = row.get('amount')
                if amount is not None:
                    try:
                        if isinstance(amount, str):
                            amount = amount.replace(',', '').replace('，', '')
                        total_amount += float(amount)
                    except (ValueError, TypeError):
                        pass
            
            self.logger.info(f"文件 {os.path.basename(file_path)}: {len(data_list)} 行数据，合计金额: {total_amount:.2f}")
                
            # 生成输出文件
            original_filename = os.path.basename(file_path)
            output_path = self.writer.create_output_file(data_list, output_dir, original_filename)
            
            if output_path:
                self.logger.info(f"成功处理文件: {file_path} -> {output_path}")
            else:
                self.logger.error(f"处理文件失败: {file_path}")
                
            return output_path
            
        except Exception as e:
            self.logger.error(f"处理文件时发生错误 {file_path}: {e}")
            return None
    
    def process_batch(self, 
                     source_dir: str, 
                     output_dir: str,
                     progress_callback: Optional[Callable[[int, int, str], None]] = None,
                     complete_callback: Optional[Callable[[List[str], List[str]], None]] = None,
                     max_workers: int = 2) -> None:  # 减少线程数避免阻塞
        """
        批量处理Excel文件
        
        Args:
            source_dir: 源目录
            output_dir: 输出目录
            progress_callback: 进度回调函数 (current, total, current_file)
            complete_callback: 完成回调函数 (success_files, failed_files)
            max_workers: 最大线程数
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
                
                self.logger.info(f"开始批量处理 {total_files} 个文件，顺序处理模式")
                
                # 改为顺序处理，避免线程池阻塞
                for i, file_path in enumerate(excel_files):
                    if self.should_stop:
                        self.logger.info("用户请求停止处理")
                        break
                        
                    # 更新进度（在处理前）
                    if progress_callback:
                        try:
                            progress_callback(i + 1, total_files, os.path.basename(file_path))
                        except Exception as e:
                            self.logger.warning(f"进度回调失败: {e}")
                    
                    # 处理单个文件
                    try:
                        result = self.process_single_file(file_path, output_dir)
                        if result:
                            success_files.append(result)
                            self.logger.info(f"处理成功 ({i+1}/{total_files}): {os.path.basename(file_path)}")
                        else:
                            failed_files.append(file_path)
                            self.logger.warning(f"处理失败 ({i+1}/{total_files}): {os.path.basename(file_path)}")
                    except Exception as e:
                        self.logger.error(f"处理文件异常 {file_path}: {e}")
                        failed_files.append(file_path)
                    
                    # 添加小延迟，让UI有机会响应
                    import time
                    time.sleep(0.1)
                
                # 处理完成回调
                if complete_callback:
                    try:
                        complete_callback(success_files, failed_files)
                    except Exception as e:
                        self.logger.error(f"完成回调失败: {e}")
                    
                self.logger.info(f"批量处理完成: 成功 {len(success_files)}, 失败 {len(failed_files)}")
                
            except Exception as e:
                self.logger.error(f"批量处理时发生错误: {e}")
                if complete_callback:
                    try:
                        complete_callback([], excel_files if 'excel_files' in locals() else [])
                    except Exception as callback_e:
                        self.logger.error(f"错误回调失败: {callback_e}")
            finally:
                self.is_processing = False
        
        # 在新线程中执行处理，避免阻塞UI
        processing_thread = threading.Thread(target=_process, name="ExcelProcessor")
        processing_thread.daemon = True
        processing_thread.start()
    
    def stop_processing(self) -> None:
        """停止处理"""
        self.should_stop = True
        self.logger.info("已请求停止处理")
    
    def get_file_preview(self, file_path: str, rows: int = 5) -> List[Dict[str, Any]]:
        """
        获取文件预览
        
        Args:
            file_path: 文件路径
            rows: 预览行数
            
        Returns:
            List[Dict]: 预览数据
        """
        return self.reader.preview_data(file_path, rows)
    
    def get_file_summary(self, file_path: str) -> Dict[str, Any]:
        """
        获取文件数据摘要
        
        Args:
            file_path: 文件路径
            
        Returns:
            Dict: 文件摘要信息
        """
        return self.reader.get_data_summary(file_path)
    
    def validate_source_directory(self, directory: str) -> Dict[str, Any]:
        """
        验证源目录
        
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
            result['files'] = [os.path.basename(f) for f in excel_files[:10]]  # 只显示前10个
            
            if result['file_count'] == 0:
                result['message'] = '目录中没有Excel文件'
            else:
                result['valid'] = True
                result['message'] = f'找到 {result["file_count"]} 个Excel文件'
                
                # 获取第一个文件的摘要作为示例
                if excel_files:
                    try:
                        sample_summary = self.get_file_summary(excel_files[0])
                        result['sample_summary'] = sample_summary
                        result['total_records'] = sample_summary.get('total_rows', 0)
                        
                        if sample_summary.get('total_rows', 0) > 0:
                            result['message'] += f"\n示例文件: {os.path.basename(excel_files[0])} ({sample_summary['total_rows']} 行数据)"
                            if sample_summary.get('total_amount'):
                                result['message'] += f", 合计金额: {sample_summary['total_amount']:.2f}"
                    except Exception as e:
                        self.logger.warning(f"获取示例文件摘要失败: {e}")
                
        except Exception as e:
            result['message'] = f'验证目录时发生错误: {e}'
            
        return result
    
    def prepare_output_directory(self, base_output_dir: str) -> str:
        """
        准备输出目录
        
        Args:
            base_output_dir: 基础输出目录
            
        Returns:
            str: 实际输出目录路径
        """
        try:
            if OUTPUT_CONFIG['create_output_dir']:
                # 创建带时间戳的子目录，使用新的命名格式
                from datetime import datetime
                timestamp = datetime.now().strftime(OUTPUT_CONFIG.get('dir_timestamp_format', '%Y-%m-%d'))
                dir_format = OUTPUT_CONFIG.get('default_output_dir', '输出模板-{timestamp}')
                output_dir_name = dir_format.format(timestamp=timestamp)
                output_dir = os.path.join(base_output_dir, output_dir_name)
            else:
                output_dir = base_output_dir
                
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
            'template_valid': self.writer.validate_template()
        } 