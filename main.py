#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel批量处理工具 - 主程序入口
"""

import sys
import os
import logging

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from views.main_window import MainWindow
from config.settings import LOG_CONFIG, get_log_path


def setup_logging():
    """设置日志系统"""
    try:
        # 配置根日志器
        logging.basicConfig(
            level=getattr(logging, LOG_CONFIG['level']),
            format=LOG_CONFIG['format'],
            filename=get_log_path(),
            filemode='a',
            encoding='utf-8'
        )
        
        # 创建控制台处理器（在UI启动前的日志输出）
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter(LOG_CONFIG['format']))
        
        # 获取根日志器并添加控制台处理器
        root_logger = logging.getLogger()
        root_logger.addHandler(console_handler)
        
        logging.info("日志系统初始化完成")
        
    except Exception as e:
        print(f"日志系统初始化失败: {e}")


def main():
    """主函数"""
    try:
        # 设置日志
        setup_logging()
        
        logging.info("Excel批量处理工具启动")
        logging.info(f"工作目录: {os.getcwd()}")
        logging.info(f"项目根目录: {project_root}")
        
        # 创建并运行主窗口
        app = MainWindow()
        app.run()
        
        logging.info("Excel批量处理工具退出")
        
    except Exception as e:
        logging.error(f"程序运行时发生错误: {e}", exc_info=True)
        import tkinter.messagebox as messagebox
        messagebox.showerror("严重错误", f"程序运行时发生严重错误:\n{e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 