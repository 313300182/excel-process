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


def setup_logging():
    """设置基础日志系统 - 仅控制台输出，UI日志由MainWindow管理"""
    try:
        # 简单的控制台日志配置（仅用于启动前的日志）
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        logging.info("Excel批量处理工具启动")
        
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