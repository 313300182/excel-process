#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel批量处理工具 - 简洁打包脚本
只保留必要的依赖，排除所有不需要的包
"""

import os
import subprocess

# --- 配置 ---
APP_NAME = "Excel数据处理器"
MAIN_SCRIPT = "main.py"
OUTPUT_DIR = "dist"  # 输出目录
BUILD_DIR = "build"   # Nuitka的构建目录

# Nuitka的基本命令
# --standalone: 创建一个独立的文件夹，包含所有依赖
# --onefile: (可选) 创建单个可执行文件，启动较慢
# --windows-disable-console: 在Windows上禁用控制台窗口
# --show-progress: 显示编译进度
# --follow-imports: 包含所有导入的模块
# --plugin-enable=tk-inter: 启用Tkinter插件
# --output-dir: 指定输出目录
# --output-filename: 指定输出文件名
# --windows-icon-from-ico: (可选) 指定程序图标
# --include-data-dir: (可选) 包含整个数据文件夹
# --include-data-file: (可选) 包含单个数据文件

# --- 构建命令 ---
command = [
    "python",
    "-m",
    "nuitka",
    "--standalone",
    "--remove-output", # <-- 添加此行
    # "--onefile",
    "--windows-disable-console",
    "--show-progress",
    "--follow-imports",
    "--plugin-enable=tk-inter",
    f"--output-dir={OUTPUT_DIR}",
    f"--output-filename={APP_NAME}.exe",
    # f"--windows-icon-from-ico=path/to/your/icon.ico", # 如果有图标，取消此行注释
    # "--include-data-file=path/to/your/template.xlsx=template.xlsx", # 如果有数据文件，取消此行注释
    MAIN_SCRIPT
]

def build():
    """执行Nuitka打包命令"""
    print("开始使用Nuitka打包...")
    print(f"命令: {' '.join(command)}")
    
    try:
        # 不再实时读取，而是等待完成
        result = subprocess.run(command, capture_output=True, text=True, encoding='gbk', check=False)

        # 打印标准输出和标准错误
        if result.stdout:
            print("--- Nuitka 输出 ---")
            print(result.stdout)
        if result.stderr:
            print("--- Nuitka 错误 ---")
            print(result.stderr)

        if result.returncode == 0:
            print("\n打包成功！")
            print(f"可执行文件位于: {os.path.abspath(OUTPUT_DIR)}")
        else:
            print(f"\n打包失败！返回码: {result.returncode}")

    except FileNotFoundError:
        print("\n错误: 'python' 或 'nuitka' 命令未找到。")
        print("请确保Python和Nuitka已正确安装并已添加到系统PATH中。")
    except Exception as e:
        print(f"\n打包过程中发生错误: {e}")

if __name__ == "__main__":
    build() 