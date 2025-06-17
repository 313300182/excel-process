#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 将Python应用打包为可执行文件
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


def clean_build_dirs():
    """清理之前的构建目录"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}")
            shutil.rmtree(dir_name)
    
    # 清理.spec文件
    for spec_file in Path('.').glob('*.spec'):
        print(f"删除文件: {spec_file}")
        spec_file.unlink()


def create_data_files():
    """创建数据文件列表"""
    data_files = []
    
    # 添加模板文件目录
    if os.path.exists('templates'):
        data_files.append(('templates', 'templates'))
    
    return data_files


def build_executable():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    # PyInstaller参数
    pyinstaller_args = [
        'pyinstaller',
        '--onefile',                    # 打包成单个文件
        '--windowed',                   # Windows下隐藏控制台窗口
        '--name=Excel批量处理工具',      # 可执行文件名
        '--distpath=dist',              # 输出目录
        '--workpath=build',             # 工作目录
        '--clean',                      # 清理临时文件
        '--noconfirm',                  # 不确认覆盖
        # '--add-data=templates;templates',  # 如果有模板文件，取消注释这行
        'main.py'                       # 主程序入口
    ]
    
    # 添加隐藏导入（如果需要）
    hidden_imports = [
        'openpyxl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext'
    ]
    
    for module in hidden_imports:
        pyinstaller_args.extend(['--hidden-import', module])
    
    # 执行PyInstaller命令
    try:
        result = subprocess.run(pyinstaller_args, check=True, capture_output=True, text=True)
        print("构建成功!")
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print("构建失败!")
        print("错误输出:", e.stderr)
        print("标准输出:", e.stdout)
        return False


def copy_additional_files():
    """复制额外的文件到dist目录"""
    dist_dir = Path('dist')
    if not dist_dir.exists():
        print("dist目录不存在")
        return
    
    # 复制模板文件
    templates_src = Path('templates')
    if templates_src.exists():
        templates_dst = dist_dir / 'templates'
        if templates_dst.exists():
            shutil.rmtree(templates_dst)
        shutil.copytree(templates_src, templates_dst)
        print(f"复制模板文件: {templates_src} -> {templates_dst}")
    
    # 复制README文件
    readme_files = ['README.md', 'readme.md', 'README.txt']
    for readme in readme_files:
        if os.path.exists(readme):
            shutil.copy2(readme, dist_dir)
            print(f"复制说明文件: {readme}")
            break


def create_installer():
    """创建安装程序（可选）"""
    print("如需创建安装程序，请使用NSIS或Inno Setup等工具")
    print("可执行文件位于: dist/Excel批量处理工具.exe")


def main():
    """主函数"""
    print("=" * 50)
    print("Excel批量处理工具 - 打包脚本")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("错误: 需要Python 3.7或更高版本")
        return False
    
    # 检查依赖
    try:
        import PyInstaller
        print(f"PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("错误: 未安装PyInstaller，请运行: pip install PyInstaller")
        return False
    
    try:
        # 1. 清理构建目录
        clean_build_dirs()
        
        # 2. 构建可执行文件
        if not build_executable():
            return False
        
        # 3. 复制额外文件
        copy_additional_files()
        
        # 4. 提示创建安装程序
        create_installer()
        
        print("\n" + "=" * 50)
        print("打包完成!")
        print("可执行文件位置: dist/Excel批量处理工具.exe")
        print("=" * 50)
        
        return True
        
    except Exception as e:
        print(f"打包过程中发生错误: {e}")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 