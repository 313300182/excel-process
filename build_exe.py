#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel批量处理工具 - 统一打包脚本
支持多种打包模式，自动解决依赖问题
"""

import os
import shutil
import subprocess
import sys
import argparse
from pathlib import Path


def clean_build_dirs():
    """清理之前的构建目录"""
    dirs_to_clean = ['build', 'dist', '__pycache__', 'hooks']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}")
            shutil.rmtree(dir_name)
    
    # 清理.spec文件
    for spec_file in Path('.').glob('*.spec'):
        print(f"删除文件: {spec_file}")
        spec_file.unlink()


def create_hook_file():
    """创建PyInstaller hook文件来解决openpyxl问题"""
    hook_dir = Path('hooks')
    hook_dir.mkdir(exist_ok=True)
    
    hook_content = '''# Hook for openpyxl
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = collect_all('openpyxl')

# 额外添加可能缺失的模块
hiddenimports += [
    'openpyxl.cell._writer',
    'openpyxl.worksheet.write_only',
    'openpyxl.worksheet.worksheet',
    'openpyxl.workbook.workbook',
    'openpyxl.styles.fonts',
    'openpyxl.styles.borders', 
    'openpyxl.styles.alignment',
    'openpyxl.styles.fills',
    'openpyxl.styles.colors',
    'openpyxl.styles.numbers',
    'openpyxl.utils.cell',
    'openpyxl.utils.exceptions',
    'openpyxl.reader.excel',
    'openpyxl.writer.excel',
    'openpyxl.writer.write_only'
]
'''
    
    hook_file = hook_dir / 'hook-openpyxl.py'
    with open(hook_file, 'w', encoding='utf-8') as f:
        f.write(hook_content)
    
    print(f"创建hook文件: {hook_file}")
    return str(hook_dir)


def build_executable(mode='auto', use_hooks=True, onedir=False, console=False):
    """
    构建可执行文件
    
    Args:
        mode: 打包模式 ('auto', 'simple', 'advanced')
        use_hooks: 是否使用自定义hooks
        onedir: 是否使用目录模式
        console: 是否显示控制台
    """
    print(f"开始构建可执行文件... 模式: {mode}")
    
    # 创建hook文件（如果需要）
    hook_path = None
    if use_hooks:
        hook_path = create_hook_file()
    
    # 基础PyInstaller参数
    pyinstaller_args = [
        'pyinstaller',
        '--onedir' if onedir else '--onefile',  # 打包模式
        '--console' if console else '--windowed',  # 控制台模式
        '--name=Excel批量处理工具',              # 可执行文件名
        '--distpath=dist',                      # 输出目录
        '--workpath=build',                     # 工作目录
        '--clean',                              # 清理临时文件
        '--noconfirm',                          # 不确认覆盖
        '--add-data=templates;templates',       # 添加模板文件
    ]
    
    # 排除不必要的大型库以减少体积（只排除最明显的大型库）
    if mode == 'simple':
        exclude_modules = [
            'numpy', 'pandas', 'scipy', 'matplotlib', 
            'jupyter', 'notebook', 'IPython',
            'sklearn', 'tensorflow', 'torch',
            'plotly', 'bokeh', 'sympy'
        ]
        
        for module in exclude_modules:
            pyinstaller_args.extend(['--exclude-module', module])
    
    # 添加hooks（如果使用）
    if use_hooks and hook_path:
        pyinstaller_args.extend([
            f'--additional-hooks-dir={hook_path}',
            '--collect-all=openpyxl'
        ])
    
    # 根据模式添加不同的参数
    if mode == 'advanced':
        # 高级模式：最全面的依赖收集
        hidden_imports = [
            'openpyxl', 'openpyxl.workbook', 'openpyxl.worksheet',
            'openpyxl.cell', 'openpyxl.cell._writer',
            'openpyxl.worksheet.write_only', 'openpyxl.worksheet.worksheet',
            'openpyxl.workbook.workbook', 'openpyxl.styles',
            'openpyxl.styles.fonts', 'openpyxl.styles.borders',
            'openpyxl.styles.alignment', 'openpyxl.styles.fills',
            'openpyxl.styles.colors', 'openpyxl.styles.numbers',
            'openpyxl.utils', 'openpyxl.utils.cell', 'openpyxl.utils.exceptions',
            'openpyxl.reader', 'openpyxl.reader.excel',
            'openpyxl.writer', 'openpyxl.writer.excel', 'openpyxl.writer.write_only',
            'xlrd', 'xlrd.sheet', 'xlrd.book',
            'tkinter', 'tkinter.ttk', 'tkinter.filedialog',
            'tkinter.messagebox', 'tkinter.scrolledtext',
            'logging', 'logging.handlers', 'subprocess', 'pathlib'
        ]
    elif mode == 'simple':
        # 简单模式：基础依赖
        hidden_imports = [
            'openpyxl', 'openpyxl.cell._writer',
            'xlrd', 'tkinter', 'tkinter.ttk',
            'tkinter.filedialog', 'tkinter.messagebox'
        ]
    else:  # auto模式
        # 自动模式：平衡的依赖配置
        hidden_imports = [
            'openpyxl', 'openpyxl.cell._writer',
            'openpyxl.worksheet.write_only', 'openpyxl.workbook.workbook',
            'xlrd', 'tkinter', 'tkinter.ttk',
            'tkinter.filedialog', 'tkinter.messagebox',
            'tkinter.scrolledtext', 'logging.handlers'
        ]
    
    # 添加隐藏导入
    for module in hidden_imports:
        pyinstaller_args.extend(['--hidden-import', module])
    
    # 添加主程序入口
    pyinstaller_args.append('main.py')
    
    # 执行PyInstaller命令
    try:
        print("执行命令:", ' '.join(pyinstaller_args))
        result = subprocess.run(pyinstaller_args, check=True, capture_output=True, text=True)
        print("构建成功!")
        if result.stdout:
            # 只显示重要信息
            output_lines = result.stdout.split('\n')
            important_lines = [line for line in output_lines if 
                             'WARNING' in line or 'ERROR' in line or 
                             'Building' in line or 'completed' in line]
            if important_lines:
                print("构建信息:")
                for line in important_lines[-10:]:  # 最后10行重要信息
                    print(f"  {line}")
        return True
    except subprocess.CalledProcessError as e:
        print("构建失败!")
        print("错误输出:", e.stderr)
        if e.stdout:
            print("标准输出:", e.stdout[-1000:])  # 最后1000字符
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


def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description='Excel批量处理工具 - 统一打包脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
打包模式说明:
  auto     - 自动模式（推荐）：平衡的依赖配置
  simple   - 简单模式：最少依赖，快速打包
  advanced - 高级模式：最全依赖，解决所有兼容性问题

使用示例:
  python build_exe.py                    # 默认自动模式
  python build_exe.py --mode advanced    # 高级模式
  python build_exe.py --onedir          # 目录模式
  python build_exe.py --console         # 显示控制台
  python build_exe.py --no-hooks        # 不使用自定义hooks
        """
    )
    
    parser.add_argument('--mode', choices=['auto', 'simple', 'advanced'], 
                       default='auto', help='打包模式 (默认: auto)')
    parser.add_argument('--onedir', action='store_true', 
                       help='使用目录模式而非单文件模式')
    parser.add_argument('--console', action='store_true', 
                       help='显示控制台窗口（用于调试）')
    parser.add_argument('--no-hooks', action='store_true', 
                       help='不使用自定义hooks')
    parser.add_argument('--clean-only', action='store_true', 
                       help='只清理构建目录，不进行打包')
    
    return parser.parse_args()


def main():
    """主函数"""
    args = parse_arguments()
    
    print("=" * 60)
    print("Excel批量处理工具 - 统一打包脚本")
    print("=" * 60)
    
    # 只清理模式
    if args.clean_only:
        print("清理构建目录...")
        clean_build_dirs()
        print("清理完成!")
        return True
    
    # 显示配置信息
    print(f"打包模式: {args.mode}")
    print(f"文件模式: {'目录' if args.onedir else '单文件'}")
    print(f"控制台: {'显示' if args.console else '隐藏'}")
    print(f"自定义Hooks: {'启用' if not args.no_hooks else '禁用'}")
    print("-" * 60)
    
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
        print("\n1. 清理构建目录...")
        clean_build_dirs()
        
        # 2. 构建可执行文件
        print("2. 构建可执行文件...")
        if not build_executable(
            mode=args.mode,
            use_hooks=not args.no_hooks,
            onedir=args.onedir,
            console=args.console
        ):
            return False
        
        # 3. 复制额外文件
        print("3. 复制额外文件...")
        copy_additional_files()
        
        # 4. 清理临时文件
        print("4. 清理临时文件...")
        if os.path.exists('hooks'):
            shutil.rmtree('hooks')
            print("已清理临时hook目录")
        
        # 5. 显示结果
        print("\n" + "=" * 60)
        print("🎉 打包完成!")
        
        if args.onedir:
            print("📁 可执行文件位置: dist/Excel批量处理工具/")
            print("   运行文件: dist/Excel批量处理工具/Excel批量处理工具.exe")
        else:
            print("📄 可执行文件位置: dist/Excel批量处理工具.exe")
        
        # 检查文件大小
        if args.onedir:
            exe_path = Path('dist/Excel批量处理工具/Excel批量处理工具.exe')
        else:
            exe_path = Path('dist/Excel批量处理工具.exe')
            
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📊 文件大小: {size_mb:.1f} MB")
        
        print("=" * 60)
        
        return True
        
    except Exception as e:
        print(f"❌ 打包过程中发生错误: {e}")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1) 