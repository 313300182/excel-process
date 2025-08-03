#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel批量处理工具 - PyInstaller打包脚本
"""

import os
import subprocess
import sys

def build_exe():
    """使用PyInstaller打包exe"""
    
    # 基本配置
    app_name = "Excel数据处理器"
    main_script = "main.py"
    
    # openpyxl需要的隐藏导入模块 - 更全面的列表
    hidden_imports = [
        # openpyxl核心模块
        'openpyxl',
        'openpyxl.cell_writer',
        'openpyxl.workbook',
        'openpyxl.workbook.workbook',
        'openpyxl.worksheet',
        'openpyxl.worksheet.worksheet', 
        'openpyxl.styles',
        'openpyxl.styles.styles',
        'openpyxl.styles.numbers',
        'openpyxl.styles.borders',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.styles.alignment',
        'openpyxl.styles.protection',
        'openpyxl.chart',
        'openpyxl.chart.chart',
        'openpyxl.comments',
        'openpyxl.comments.comments',
        'openpyxl.drawing',
        'openpyxl.drawing.drawing',
        'openpyxl.packaging',
        'openpyxl.packaging.relationship',
        'openpyxl.packaging.manifest',
        'openpyxl.packaging.core',
        'openpyxl.packaging.extended',
        'openpyxl.packaging.custom',
        'openpyxl.utils',
        'openpyxl.utils.exceptions',
        'openpyxl.utils.indexed_list',
        'openpyxl.utils.datetime',
        'openpyxl.utils.units',
        'openpyxl.utils.dataframe',
        'openpyxl.xml',
        'openpyxl.xml.functions',
        'openpyxl.xml.constants',
        'openpyxl.reader',
        'openpyxl.reader.excel',
        'openpyxl.writer',
        'openpyxl.writer.excel',
        # 依赖库
        'et_xmlfile',
        'et_xmlfile.xmlfile',
        'xlrd',
        # tkinter模块
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'tkinter.ttk',
    ]
    
    # 需要排除的包列表 - 这些包通常不需要，可以显著减小exe体积
    excluded_modules = [
        # 测试相关
        'pytest', 'unittest', 'nose', 'coverage',
        # 开发工具
        'IPython', 'jupyter', 'notebook', 'ipykernel',
        # 网络相关 (如果不需要网络功能)
        'requests', 'urllib3', 'certifi', 'chardet', 'idna',
        # 数据科学库 (如果不需要)
        'numpy', 'pandas', 'scipy', 'matplotlib', 'seaborn',
        # 图像处理
        'PIL', 'Pillow', 'cv2',
        # Web框架
        'flask', 'django', 'tornado', 'fastapi',
        # 其他大型库
        'tensorflow', 'torch', 'sklearn', 'pyqt5', 'pyqt6',
        # 文档生成
        'sphinx', 'docutils',
        # 调试工具 (不包括pdb，可能需要)
        'debugpy',
        # 加密相关 (如果不需要)
        'cryptography', 'pyopenssl',
    ]
    
    # 构建PyInstaller命令
    cmd = [
        'pyinstaller',
        '--onefile',                    # 打包成单个exe文件
        '--windowed',                   # 无控制台窗口(GUI应用)
        '--clean',                      # 清理临时文件
        f'--name={app_name}',          # 指定exe文件名
        '--optimize=2',                 # Python字节码优化
        '--noconfirm',                  # 覆盖输出目录而不询问
        '--additional-hooks-dir=.',     # 使用当前目录的hook文件
        '--collect-all=openpyxl',       # 收集openpyxl的所有子模块
        '--collect-all=et_xmlfile',     # 收集et_xmlfile的所有子模块
    ]
    
    # 添加隐藏导入的模块
    for module in hidden_imports:
        cmd.append(f'--hidden-import={module}')
    
    # 添加排除的模块
    for module in excluded_modules:
        cmd.append(f'--exclude-module={module}')
    
    # 如果有模板文件，添加数据文件
    template_path = 'templates/salary_template_example.xlsx'
    if os.path.exists(template_path):
        cmd.append(f'--add-data={template_path};templates/')
    
    # 如果有配置文件，添加数据文件
    config_file = 'salary_user_config.json'
    if os.path.exists(config_file):
        cmd.append(f'--add-data={config_file};.')
    
    # 添加主脚本
    cmd.append(main_script)
    
    print("开始使用PyInstaller打包...")
    print(f"命令: {' '.join(cmd)}")
    print()
    
    try:
        # 执行打包命令
        result = subprocess.run(cmd, check=True, text=True, encoding='utf-8')
        
        print("\n✅ 打包成功！")
        print(f"可执行文件位于: dist/{app_name}.exe")
        
        # 显示文件大小
        exe_path = f"dist/{app_name}.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"文件大小: {size_mb:.1f} MB")
        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 打包失败！返回码: {e.returncode}")
        print("请检查错误信息并重试")
        
    except FileNotFoundError:
        print("\n❌ 错误: PyInstaller未找到")
        print("请先安装PyInstaller: pip install pyinstaller")
        return False
        
    except Exception as e:
        print(f"\n❌ 打包过程中发生错误: {e}")
        return False
    
    return True

def clean_build():
    """清理构建文件"""
    import shutil
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['*.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已删除: {dir_name}/")
    
    # 清理spec文件
    for file in os.listdir('.'):
        if file.endswith('.spec'):
            os.remove(file)
            print(f"已删除: {file}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == 'clean':
        print("清理构建文件...")
        clean_build()
        print("清理完成！")
    else:
        success = build_exe()
        if success:
            print("\n🎉 打包完成！运行以下命令测试:")
            print(f"   .\\dist\\Excel数据处理器.exe") 