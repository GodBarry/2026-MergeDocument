#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动环境配置脚本 / Automatic Environment Setup Script
用途: 自动检测并安装项目所需的所有依赖
Usage: Automatically detect and install all project dependencies

运行方式 / How to run:
    python setup_env.py
    python setup_env.py --full-install
    python setup_env.py --skip-ocr --skip-ai
"""

import sys
import subprocess
import argparse
import os
from pathlib import Path

# ANSI 颜色代码
class Colors:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    GRAY = '\033[90m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

def print_colored(text, color=''):
    """打印彩色文本"""
    print(f"{color}{text}{Colors.RESET}")

def check_command(cmd):
    """检查命令是否可用"""
    try:
        subprocess.run([cmd, '--version'], capture_output=True, check=True)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False

def check_python_package(package_name):
    """检查 Python 包是否已安装"""
    try:
        __import__(package_name)
        return True
    except ImportError:
        return False

def install_package(package_name, display_name=None):
    """安装 Python 包"""
    display_name = display_name or package_name
    print(f"  正在安装 {display_name}...", end=' ')
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', package_name], 
                      capture_output=True, check=True)
        print_colored("✓ 成功", Colors.GREEN)
        return True
    except subprocess.CalledProcessError as e:
        print_colored("✗ 失败", Colors.RED)
        print(f"    错误: {e.stderr.decode()}")
        return False

def main():
    parser = argparse.ArgumentParser(description='自动配置项目环境')
    parser.add_argument('--full-install', action='store_true', 
                       help='安装所有可选依赖')
    parser.add_argument('--skip-ocr', action='store_true', 
                       help='跳过 OCR 依赖')
    parser.add_argument('--skip-ai', action='store_true', 
                       help='跳过 AI 依赖')
    args = parser.parse_args()

    print_colored("=" * 60, Colors.CYAN)
    print_colored("  文档处理环境自动配置脚本", Colors.CYAN)
    print_colored("  Document Processing Environment Setup", Colors.CYAN)
    print_colored("=" * 60, Colors.CYAN)
    print()

    # [1/6] 检测 Python 版本
    print_colored("[1/6] 检测 Python 环境...", Colors.YELLOW)
    python_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    print_colored(f"  ✓ Python 版本: {python_version}", Colors.GREEN)
    
    if sys.version_info < (3, 8):
        print_colored("  ✗ 错误: 需要 Python 3.8+", Colors.RED)
        sys.exit(1)

    # [2/6] 检测虚拟环境
    print()
    print_colored("[2/6] 检测虚拟环境...", Colors.YELLOW)
    venv_active = hasattr(sys, 'real_prefix') or (
        hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix
    )
    
    if venv_active:
        print_colored(f"  ✓ 虚拟环境已激活: {sys.prefix}", Colors.GREEN)
    else:
        print_colored("  ⚠ 未检测到虚拟环境", Colors.YELLOW)
        print_colored("  建议: 使用虚拟环境 (conda activate office)", Colors.YELLOW)

    # [3/6] 安装核心依赖
    print()
    print_colored("[3/6] 安装核心依赖...", Colors.YELLOW)
    
    core_packages = [
        ('docbuilder', 'document-builder', 'ONLYOFFICE Document Builder'),
        ('lxml', 'lxml', 'lxml (XML 处理)'),
        ('PIL', 'Pillow', 'Pillow (图像处理)'),
    ]
    
    for import_name, package_name, display_name in core_packages:
        print(f"  检查 {display_name}...", end=' ')
        if check_python_package(import_name):
            print_colored("✓ 已安装", Colors.GREEN)
        else:
            print_colored("✗ 未安装", Colors.RED)
            install_package(package_name, display_name)

    # [4/6] 安装 OCR 依赖
    print()
    if not args.skip_ocr and args.full_install:
        print_colored("[4/6] 安装 OCR 依赖 (可选)...", Colors.YELLOW)
        
        ocr_packages = [
            ('pix2tex', 'pix2tex', 'pix2tex (OCR 引擎)'),
            ('latex2mathml', 'latex2mathml', 'latex2mathml (LaTeX 转换)'),
        ]
        
        for import_name, package_name, display_name in ocr_packages:
            print(f"  检查 {display_name}...", end=' ')
            if check_python_package(import_name):
                print_colored("✓ 已安装", Colors.GREEN)
            else:
                print_colored("✗ 未安装", Colors.RED)
                install_package(package_name, display_name)
        
        # 检查 ImageMagick
        print("  检查 ImageMagick...", end=' ')
        if check_command('convert'):
            print_colored("✓ 已安装", Colors.GREEN)
        else:
            print_colored("✗ 未安装", Colors.RED)
            print_colored("  ⚠ ImageMagick 未安装 (OCR 模式需要)", Colors.YELLOW)
            print_colored("  下载地址: https://imagemagick.org/script/download.php", Colors.CYAN)
    else:
        print_colored("[4/6] 跳过 OCR 依赖安装 (使用 --full-install 启用)", Colors.GRAY)

    # [5/6] 安装 AI 文本校对依赖
    print()
    if not args.skip_ai and args.full_install:
        print_colored("[5/6] 安装 AI 文本校对依赖 (可选)...", Colors.YELLOW)
        
        print("  检查 pycorrector...", end=' ')
        if check_python_package('pycorrector'):
            print_colored("✓ 已安装", Colors.GREEN)
        else:
            print_colored("✗ 未安装", Colors.RED)
            install_package('pycorrector', 'pycorrector (本地校对)')
    else:
        print_colored("[5/6] 跳过 AI 依赖安装 (使用 --full-install 启用)", Colors.GRAY)

    # [6/6] 验证安装
    print()
    print_colored("[6/6] 验证安装...", Colors.YELLOW)
    
    verify_script = Path(__file__).parent / 'verify_installation.py'
    if verify_script.exists():
        print("  运行验证脚本...")
        subprocess.run([sys.executable, str(verify_script)])
    else:
        print_colored("  ⚠ 未找到 verify_installation.py", Colors.YELLOW)

    # 总结
    print()
    print_colored("=" * 60, Colors.CYAN)
    print_colored("  环境配置完成!", Colors.GREEN)
    print_colored("=" * 60, Colors.CYAN)
    print()
    print_colored("下一步:", Colors.YELLOW)
    print("  1. 确保已激活 'office' 虚拟环境: conda activate office")
    print("  2. 运行文档合并: python run_merge.py --help")
    print("  3. 运行公式转换: python convert_docx_formulas.py --help")
    print()
    print_colored("文档:", Colors.YELLOW)
    print("  - setup_environment.md (环境配置指南)")
    print("  - QUICK_START.md (快速开始)")
    print("  - INSTALLATION_GUIDE.md (详细安装说明)")
    print()

if __name__ == '__main__':
    main()
