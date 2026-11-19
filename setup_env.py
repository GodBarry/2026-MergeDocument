#!/usr/bin/env python3
"""Install this project into the current Python environment."""

import argparse
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description="在当前 Python 环境中安装项目")
    parser.add_argument("--full-install", action="store_true", help="同时安装可选 OCR 和本地校对依赖")
    parser.add_argument("--skip-ocr", action="store_true", help="跳过可选 OCR 依赖")
    parser.add_argument("--skip-ai", action="store_true", help="跳过可选本地校对依赖")
    args = parser.parse_args()

    project = Path(__file__).resolve().parent
    subprocess.run([sys.executable, "-m", "pip", "install", str(project)], check=True)
    extras = []
    if args.full_install and not args.skip_ocr:
        extras.extend(("pix2tex", "latex2mathml"))
    if args.full_install and not args.skip_ai:
        extras.append("pycorrector")
    if extras:
        subprocess.run([sys.executable, "-m", "pip", "install", *extras], check=True)
    if args.full_install and not args.skip_ai:
        print("文本校对功能仍需要项目缺失的 ai_heading_analyzer 模块。")
    print("安装完成。启动桌面界面：python -m mddm.merge_gui")
    print("使用说明：docs/部署与使用.md")


if __name__ == "__main__":
    main()
