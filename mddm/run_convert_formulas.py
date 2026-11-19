#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_convert_formulas.py
=======================
将 .docx 文件中的 MathType OLE 公式批量转换为 MS OMML（Office 数学公式）。

内部调用 mathtype_converter.py 的转换函数，转换流程：
  MathType OLE 二进制  →  MathML（via Ruby bridge）
                       →  OMML（via MML2OMML.XSL）
                       →  替换原文档中的 OLE 对象

用法示例：
  # 最简：转换当前目录下的 BY1617126.docx（输出 converted_BY1617126.docx）
  python3 run_convert_formulas.py

  # 指定输入输出
  python3 run_convert_formulas.py --input 论文.docx --output 论文_converted.docx

  # 二进制模式（只转 MathType OLE，速度最快，推荐）
  python3 run_convert_formulas.py --input BY1617126.docx --strategy binary

  # 混合模式（优先二进制，失败后 OCR 图片公式兜底）
  python3 run_convert_formulas.py --input BY1617126.docx --strategy hybrid --priority binary

  # 启用图片公式扫描（非 OLE 的内嵌图片公式也尝试识别）
  python3 run_convert_formulas.py --input BY1617126.docx --scan-images

  # 调试单个 OLE bin 文件
  python3 run_convert_formulas.py --debug-bin word/embeddings/oleObject1.bin
"""

import argparse
import os
import sys

def main():
    parser = argparse.ArgumentParser(
        prog="run_convert_formulas.py",
        description="将 DOCX 中的 MathType 公式批量转换为 MS OMML（Office 数学公式）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )

    # ── 输入输出 ──────────────────────────────────────────────────────
    parser.add_argument(
        "--input", "-i",
        default="BY1617126.docx",
        metavar="DOCX",
        help="待转换的 .docx 文件路径（默认: BY1617126.docx）"
    )
    parser.add_argument(
        "--output", "-o",
        default="xx.docx",
        metavar="DOCX",
        help="输出文件路径（默认: 与原文件同目录，前缀加 converted_）"
    )

    # ── 转换策略 ──────────────────────────────────────────────────────
    parser.add_argument(
        "--strategy", "-s",
        choices=["binary", "ocr", "hybrid"],
        default="binary",
        help=(
            "公式转换策略：\n"
            "  binary  仅读取 MathType OLE 二进制（最快、最稳定，推荐）\n"
            "  ocr     仅对公式图片做 OCR 识别（需安装 pix2tex）\n"
            "  hybrid  先按 --priority 顺序尝试，失败再换另一路\n"
            "（默认: binary）"
        )
    )
    parser.add_argument(
        "--priority",
        choices=["binary", "ocr"],
        default="binary",
        help="hybrid 模式下的优先转换方法（默认: binary）"
    )

    # ── 图片公式扫描 ──────────────────────────────────────────────────
    parser.add_argument(
        "--scan-images",
        action="store_true",
        default=False,
        help=(
            "同时扫描 <w:drawing> 内嵌非 OLE 图片，"
            "尝试用 OCR 识别图片中的公式（需安装 pix2tex）"
        )
    )
    parser.add_argument(
        "--min-height",
        type=float,
        default=2.5,
        metavar="MM",
        help="图片扫描：最小图片高度（mm），小于此值不视为公式图片（默认: 2.5）"
    )
    parser.add_argument(
        "--max-height",
        type=float,
        default=38,
        metavar="MM",
        help="图片扫描：最大图片高度（mm），大于此值不视为公式图片（默认: 38）"
    )

    # ── 并行 ──────────────────────────────────────────────────────────
    parser.add_argument(
        "--workers", "-j",
        type=int,
        default=None,
        metavar="N",
        help="并行 OCR 线程数（默认: CPU 核心数）"
    )

    # ── OCR 补救 ─────────────────────────────────────────────────────
    parser.add_argument(
        "--ocr-rescue",
        action="store_true",
        default=False,
        help="对二进制转换失败的公式尝试 OCR 补救，结果以蓝色标注（默认关闭）"
    )

    # ── 调试工具 ──────────────────────────────────────────────────────
    parser.add_argument(
        "--debug-bin",
        default=None,
        metavar="FILE",
        help=(
            "调试模式：转换单个 OLE .bin 文件，打印 MathML 和 OMML 结果，"
            "不修改任何文档（与 --input 互斥）"
        )
    )
    parser.add_argument(
        "--xsl",
        default=None,
        metavar="FILE",
        help="自定义 MML2OMML.XSL 路径（默认：与脚本同目录）"
    )

    args = parser.parse_args()

    # ── 调试模式：单个 .bin 文件 ──────────────────────────────────────
    if args.debug_bin:
        from . import mathtype_converter

        if not os.path.exists(args.debug_bin):
            print(f"错误：文件不存在: {args.debug_bin}")
            sys.exit(1)

        with open(args.debug_bin, 'rb') as f:
            data = f.read()

        if not mathtype_converter.is_mathtype_ole(data):
            print("⚠ 该文件不是有效的 MathType OLE 二进制（魔数不匹配）")
            sys.exit(1)

        print(f"=== 文件: {args.debug_bin} ({len(data)} bytes) ===")
        print("\n--- Step 1: MathML ---")
        try:
            mathml = mathtype_converter.convert_mathtype_to_mathml(data)
            print(mathml)
        except Exception as e:
            print(f"MathML 转换失败: {e}")
            sys.exit(1)

        print("\n--- Step 2: OMML (含矩阵括号修复) ---")
        try:
            omml = mathtype_converter.convert_mathml_to_omml(mathml, args.xsl)
            print(omml)
        except Exception as e:
            print(f"OMML 转换失败: {e}")
            sys.exit(1)
        return

    # ── 批量转换模式 ──────────────────────────────────────────────────
    from .convert_docx_formulas import convert_document

    docx_path = args.input
    if not os.path.exists(docx_path):
        print(f"错误：输入文件不存在: {docx_path}")
        sys.exit(1)

    # 如果用户指定了 --output，把 convert_document 的输出移动过去
    # （convert_document 内部固定以 converted_ 为前缀，这里通过重命名实现）
    import shutil
    auto_output = os.path.join(
        os.path.dirname(os.path.abspath(docx_path)),
        f"converted_{os.path.basename(docx_path)}"
    )

    print(f"{'='*55}")
    print(f"  输入文件  : {docx_path}")
    print(f"  输出文件  : {args.output or auto_output}")
    print(f"  转换策略  : {args.strategy}" +
          (f"（优先 {args.priority}）" if args.strategy == 'hybrid' else ''))
    print(f"  图片扫描  : {'开启' if args.scan_images else '关闭'}")
    print(f"  OCR补救   : {'开启' if args.ocr_rescue else '关闭'}")
    print(f"  并行线程  : {args.workers or 'auto'}")
    print(f"{'='*55}\n")

    convert_document(
        docx_path=docx_path,
        strategy=args.strategy,
        max_workers=args.workers,
        scan_images=args.scan_images,
        min_height_mm=args.min_height,
        max_height_mm=args.max_height,
        priority=args.priority,
        ocr_rescue=args.ocr_rescue,
    )

    # 如果用户指定了 --output 且与默认输出路径不同，重命名
    if args.output and os.path.abspath(args.output) != os.path.abspath(auto_output):
        if os.path.exists(auto_output):
            os.makedirs(os.path.dirname(os.path.abspath(args.output)) or '.', exist_ok=True)
            shutil.move(auto_output, args.output)
            print(f"已重命名输出文件为: {args.output}")


if __name__ == "__main__":
    main()
