"""
标题处理工具（纯 Python 实现，不依赖 ONLYOFFICE Document Builder）
可用于单个文档或合并过程中的标题处理

用法:
    # 处理单个文档
    python process_headings.py document.docx

    # 指定输出文件
    python process_headings.py document.docx -o output.docx

    # 指定编号风格
    python process_headings.py document.docx --style numeric

    # 仅检测不修改
    python process_headings.py document.docx --detect-only
"""

import argparse
import sys
import os
import zipfile
import tempfile
import shutil
import re
import time
from lxml import etree

# 添加项目路径
sys.path.insert(0, os.path.dirname(__file__))

from heading_processor import (
    HeadingProcessor,
    HeadingInfo,
    ProcessingReport
)

# OOXML 命名空间
WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W_TAG = lambda t: f'{{{WNS}}}{t}'

# 注册命名空间（避免 lxml 输出 ns0 前缀）
etree.register_namespace('w', WNS)
etree.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
etree.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
etree.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
etree.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')


def _safe_parse_xml(xml_path: str) -> etree._ElementTree:
    """安全解析 XML，处理 BOM 和编码问题"""
    try:
        with open(xml_path, 'rb') as f:
            raw = f.read()
        # 去除 UTF-8 BOM
        if raw.startswith(b'\xef\xbb\xbf'):
            raw = raw[3:]
        # 去除 UTF-16 BOM
        if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
            text = raw.decode('utf-16')
            raw = text.encode('utf-8')
        # 修正编码声明
        content = raw.decode('utf-8', errors='replace')
        content = re.sub(
            r'<\?xml[^>]*encoding="[^"]*"',
            '<?xml version="1.0" encoding="UTF-8"',
            content,
            count=1
        )
        root = etree.fromstring(content.encode('utf-8'))
        return etree.ElementTree(root)
    except Exception:
        raise


def extract_paragraphs_from_document(doc_path: str) -> list:
    """
    从文档中提取段落信息（纯 Python zipfile + lxml 实现）
    """
    print(f"\n正在读取文档: {doc_path}")

    paragraphs = []

    try:
        with zipfile.ZipFile(doc_path, 'r') as zf:
            if 'word/document.xml' not in zf.namelist():
                print(f"⚠ 文档中无 document.xml")
                return paragraphs

            with zf.open('word/document.xml') as f:
                doc_tree = etree.parse(f)
            doc_root = doc_tree.getroot()
            body = doc_root.find(W_TAG('body'))

            if body is None:
                return paragraphs

            # 预读样式映射
            style_map = _load_styles(zf)

            # 预读自动编号
            numbering_map = _load_numbering(zf)

            para_index = 0
            for child in body:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

                if tag != 'p':
                    continue

                # 提取纯文本
                all_texts = [t.text for t in child.iter(W_TAG('t')) if t.text]
                raw_text = ''.join(all_texts).replace('\r', '').replace('\n', '').strip()

                if not raw_text:
                    continue

                # 提取段落属性
                pPr = child.find(W_TAG('pPr'))

                # 字体大小
                font_size = 12
                if pPr is not None:
                    rPr = pPr.find(W_TAG('rPr'))
                    if rPr is not None:
                        sz = rPr.find(W_TAG('sz'))
                        if sz is not None:
                            val = sz.get(W_TAG('val'), '')
                            if val.isdigit():
                                font_size = int(val) // 2  # Word 字号是半磅值

                # 是否加粗
                is_bold = False
                if pPr is not None:
                    rPr = pPr.find(W_TAG('rPr'))
                    if rPr is not None:
                        b = rPr.find(W_TAG('b'))
                        if b is not None:
                            val = b.get(W_TAG('val'), '1')
                            is_bold = val not in ('0', 'false', 'off')

                # 对齐
                is_centered = False
                if pPr is not None:
                    jc = pPr.find(W_TAG('jc'))
                    if jc is not None:
                        val = jc.get(W_TAG('val'), 'left')
                        is_centered = (val == 'center')

                # 间距
                spacing_before = 0
                spacing_after = 0
                if pPr is not None:
                    sp = pPr.find(W_TAG('spacing'))
                    if sp is not None:
                        bef = sp.get(W_TAG('before'), '0')
                        aft = sp.get(W_TAG('after'), '0')
                        spacing_before = int(bef) if bef.isdigit() else 0
                        spacing_after = int(aft) if aft.isdigit() else 0

                # 样式名称
                style_name = ""
                if pPr is not None:
                    pStyle = pPr.find(W_TAG('pStyle'))
                    if pStyle is not None:
                        sid = pStyle.get(W_TAG('val'), '')
                        style_name = style_map.get(sid, sid)

                # 自动编号
                ilvl = None
                num_id = None
                if pPr is not None:
                    numPr = pPr.find(W_TAG('numPr'))
                    if numPr is not None:
                        nid = numPr.find(W_TAG('numId'))
                        ilv = numPr.find(W_TAG('ilvl'))
                        if nid is not None:
                            num_id_str = nid.get(W_TAG('val'), '')
                            if num_id_str.isdigit():
                                num_id = int(num_id_str)
                        if ilv is not None:
                            ilvl_str = ilv.get(W_TAG('val'), '')
                            if ilvl_str.isdigit():
                                ilvl = int(ilvl_str)

                paragraphs.append({
                    'index': para_index,
                    'text': raw_text,
                    'font_size': font_size,
                    'is_bold': is_bold,
                    'is_centered': is_centered,
                    'spacing_before': spacing_before,
                    'spacing_after': spacing_after,
                    'ilvl': ilvl,
                    'num_id': num_id,
                    'style_name': style_name,
                    '_xml_element': child,  # 保留引用用于应用修改
                })
                para_index += 1

        print(f"✓ 读取完成，共 {len(paragraphs)} 个段落")
        return paragraphs

    except Exception as e:
        print(f"⚠ 读取文档失败: {e}")
        import traceback; traceback.print_exc()
        return paragraphs


def _load_styles(zf: zipfile.ZipFile) -> dict:
    """从 word/styles.xml 加载样式映射"""
    styles = {}
    try:
        if 'word/styles.xml' not in zf.namelist():
            return styles
        with zf.open('word/styles.xml') as f:
            st_root = etree.parse(f).getroot()
        for sty in st_root.findall(f'.//{W_TAG("style")}'):
            sid = sty.get(W_TAG('styleId'), '')
            if not sid:
                continue
            name_el = sty.find(W_TAG('name'))
            sname = ''
            if name_el is not None:
                sname = name_el.get(W_TAG('val'), '')
            styles[sid] = sname
    except Exception:
        pass
    return styles


def _load_numbering(zf: zipfile.ZipFile) -> dict:
    """从 word/numbering.xml 加载编号映射"""
    # 简单占位实现，可扩展
    return {}


def apply_headings_to_document(doc_path: str, headings: list, output_path: str):
    """
    将处理后的标题应用到文档（纯 Python zipfile + lxml 实现）
    """
    print(f"\n正在应用标题修正到文档...")

    temp_dir = tempfile.mkdtemp()
    temp_output = doc_path + '.heading.tmp'

    try:
        # 1. 解压
        with zipfile.ZipFile(doc_path, 'r') as zf:
            zf.extractall(temp_dir)

        doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

        # 2. 解析 document.xml
        tree = _safe_parse_xml(doc_xml_path)
        root = tree.getroot()
        body = root.find(W_TAG('body'))

        if body is None:
            print("⚠ 未找到 body 元素，跳过")
            return

        # 建立段落索引映射
        block_tags = {W_TAG('p'), W_TAG('tbl'), W_TAG('sdt')}
        body_children = [c for c in body if c.tag in block_tags]

        # 建立索引 -> 元素的映射
        idx_to_element = {}
        for i, child in enumerate(body_children):
            if child.tag == W_TAG('p'):
                idx_to_element[i] = child

        # 3. 应用标题修改
        heading_map = {h.index: h for h in headings if not h.is_placeholder and h.index >= 0}

        applied_count = 0
        for idx, heading in heading_map.items():
            if idx not in idx_to_element:
                continue

            para = idx_to_element[idx]

            # 清空段落内容，保留 pPr
            pPr = para.find(W_TAG('pPr'))
            para.clear()  # 清空所有子元素
            if pPr is not None:
                para.append(pPr)

            # 设置样式
            if pPr is None:
                pPr = etree.SubElement(para, W_TAG('pPr'))

            # 清除现有样式设置
            existing_style = pPr.find(W_TAG('pStyle'))
            if existing_style is not None:
                pPr.remove(existing_style)

            # 清除自动编号（防止双重编号）
            existing_numPr = pPr.find(W_TAG('numPr'))
            if existing_numPr is not None:
                pPr.remove(existing_numPr)

            # 设置新的样式
            style_name = f"Heading {heading.true_level}"
            pStyle = etree.SubElement(pPr, W_TAG('pStyle'))
            pStyle.set(W_TAG('val'), style_name)

            # 清除现有的文字运行
            for r in list(para.findall(W_TAG('r'))):
                para.remove(r)

            # 添加新文本 run
            r_el = etree.SubElement(para, W_TAG('r'))
            rPr_el = etree.SubElement(r_el, W_TAG('rPr'))

            # 设置加粗
            if heading.is_bold:
                etree.SubElement(rPr_el, W_TAG('b'))

            # 设置字体大小
            if heading.font_size and heading.font_size > 0:
                sz_el = etree.SubElement(rPr_el, W_TAG('sz'))
                sz_el.set(W_TAG('val'), str(heading.font_size * 2))  # 恢复半磅值
                szCs_el = etree.SubElement(rPr_el, W_TAG('szCs'))
                szCs_el.set(W_TAG('val'), str(heading.font_size * 2))

            t_el = etree.SubElement(r_el, W_TAG('t'))
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t_el.text = heading.text

            applied_count += 1

        # 4. 写回 document.xml
        tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        # 5. 重新打包
        with zipfile.ZipFile(temp_output, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            mimetype_path = os.path.join(temp_dir, 'mimetype')
            if os.path.exists(mimetype_path):
                zf_out.write(mimetype_path, 'mimetype', compress_type=zipfile.ZIP_STORED)
            for root_dir, _, files in os.walk(temp_dir):
                for fname in files:
                    if fname == 'mimetype':
                        continue
                    fp = os.path.join(root_dir, fname)
                    arc = os.path.relpath(fp, temp_dir)
                    zf_out.write(fp, arc, compress_type=zipfile.ZIP_DEFLATED)

        # 6. 原子替换
        replaced = False
        last_err = None
        for _ in range(6):
            try:
                os.replace(temp_output, output_path)
                replaced = True
                break
            except Exception as _re:
                last_err = _re
                time.sleep(0.2)
        if not replaced:
            raise last_err

        print(f"✓ 文档已保存: {output_path} (应用了 {applied_count} 处标题修正)")

    except Exception as e:
        print(f"⚠ 应用标题失败: {e}")
        import traceback; traceback.print_exc()
        if os.path.exists(temp_output):
            try: os.remove(temp_output)
            except Exception: pass
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description='智能标题处理工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 处理单个文档
  python process_headings.py document.docx

  # 指定输出文件和编号风格
  python process_headings.py document.docx -o output.docx --style numeric

  # 仅检测标题，不修改
  python process_headings.py document.docx --detect-only

编号风格:
  auto     - 自动检测（默认）
  numeric  - 数字编号 (1, 1.1, 1.1.1)
  chinese  - 中文编号 (一、二、三)
  mixed    - 混合编号 (第一章, 1.1, 1.1.1)
        """
    )

    parser.add_argument('--input', default='BY1617126.docx', help='输入文档路径')
    parser.add_argument('-o', '--output', default='BY1617126_out.docx', help='输出文档路径（默认覆盖原文件）')
    parser.add_argument('--style', choices=['auto', 'numeric', 'chinese', 'mixed'],
                       default='auto', help='编号风格（默认: auto）')
    parser.add_argument('--detect-only', action='store_true',
                       help='仅检测标题，不修改文档')
    parser.add_argument('--no-numbering', action='store_true',
                       help='不修正序号')
    parser.add_argument('--no-hierarchy', action='store_true',
                       help='不修正层次结构')
    parser.add_argument('--report', help='报告输出路径')

    args = parser.parse_args()

    # 检查输入文件
    if not os.path.exists(args.input):
        print(f"错误: 文件不存在: {args.input}")
        sys.exit(1)

    # 确定输出路径
    if args.output:
        output_path = args.output
    elif args.detect_only:
        output_path = None
    else:
        output_path = args.input

    # 确定报告路径
    if args.report:
        report_path = args.report
    else:
        base_name = os.path.splitext(args.input)[0]
        report_path = f"{base_name}_heading_report.txt"

    try:
        # 1. 读取文档
        paragraphs = extract_paragraphs_from_document(args.input)

        # 2. 处理标题
        processor = HeadingProcessor(numbering_style=args.style)
        headings, report = processor.process_document(
            paragraphs,
            detect_headings=True,
            infer_levels=True,
            correct_numbering=not args.no_numbering,
            correct_hierarchy=not args.no_hierarchy
        )

        # 3. 生成报告
        report_text = processor.generate_report(report_path)
        print("\n" + report_text)

        # 4. 应用到文档
        if not args.detect_only and output_path:
            apply_headings_to_document(args.input, headings, output_path)

        print("\n✓ 处理完成！")

        if args.detect_only:
            print("\n提示: 使用 --detect-only 模式，文档未被修改")
            print(f"     查看报告: {report_path}")

    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
