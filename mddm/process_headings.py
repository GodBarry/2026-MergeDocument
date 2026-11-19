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
from lxml import etree

from .heading_processor import (
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


def _heading_text_nodes(paragraph):
    """Visible heading text, excluding nested paragraphs and deleted runs."""
    return [node for node in paragraph.iter(W_TAG('t'))
            if next(node.iterancestors(W_TAG('p')), None) is paragraph
            and next(node.iterancestors(W_TAG('del')), None) is None]


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

            for para_index, child in enumerate(body.findall(W_TAG('p'))):
                # 提取纯文本
                all_texts = [t.text for t in _heading_text_nodes(child) if t.text]
                raw_text = ''.join(all_texts).strip()

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

        print(f"✓ 读取完成，共 {len(paragraphs)} 个段落")
        return paragraphs

    except (OSError, zipfile.BadZipFile, etree.XMLSyntaxError) as error:
        raise ValueError(f"读取标题失败 {doc_path}: {error}") from error


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


def _replace_heading_number(nodes, original, updated):
    """Replace the changed prefix while preserving runs and comment or bookmark anchors."""
    text = ''.join(node.text or '' for node in nodes)
    start = len(text) - len(text.lstrip())
    left = 0
    while left < min(len(original), len(updated)) and original[left] == updated[left]:
        left += 1
    right = 0
    while (right < min(len(original), len(updated)) - left
           and original[-right-1] == updated[-right-1]):
        right += 1
    begin, end = start + left, start + len(original) - right
    insert = updated[left:len(updated)-right]
    pos = 0
    inserted = False
    for node in nodes:
        value = node.text or ''
        limit = pos + len(value)
        if (pos <= begin < limit or begin == limit == len(text)) and not inserted:
            node.text = value[:begin-pos] + insert + value[max(end-pos, 0):]
            inserted = True
        elif pos < end and limit > begin:
            node.text = value[max(end-pos, 0):]
        if node.text and (node.text[0].isspace() or node.text[-1].isspace()):
            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        pos = limit


def apply_headings_to_document(doc_path: str, headings: list, output_path: str):
    """Apply heading edits to their original top-level paragraphs."""
    output_dir = os.path.dirname(os.path.abspath(output_path))
    temp_path = None
    try:
        with zipfile.ZipFile(doc_path) as package:
            root = etree.fromstring(package.read('word/document.xml'))
            body = root.find(W_TAG('body'))
            if body is None:
                raise ValueError("文档中缺少正文")
            paragraphs = body.findall(W_TAG('p'))
            heading_styles = {}
            for style_id, name in _load_styles(package).items():
                match = re.fullmatch(r'(?:heading|标题)\s*([1-9])', name, re.I)
                if match:
                    heading_styles[int(match.group(1))] = style_id

            applied = 0
            for heading in headings:
                if heading.is_placeholder or heading.index < 0:
                    continue
                para = paragraphs[heading.index]
                nodes = _heading_text_nodes(para)
                actual = ''.join(node.text or '' for node in nodes).strip()
                original = heading.original_text or heading.text
                if actual != original:
                    raise ValueError(f"标题位置或文本发生变化，段落 {heading.index + 1}")
                before = etree.tostring(para)
                props = para.find(W_TAG('pPr'))
                if props is None:
                    props = etree.Element(W_TAG('pPr'))
                    para.insert(0, props)
                style_id = heading_styles.get(heading.true_level)
                if style_id:
                    style = props.find(W_TAG('pStyle'))
                    if style is None:
                        style = etree.Element(W_TAG('pStyle'))
                        props.insert(0, style)
                    style.set(W_TAG('val'), style_id)
                if heading.true_level:
                    outline = props.find(W_TAG('outlineLvl'))
                    if outline is None and not style_id:
                        outline = etree.Element(W_TAG('outlineLvl'))
                        following_tags = {W_TAG(tag) for tag in ('divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')}
                        insert_at = next((i for i, item in enumerate(props) if item.tag in following_tags), len(props))
                        props.insert(insert_at, outline)
                    if outline is not None:
                        outline.set(W_TAG('val'), str(heading.true_level - 1))
                if heading.text != original:
                    _replace_heading_number(nodes, original, heading.text)
                    numbering = props.find(W_TAG('numPr'))
                    if numbering is not None:
                        props.remove(numbering)
                if before != etree.tostring(para):
                    applied += 1

            fd, temp_path = tempfile.mkstemp(dir=output_dir, suffix='.heading.tmp')
            os.close(fd)
            with zipfile.ZipFile(temp_path, 'w') as output:
                for part in package.infolist():
                    if part.filename == 'word/document.xml':
                        output.writestr(part, etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))
                    else:
                        with package.open(part) as source, output.open(part, 'w') as target:
                            shutil.copyfileobj(source, target)
        os.replace(temp_path, output_path)
        print(f"✓ 文档已保存: {output_path} (调整了 {applied} 个标题)")
    finally:
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)


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
