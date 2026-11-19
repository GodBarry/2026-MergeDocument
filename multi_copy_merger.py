"""
多副本文档合并系统 - 累积式三路合并
支持 n 个副本的自动合并,冲突检测和详细报告
"""

import os
import sys
import json
import posixpath
import re
from itertools import combinations, groupby
import difflib
import hashlib
import shutil
import zipfile
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Any
from datetime import datetime
from pathlib import Path
import tempfile

from lxml import etree

# 锚点引擎
from anchor_engine import (
    _anchor_norm_text,
    build_anchor_keys as _anchor_build_keys,
    score_anchor_candidate as _anchor_score_candidate,
)

# 自动编号解析器（读取 word/numbering.xml）
try:
    from numbering_resolver import NumberingResolver
    _NUMBERING_RESOLVER_AVAILABLE = True
except ImportError:
    _NUMBERING_RESOLVER_AVAILABLE = False
    class NumberingResolver:  # type: ignore
        def __init__(self, *a, **kw): self._loaded = False
        def resolve(self, *a, **kw): return ""
        def reset(self): pass
        def has_numbering(self): return False


@dataclass
class DocumentChange:
    """文档变更记录"""
    change_id: str
    change_type: str  # 'insert', 'delete', 'format', 'table'
    position: int  # 变更位置
    content: str  # 变更内容
    author: str  # 编辑者
    timestamp: datetime
    copy_index: int = -1  # 副本索引，用于冲突检测
    
    # 位置信息
    section_path: List[str] = field(default_factory=list)  # 章节路径 ['第一章', '1.1 概述']
    page_number: int = 0  # 页码
    paragraph_index: int = 0  # 绝对索引（已弃用，仅保留兼容）
    para_order: int = -1  # 段落序号（同类元素中第N个，用于写回定位）
    
    # 格式信息
    format_info: Optional[Dict] = None
    
    # 表格相关
    table_info: Optional[Dict] = None  # 表格信息


@dataclass
class MergeConflict:
    """合并冲突"""
    conflict_id: str
    position: int  # 冲突位置
    changes: List[DocumentChange]  # 冲突的多个变更
    severity: str  # 'low', 'medium', 'high'
    description: str = ""


@dataclass
class MergeResult:
    """合并结果"""
    merged_document: str  # 合并后的文档路径
    conflicts: List[MergeConflict]
    auto_merged_count: int
    manual_review_count: int
    conflict_report: str = ""  # 冲突报告路径


# =====================================================================
# 新的 OpenXML 文档引擎：纯 Python / lxml 实现，无需 ONLYOFFICE
# =====================================================================

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPC = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
WPG = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingGroup'
WPI = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingInk'
WPM = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingMatric'
WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
COMMENTS_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'

_AUTHOR_COLORS = [
    "FF6B6B",   # 红
    "4ECDC4",   # 青
    "45B7D1",   # 蓝
    "96CEB4",   # 绿
    "FFEAA7",   # 黄
    "DDA0DD",   # 紫
    "FFA07A",   # 橙
    "98D8C8",   # 薄荷
]


def _wtag(n): return f'{{{WNS}}}{n}'
def _atag(n): return f'{{{A}}}{n}'
def _wptag(n): return f'{{{WP}}}{n}'
def _wpeqtag(n): return f'{{{WPC}}}{n}'


_RUN_CONTAINERS = {
    _wtag(name) for name in ('hyperlink', 'ins', 'moveTo', 'fldSimple', 'smartTag',
                            'customXml', 'sdt', 'sdtContent')
}
_REVISION_TAGS = {
    _wtag(name) for name in ('ins', 'del', 'moveFrom', 'moveTo',
                            'pPrChange', 'rPrChange', 'sectPrChange', 'numberingChange')
}
_MOVE_RANGE_TAGS = {
    _wtag(name) for name in ('moveFromRangeStart', 'moveFromRangeEnd',
                            'moveToRangeStart', 'moveToRangeEnd')
}
_REVISION_MARKUP = _REVISION_TAGS | _MOVE_RANGE_TAGS
_IGNORED_PARAGRAPH_TAGS = {
    _wtag(name) for name in ('proofErr', 'commentRangeStart', 'commentRangeEnd', 'commentReference')
}
_IGNORED_PARAGRAPH_ATTRIBUTES = {
    '{http://www.w3.org/XML/1998/namespace}space',
    '{http://schemas.microsoft.com/office/word/2010/wordml}paraId',
    '{http://schemas.microsoft.com/office/word/2010/wordml}textId',
    '{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}anchorId',
    '{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}editId',
}


def _visible_runs(element):
    for child in element:
        if child.tag == _wtag('r'):
            yield child
        elif child.tag in _RUN_CONTAINERS:
            yield from _visible_runs(child)


def _relationship_attribute(name):
    return name.startswith(f'{{{R}}}') or name == '{urn:schemas-microsoft-com:office:office}relid'


def _relationship_signatures(package):
    rels_path = 'word/_rels/document.xml.rels'
    if rels_path not in package.namelist():
        return {}

    signatures = {}
    image_digests = {}
    for rel in _safe_xml_from_zip(package, rels_path):
        rel_type = rel.get('Type', '')
        mode = rel.get('TargetMode', 'Internal')
        target = rel.get('Target', '')
        if mode == 'Internal' and rel_type.endswith('/image'):
            part = posixpath.normpath(posixpath.join('/word', target)).lstrip('/')
            if part not in image_digests:
                try:
                    with package.open(part) as stream:
                        digest = hashlib.sha256()
                        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                            digest.update(chunk)
                        image_digests[part] = digest.hexdigest()
                except KeyError:
                    # Unreferenced legacy relationships may be stale; copying a missing image still fails.
                    image_digests[part] = ('missing', part)
            target = image_digests[part]
        signatures[rel.get('Id')] = (rel_type, mode, target)
    return signatures


def _paragraph_signature(element, relationships):
    """Compare paragraph content independently of Word's local IDs and simple run boundaries."""
    attributes = tuple(
        (name, relationships.get(value, value) if _relationship_attribute(name) else value)
        for name, value in sorted(element.attrib.items())
        if not name.startswith(_wtag('rsid'))
        and name not in _IGNORED_PARAGRAPH_ATTRIBUTES
        and not (name == _wtag('id') and element.tag in _REVISION_MARKUP)
        and not (name == 'id' and element.tag in (_wptag('docPr'), '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr'))
    )
    children = []
    for child in element:
        if not isinstance(child.tag, str) or child.tag in _IGNORED_PARAGRAPH_TAGS:
            continue
        part = _paragraph_signature(child, relationships)
        if child.tag in (_wtag('pPr'), _wtag('rPr')) and not any(part[1:]):
            continue
        if child.tag == _wtag('r'):
            run_children = part[3]
            if not run_children:
                continue
            if (
                (len(run_children) == 1 or len(run_children) == 2 and run_children[0][0] == _wtag('rPr'))
                and run_children[-1][0] == _wtag('t')
                and not run_children[-1][3]
            ):
                run_style = (part[1], run_children[:-1])
                text = run_children[-1][2]
                if children and children[-1][0] == 'text-run' and children[-1][1] == run_style:
                    text = children.pop()[2] + text
                part = ('text-run', run_style, text)
        children.append(part)
    text = element.text or ''
    if element.tag not in (_wtag('t'), _wtag('delText'), _wtag('instrText'), _wtag('delInstrText')) and not text.strip():
        text = ''
    return element.tag, attributes, text, tuple(children)


def _safe_xml_from_zip(zf: zipfile.ZipFile, xml_path: str) -> etree._Element:
    """
    从 ZIP 文件流安全解析 XML，处理以下常见问题：
    1. UTF-8 BOM 头 (EF BB BF) 导致解析失败
    2. 编码声明与实际编码不一致
    3. 命名空间未注册导致 lxml 重写成 ns0 前缀

    Args:
        zf: zipfile.ZipFile 实例
        xml_path: ZIP 内的 XML 文件路径，如 'word/document.xml'

    Returns:
        XML 根元素
    """
    import io
    import re

    with zf.open(xml_path) as f:
        raw = f.read()

    # 去除 BOM
    if raw.startswith(b'\xef\xbb\xbf'):
        raw = raw[3:]
    elif raw.startswith(b'\xff\xfe'):
        raw = raw.decode('utf-16-le')[1:].encode('utf-16-le')
    elif raw.startswith(b'\xfe\xff'):
        raw = raw.decode('utf-16-be')[1:].encode('utf-16-be')

    # 动态注册命名空间
    xml_content = raw.decode('utf-8', errors='replace')
    for m in re.finditer(r'xmlns:([a-zA-Z0-9_\-]+)="([^"]+)"', xml_content[:4096]):
        try:
            etree.register_namespace(m.group(1), m.group(2))
        except ValueError:
            pass

    # 修正编码声明
    xml_content = re.sub(
        r'<\?xml[^>]*encoding="[^"]*"',
        '<?xml version="1.0" encoding="UTF-8"',
        xml_content,
        count=1
    )

    return etree.fromstring(xml_content.encode('utf-8'))


class DocItem:
    """统一文档块：段落 / 表格 / 图片 / SDT / 其他 body 子元素"""
    __slots__ = (
        'index', 'type', 'content', 'item_id', 'xml',
        'para_props', 'runs', 'images', 'rows', 'table_props',
        'heading_score', 'style_name', 'page',
        'level', 'detected_level', 'body_tag', 'paragraph_signature',
    )
    def __init__(self):
        self.index: int = 0
        self.type: str = 'paragraph'
        self.content: str = ''
        self.item_id: str = ''
        self.xml: etree._Element = None
        self.para_props: dict = {}
        self.runs: List[dict] = []
        self.images: List[dict] = []
        self.rows: List[List[dict]] = []
        self.table_props: dict = {}
        self.heading_score: float = 0.0
        self.style_name: str = ''
        self.page: int = 1
        self.level: str = '?'          # 原判级别 (从样式提取，如 '1', '?' )
        self.detected_level: int = 0   # 推断级别 (0=正文, 1=1级...)
        self.body_tag: str = ''   # body 子元素标签名
        self.paragraph_signature = None

    def table_shape(self) -> tuple:
        return tuple(tuple((cell['rowspan'], cell['colspan']) for cell in row) for row in self.rows)

    def table_diff_key(self) -> tuple:
        cell_text = tuple(tuple(cell['text'] for cell in row) for row in self.rows)
        return self.table_shape(), cell_text

    def __repr__(self):
        return f"<DocItem {self.type} #{self.index}: {self.content[:30]!r}>"


class OpenXMLDocumentEngine:
    """纯 Python / lxml OpenXML 文档处理引擎，无需 ONLYOFFICE"""

    def __init__(self):
        self.use_python_api = True
        print("[OK]")  # was ✓

    # ------------------------------------------------------------------
    # 提取文档结构（统一全局 index）
    # ------------------------------------------------------------------
    def extract_document_structure(self, doc_path: str) -> Dict:
        """直接用 lxml 解析 document.xml，统一提取所有 body 直接子元素"""
        try:
            items: List[DocItem] = []
            para_counter = 0
            table_counter = 0
            image_counter = 0

            # 用于估算页码
            current_page = 1
            chars_on_page = 0
            CHARS_PER_PAGE = 1250

            # 预读样式映射
            style_map = self._load_styles(doc_path)

            with zipfile.ZipFile(doc_path) as zf:
                required_parts = {'[Content_Types].xml', '_rels/.rels', 'word/document.xml'}
                missing = required_parts.difference(zf.namelist())
                if missing:
                    raise ValueError(f"缺少 DOCX 部件: {', '.join(sorted(missing))}")

                if _safe_xml_from_zip(zf, '[Content_Types].xml').tag != f'{{{CT_NS}}}Types':
                    raise ValueError("缺少有效的 DOCX 内容类型表")
                if _safe_xml_from_zip(zf, '_rels/.rels').tag != f'{{{REL}}}Relationships':
                    raise ValueError("缺少有效的 DOCX 包关系表")

                doc_root = _safe_xml_from_zip(zf, 'word/document.xml')
                body = doc_root.find(_wtag('body'))
                relationships = _relationship_signatures(zf)

            if doc_root.tag != _wtag('document') or body is None:
                raise ValueError("缺少有效的 Word 正文")

            print(f"    正在分析文档...")

            # 统计 body 直接子元素数量（用于提示）
            all_children_count = len(list(body))
            known_tags = {_wtag('p'), _wtag('tbl'), _wtag('sdt')}
            known_count = len([c for c in body if c.tag in known_tags])
            if all_children_count > known_count:
                print(f"    检测到 {all_children_count} 个 body 子元素（已知类型 {known_count} 个）")
            else:
                print(f"    检测到 {known_count} 个块级元素")

            # 1. 预解析自动编号定义（单实例复用以保持计数状态）
            try:
                nr = NumberingResolver(doc_path)
                if nr.has_numbering():
                    print(f"    [OK] 检测到自动编号定义 ({len(nr.get_num_ids())} 个实例)")
            except Exception:
                nr = None

            # 遍历 body 的所有直接子元素，统一处理
            for child in body:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                item_idx = len(items)

                if tag == 'p':
                    # 段落：提取详细内容
                    item = self._extract_para_item(child, style_map, nr=nr)
                    item.paragraph_signature = _paragraph_signature(child, relationships)
                    para_counter += 1
                    item.body_tag = 'p'
                    items.append(item)
                    item.index = len(items) - 1

                    # 图片由所属段落写回，避免为同一图片额外生成正文占位符。
                    image_counter += len(item.images)

                    # 页码估算
                    chars_on_page += len(item.content) + 50
                    if chars_on_page >= CHARS_PER_PAGE:
                        current_page += 1
                        chars_on_page = 0
                    item.page = current_page

                elif tag == 'tbl':
                    # 表格
                    item = self._extract_table_item(child, table_counter)
                    table_counter += 1
                    item.body_tag = 'tbl'
                    items.append(item)
                    item.index = len(items) - 1
                    # 表格占较大空间
                    chars_on_page += 300
                    if chars_on_page >= CHARS_PER_PAGE:
                        current_page += 1
                        chars_on_page = 0
                    item.page = current_page

                elif tag == 'sdt':
                    # 结构化文档标签
                    item = self._extract_sdt_item(child, item_idx)
                    item.body_tag = 'sdt'
                    items.append(item)
                    item.index = len(items) - 1

                else:
                    # 其他所有 body 子元素（sectPr、bookmarkEnd 等）
                    item = self._extract_misc_item(child, item_idx)
                    items.append(item)
                    item.index = len(items) - 1

            print(f"    [OK] 成功提取: {para_counter} 个段落, {table_counter} 个表格, {image_counter} 个图片, {len(items) - para_counter - table_counter} 个其他元素")

            return {
                'items': items,
                'item_count': len(items),
                '_para_count': para_counter,
                '_table_count': table_counter,
            }

        except Exception as error:
            raise ValueError(f"无法读取 DOCX {doc_path}: {error}") from error

    # ------------------------------------------------------------------
    # 段落提取
    # ------------------------------------------------------------------
    def _load_styles(self, doc_path: str) -> Dict:
        """预读 word/styles.xml，构建完整的样式属性查找表（含继承关系）"""
        styles = {}
        try:
            with zipfile.ZipFile(doc_path) as zf:
                if 'word/styles.xml' not in zf.namelist():
                    return styles
                st_root = _safe_xml_from_zip(zf, 'word/styles.xml')

                # ── 1. 读取 docDefaults ──────────────────────────────────
                defaults = {'east_asia': '', 'sz': '21', 'bold': False, 'outlineLvl': None}
                docDef = st_root.find(_wtag('docDefaults'))
                if docDef is not None:
                    rPrDef = docDef.find(_wtag('rPrDefault'))
                    if rPrDef is not None:
                        rpr = rPrDef.find(_wtag('rPr'))
                        if rpr is not None:
                            rf = rpr.find(_wtag('rFonts'))
                            if rf is not None:
                                defaults['east_asia'] = rf.get(_wtag('eastAsia'), '') or rf.get(_wtag('hAnsi'), '')
                            sz_el = rpr.find(_wtag('sz'))
                            if sz_el is not None:
                                defaults['sz'] = sz_el.get(_wtag('val'), '21')
                            b_el = rpr.find(_wtag('b'))
                            if b_el is not None:
                                val = b_el.get(_wtag('val'), '1')
                                defaults['bold'] = val not in ('0', 'false', 'off')
                styles['_defaults_'] = defaults

                # ── 2. 预读所有样式的原始定义 ──────────────────────────────
                raw_styles = {}
                default_para_style_id = None
                for sty in st_root.findall(f'.//{_wtag("style")}'):
                    sid = sty.get(_wtag('styleId'), '')
                    if not sid: continue
                    
                    name_el = sty.find(_wtag('name'))
                    sname = name_el.get(_wtag('val'), '') if name_el is not None else sid
                    
                    if sty.get(_wtag('type')) == 'paragraph' and sty.get(_wtag('default')) == '1':
                        default_para_style_id = sid

                    # 提取该样式的显式定义内容
                    pPr = sty.find(_wtag('pPr'))
                    rPr = sty.find(_wtag('rPr'))
                    
                    props = {}
                    if pPr is not None:
                        ol = pPr.find(_wtag('outlineLvl'))
                        if ol is not None:
                            props['outlineLvl'] = ol.get(_wtag('val'))
                    
                    if rPr is not None:
                        rf = rPr.find(_wtag('rFonts'))
                        if rf is not None:
                            ea = rf.get(_wtag('eastAsia'), '') or rf.get(_wtag('hAnsi'), '')
                            if ea: props['east_asia'] = ea
                        sz_el = rPr.find(_wtag('sz'))
                        if sz_el is not None:
                            props['sz'] = sz_el.get(_wtag('val'), '')
                        b_el = rPr.find(_wtag('b'))
                        if b_el is not None:
                            val = b_el.get(_wtag('val'), '1')
                            props['bold'] = val not in ('0', 'false', 'off')
                    
                    based_on = sty.find(_wtag('basedOn'))
                    parent_id = based_on.get(_wtag('val')) if based_on is not None else None
                    
                    raw_styles[sid] = {
                        'name': sname,
                        'parent': parent_id,
                        'props': props
                    }

                # ── 3. 递归/迭代解析继承链链 ──────────────────────────────
                def resolve_style(sid, visited):
                    if sid in styles or sid not in raw_styles:
                        return styles.get(sid, {}).get('props', {})
                    if sid in visited: # 防环
                        return {}
                    
                    visited.add(sid)
                    raw = raw_styles[sid]
                    parent_props = {}
                    if raw['parent']:
                        parent_props = resolve_style(raw['parent'], visited)
                    
                    # 合并：自身属性覆盖父类属性
                    final_props = {**defaults, **parent_props, **raw['props']}
                    styles[sid] = {'name': raw['name'], 'props': final_props}
                    return final_props

                for sid in list(raw_styles.keys()):
                    resolve_style(sid, set())

                styles['_default_para_style_'] = default_para_style_id
        except Exception:
            pass
        return styles

    def _extract_para_item(self, para: etree._Element, style_map: Dict,
                           nr: NumberingResolver = None) -> DocItem:
        item = DocItem()
        item.index = -1  # 临时，后续填充
        item.type = 'paragraph'
        item.xml = para

        pPr = para.find(_wtag('pPr'))
        style_id = None
        if pPr is not None:
            pStyle = pPr.find(_wtag('pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(_wtag('val'))

        # 样式继承链：
        # 优先级：本段 pPr/rPr 直接定义 > 样式的 props 定义 (含继承链) > docDefaults
        style_node = style_map.get(style_id)
        if not style_node and not style_id:
            default_sid = style_map.get('_default_para_style_')
            style_node = style_map.get(default_sid)
        
        style_props = style_node.get('props', style_map.get('_defaults_', {})) if style_node else style_map.get('_defaults_', {})
        
        item.style_name = style_node.get('name', style_id or '') if style_node else (style_id or '')
        style_bold = style_props.get('bold', False)
        style_sz = style_props.get('sz', '21')
        style_font = style_props.get('east_asia', '')

        # 收集 runs 和文字
        runs = []
        text_parts = []
        images = []
        cur_bold = style_bold

        for run in _visible_runs(para):
            rPr = run.find(_wtag('rPr'))
            run_bold = cur_bold
            run_font = style_font
            run_size = style_sz
            run_italic = False
            run_underline = ''
            run_color = ''

            if rPr is not None:
                b_el = rPr.find(_wtag('b'))
                if b_el is not None:
                    val = b_el.get(_wtag('val'), '1')
                    run_bold = val not in ('0', 'false', 'off')
                rf = rPr.find(_wtag('rFonts'))
                if rf is not None:
                    # rFonts 中显式指定的字体优先级最高
                    ea = rf.get(_wtag('eastAsia'), '')
                    if ea:
                        run_font = ea
                    else:
                        ha = rf.get(_wtag('hAnsi'), '')
                        if ha:
                            run_font = ha
                sz_el = rPr.find(_wtag('sz'))
                if sz_el is not None:
                    run_size = sz_el.get(_wtag('val'), '')
                i_el = rPr.find(_wtag('i'))
                if i_el is not None:
                    run_italic = True
                u_el = rPr.find(_wtag('u'))
                if u_el is not None:
                    run_underline = u_el.get(_wtag('val'), 'single')
                color_el = rPr.find(_wtag('color'))
                if color_el is not None:
                    run_color = color_el.get(_wtag('val'), '')

            text = ''.join(node.text or '' for node in run.findall(_wtag('t')))
            if text:
                text_parts.append(text)
                runs.append({
                    'text': text,
                    'bold': run_bold,
                    'italic': run_italic,
                    'underline': run_underline,
                    'color': run_color or '000000',
                    'font': run_font,
                    'fontSize': run_size,
                    'background': '',
                })
            for child in run:
                if child.tag in (_wtag('drawing'), _wtag('pict')):
                    text_parts.append(f"[图片{len(images)}]")
                    images.append(child)

        # 合并单元格内容
        for mcg in para.findall(_wtag('mc')):
            for ch in mcg.iter():
                if ch.tag == _wtag('t') and ch.text:
                    text_parts.append(ch.text)

        content = ''.join(text_parts).strip()
        item.content = content
        item.runs = runs
        item.images = images
        item.item_id = hashlib.md5(
            f"{content}|para".encode()).hexdigest()[:16]

        # 段落属性（合并样式定义）
        item.para_props = self._extract_para_props(pPr) if pPr is not None else {}
        # 补齐样式中的大纲级别等属性
        if 'outlineLvl' not in item.para_props and 'outlineLvl' in style_props:
            item.para_props['outlineLvl'] = style_props['outlineLvl']

        # 自动编号前缀
        if pPr is not None:
            numPr = pPr.find(_wtag('numPr'))
            if numPr is not None:
                try:
                    nid = numPr.find(_wtag('numId'))
                    ilv = numPr.find(_wtag('ilvl'))
                    if nid is not None and ilv is not None:
                        nid_val = nid.get(_wtag('val'), '')
                        ilv_val = ilv.get(_wtag('val'), '')
                        prefix = nr.resolve(int(nid_val), int(ilv_val))
                        if prefix and not content.startswith(prefix.strip()):
                            item.content = prefix + content
                except Exception:
                    pass

        # 初始级别（仅从样式获取）
        item.level = self._get_heading_level_from_style(item)
        item.detected_level = 0  # 默认正文
        item.page = 1

        return item

    def _extract_para_props(self, pPr: etree._Element) -> Dict:
        """Extract paragraph properties."""
        props = {}
        if pPr is None:
            return props

        def gi(tag):
            el = pPr.find(_wtag(tag))
            return int(el.get(_wtag('val'), 0)) if el is not None else None

        def gs(tag):
            el = pPr.find(_wtag(tag))
            return el.get(_wtag('val'), '') if el is not None else None

        props['ind_left'] = gi('indLeft')
        props['ind_right'] = gi('indRight')
        props['ind_first_line'] = gi('indFirstLine')
        props['spacing_before'] = gi('spacingBefore')
        props['spacing_after'] = gi('spacingAfter')
        props['outlineLvl'] = gs('outlineLvl')
        sl = gi('spacingLine')
        if sl:
            props['spacing_line'] = sl
        props['spacing_line_rule'] = gs('spacingLineRule') or 'auto'
        props['jc'] = gs('jc') or 'left'

        numPr = pPr.find(_wtag('numPr'))
        if numPr is not None:
            nid = numPr.find(_wtag('numId'))
            ilv = numPr.find(_wtag('ilvl'))
            if nid is not None:
                props['numId'] = nid.get(_wtag('val'))
            if ilv is not None:
                props['ilvl'] = ilv.get(_wtag('val'))

        return props

    # ------------------------------------------------------------------
    # 表格提取
    # ------------------------------------------------------------------
    def _extract_table_item(self, tbl: etree._Element, tbl_idx: int) -> DocItem:
        item = DocItem()
        item.index = -1
        item.type = 'table'
        item.xml = tbl

        rows = []
        for tr in tbl.findall(_wtag('tr')):
            row = []
            for tc in tr.findall(_wtag('tc')):
                tcPr = tc.find(_wtag('tcPr'))
                vMerge = None
                hMerge = 1
                if tcPr is not None:
                    vm = tcPr.find(_wtag('vMerge'))
                    if vm is not None:
                        vMerge = vm.get(_wtag('val'), 'continue')
                    hm = tcPr.find(_wtag('gridSpan'))
                    if hm is not None:
                        hMerge = int(hm.get(_wtag('val'), 1))
                cell_text = self._get_cell_text(tc)
                row.append({
                    'text': cell_text,
                    'rowspan': vMerge,
                    'colspan': hMerge,
                })
            rows.append(row)

        content = ' | '.join('  '.join(c['text'] for c in row) for row in rows)
        item.content = content
        item.rows = rows
        item.item_id = hashlib.md5(f"{content}|table".encode()).hexdigest()[:16]

        # 表格属性
        tblPr = tbl.find(_wtag('tblPr'))
        if tblPr is not None:
            tblW = tblPr.find(_wtag('tblW'))
            if tblW is not None:
                item.table_props['width'] = tblW.get(_wtag('val'), '')
                item.table_props['width_type'] = tblW.get(_wtag('type'), 'dxa')

        return item

    def _get_cell_text(self, tc: etree._Element) -> str:
        parts = []
        for t in tc.iter(_wtag('t')):
            if t.text:
                parts.append(t.text)
        return ''.join(parts).strip()

    # ------------------------------------------------------------------
    # SDT 提取
    # ------------------------------------------------------------------
    def _extract_sdt_item(self, sdt: etree._Element, idx: int) -> DocItem:
        item = DocItem()
        item.index = idx
        item.type = 'sdt'
        tag_el = sdt.find(f'.//{_wtag("tag")}')
        item.item_id = tag_el.get(_wtag('val'), '') if tag_el is not None else f'sdt_{idx}'
        item.xml = sdt
        # 取 SDT 内第一个段落文字
        first_p = sdt.find(_wtag('p'))
        if first_p is not None:
            texts = [t.text for t in first_p.iter(_wtag('t')) if t.text]
            item.content = ''.join(texts).strip()
        item.item_id = hashlib.md5(f"{item.content}|sdt".encode()).hexdigest()[:16]
        return item

    def _extract_misc_item(self, el: etree._Element, idx: int) -> DocItem:
        """提取非 p/tbl/sdt 的 body 子元素（如 sectPr、bookmarkEnd 等）"""
        item = DocItem()
        item.index = idx
        item.xml = el
        item.body_tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag

        tag = item.body_tag
        if tag == 'sectPr':
            item.type = 'section'
            item.content = '[节属性]'
            # 从 sectPr 中提取页面布局信息作为 item_id
            pgSz = el.find(_wtag('pgSz'))
            if pgSz is not None:
                w = pgSz.get(_wtag('w'), '')
                h = pgSz.get(_wtag('h'), '')
                item.item_id = f"sectPr_{w}x{h}"
            else:
                item.item_id = 'sectPr'
        elif tag in ('bookmarkStart', 'bookmarkEnd'):
            item.type = 'bookmark'
            name_el = el.find(_wtag('name'))
            name = name_el.get(_wtag('val'), '') if name_el is not None else ''
            item.item_id = name or f'bookmark_{idx}'
            item.content = f'[书签: {name}]'
        elif tag in ('commentRangeStart', 'commentRangeEnd', 'commentExRangeStart', 'commentExRangeEnd'):
            item.type = 'comment_marker'
            item.content = f'[{tag}]'
            item.item_id = f'{tag}_{idx}'
        elif tag == 'hyperlink':
            item.type = 'hyperlink'
            texts = [t.text for t in el.iter(_wtag('t')) if t.text]
            item.content = ''.join(texts).strip()
            item.item_id = f'hyperlink_{idx}'
        elif tag == 'ins' or tag == 'del':
            item.type = tag  # 'ins' 或 'del'
            texts = [t.text for t in el.iter(_wtag('t')) if t.text]
            item.content = ''.join(texts).strip()
            item.item_id = f'{tag}_{idx}'
        elif tag == 'customXml':
            item.type = 'custom_xml'
            item.content = '[自定义XML]'
            item.item_id = f'customXml_{idx}'
        else:
            # 未知类型，保留原始标签名
            item.type = tag
            item.content = f'[{tag}]'
            item.item_id = f'{tag}_{idx}'

        return item

    # ------------------------------------------------------------------
    # V2 合并核心：items 直接映射到 XML，无 blocks 层级
    # ------------------------------------------------------------------
    # 辅助方法
    # ------------------------------------------------------------------
    def merge_documents_with_revisions(
        self, base_doc: str, copies: List[Tuple[str, str]],
        output: str, changes: List[DocumentChange] = None,
        original_para_by_order: Dict = None,
        original_para_anchor_map: Dict = None,
        conflicting_changes_by_para: Dict[int, List] = None
    ) -> str:
        """
        纯 XML 合并：复制原稿 → 解压 → 用 lxml 修改 XML → 重新打包
        """
        print(f"    [合并] 使用纯 XML 引擎合并...")
        self._replaced_positions = set()
        self._move_revision_ids = {}

        if os.path.exists(output):
            input_paths = [base_doc, *(path for path, _ in copies)]
            if any(os.path.samefile(path, output) for path in input_paths):
                raise ValueError(f"输出文件不能覆盖输入文档: {output}")

        output_dir = os.path.dirname(output) or '.'
        os.makedirs(output_dir, exist_ok=True)
        conflict_ids = {
            change.change_id
            for group in (conflicting_changes_by_para or {}).values()
            for change in group
        }

        # 建立副本颜色映射（使用 copy_index 保证唯一性）
        copy_colors: Dict[int, str] = {}
        if changes:
            # 💡 核心改进：所有颜色分配均基于副本索引，而非作者名
            uniq_indices = sorted(set(c.copy_index for c in changes))
            for i, idx in enumerate(uniq_indices):
                copy_colors[idx] = _AUTHOR_COLORS[i % len(_AUTHOR_COLORS)]

        # changes 已在 _auto_merge 中按 [删除逆序, 修改逆序, 新增逆序] 排序
        # 这里直接使用，不再重新分类排序

        print(f"    [合并] 变更统计: 删除/修改/新增 共 {len(changes) if changes else 0} 条")

        # 解压 docx
        extract_dir = tempfile.mkdtemp(prefix='docx_merge_')
        try:
            with zipfile.ZipFile(base_doc, 'r') as zf:
                zf.extractall(extract_dir)

            doc_xml_path = os.path.join(extract_dir, 'word', 'document.xml')

            # 解析 document.xml
            parser = etree.XMLParser(remove_blank_text=False, ns_clean=False, recover=True)
            doc_tree = etree.parse(doc_xml_path, parser)
            doc_root = doc_tree.getroot()
            body = doc_root.find(_wtag('body'))
            revision_ids = (
                int(value) for node in body.iter()
                if node.tag in _REVISION_MARKUP
                for value in [node.get(_wtag('id'), '')] if value.isdecimal()
            )
            self._next_revision_id = max(revision_ids, default=-1) + 1
            drawing_ids = (int(node.get('id')) for node in body.iter(_wptag('docPr'))
                           if node.get('id', '').isdecimal())
            self._next_drawing_id = max(drawing_ids, default=0) + 1

            # 解析现有评论 ID 确保唯一性
            comments_path = os.path.join(extract_dir, 'word', 'comments.xml')
            global_next_cid = 0
            if os.path.exists(comments_path):
                try:
                    existing_c_tree = etree.parse(comments_path)
                    existing_c_ids = [int(el.get(_wtag('id'), 0)) for el in existing_c_tree.getroot().findall('.//' + _wtag('comment'))]
                    if existing_c_ids:
                        global_next_cid = max(existing_c_ids) + 1
                except Exception:
                    pass

            # 💡 核心改进：按 (副本索引, 作者) 存储批注，确保颜色一致
            self.copy_comment_ids = {} # {copy_idx: [(cid, author, text)]}
            def get_comment_id(copy_idx: int, author: str, comment_text: str = '') -> str:
                nonlocal global_next_cid
                cid = str(global_next_cid)
                global_next_cid += 1
                if copy_idx not in self.copy_comment_ids:
                    self.copy_comment_ids[copy_idx] = []
                # 记录副本索引及其对应的作者信息
                self.copy_comment_ids[copy_idx].append((cid, author, comment_text))
                return cid

            # 建立 body 子元素列表（仅块级），每次操作后重新构建以保持同步
            def rebuild_blocks():
                block_tags = {_wtag('p'), _wtag('tbl'), _wtag('sdt')}
                return [c for c in body if c.tag in block_tags]

            blocks = rebuild_blocks()
            block_by_idx: Dict[int, etree._Element] = {}
            for i, b in enumerate(blocks):
                block_by_idx[i] = b

            # 建立 body.children 索引缓存（避免重复 list(body) 调用）
            def get_body_children_list():
                return list(body)

            # ── 统一逆序处理所有变更 ───────────────────────────────────────
            if changes:
                # 预处理：标记连续变更以减少重复生成批注气泡
                # 注意：只在完全相同位置时才跳过批注，保留每个不同位置变更的批注
                last_pos_by_copy = {}
                for change in changes:
                    change._skip_comment = False
                    pos = getattr(change, 'true_block_index', change.position)
                    ctype = change.change_type

                    last_info = last_pos_by_copy.get(change.copy_index)
                    if last_info:
                        last_pos, last_ctype = last_info
                        # 只有在同一位置、同一类型时，才跳过批注
                        # 这样每个不同的修改位置都会生成批注
                        if pos == last_pos and ctype == last_ctype:
                            change._skip_comment = True

                    last_pos_by_copy[change.copy_index] = (pos, ctype)

                for change in changes:
                    # 优先使用 true_block_index（已在 _write_changes 中计算），否则用 position
                    pos = getattr(change, 'true_block_index', change.position)
                    ctype = change.change_type
                    author = change.author
                    copy_idx = change.copy_index
                    color = copy_colors.get(copy_idx, 'FF6B6B')
                    skip_comment = getattr(change, '_skip_comment', False)

                    # 不由于新增或删除而重建 blocks，保持块指针长期有效
                    copy_path = copies[copy_idx][0] if 0 <= copy_idx < len(copies) else None
                    is_conflict = change.change_id in conflict_ids

                    if ctype == 'delete' or ctype.endswith('_delete'):
                        self._apply_delete(body, blocks, pos, change, author, color, get_comment_id, skip_comment)
                        if is_conflict:
                            self._replaced_positions.add(pos)

                    elif ctype in ('modify', 'table', 'image'):
                        self._apply_modify(body, blocks, pos, change, author, color, get_comment_id,
                                           extract_dir, copy_path, skip_comment, is_conflict)

                    elif ctype == 'insert' or ctype.endswith('_insert'):
                        self._apply_insert(body, blocks, pos, change, author, color, get_comment_id, extract_dir, copy_path, skip_comment)

                    # 移除了块索引的灾难性刷新

            # 写回 document.xml
            doc_tree.write(doc_xml_path, xml_declaration=True,
                           encoding='UTF-8', standalone=True)

            # 更新 comments.xml（如有新增批注）
            if self.copy_comment_ids:
                self._update_comments_xml(extract_dir, self.copy_comment_ids, copy_colors)

            # 先在输出目录生成完整文件，再替换目标，避免留下半写入的合稿。
            with tempfile.NamedTemporaryFile(dir=output_dir, suffix='.tmp', delete=False) as staged:
                staged_path = staged.name
            try:
                self._repack_docx(extract_dir, staged_path)
                os.replace(staged_path, output)
            finally:
                if os.path.exists(staged_path):
                    os.unlink(staged_path)
            print(f"    [合并] 完成: {output}")

        finally:
            shutil.rmtree(extract_dir, ignore_errors=True)

        return output

    # ------------------------------------------------------------------
    # 统一写回辅助方法（逆序处理）
    # ------------------------------------------------------------------
    def _apply_delete(self, body: etree._Element,
                      blocks: List[etree._Element],
                      pos: int,
                      change: DocumentChange,
                      author: str, color: str,
                      get_cid_fn, skip_comment: bool = False) -> None:
        """删除指定位置的块（标记删除线 + 批注）"""
        if pos < 0 or pos >= len(blocks):
            print(f"      [DEL] 跳过: position={pos} 超出范围")
            return
        blk = blocks[pos]
        copy_idx = change.copy_index
        
        fmt = change.format_info or {}
        item_type = fmt.get('item_type', 'paragraph')
        
        cid = None
        if not skip_comment:
            if item_type == 'table':
                cid = get_cid_fn(copy_idx, author, f"[{author} 删除] 表格")
            elif item_type == 'image':
                cid = get_cid_fn(copy_idx, author, f"[{author} 删除] 图片")
            else:
                content_preview = change.content[:50]
                cid = get_cid_fn(copy_idx, author, f"[{author} 删除] 内容: {content_preview}")

        # 1. 挂载批注引用（取代难看的独立正文段落）
        if cid is not None:
            if item_type == 'table':
                first_p = blk.find('.//' + _wtag('p'))
                if first_p is not None:
                    self._add_comment_to_element(first_p, cid, author)
            else:
                self._add_comment_to_element(blk, cid, author)

        # 2. 标记删除线（带代理副本颜色）
        self._mark_deleted(body, blk, author, color)

        print(f"      [DEL] pos={pos} author={author} color={color}")

    def _apply_modify(self, body: etree._Element,
                       blocks: List[etree._Element],
                       pos: int,
                       change: DocumentChange,
                       author: str, color: str,
                       get_cid_fn, extract_dir: str = None, copy_path: str = None,
                       skip_comment: bool = False, is_conflict: bool = False) -> None:
        """修改/替换指定位置的块"""
        if pos < 0 or pos >= len(blocks):
            print(f"      [MOD] 跳过: position={pos} 超出范围")
            return
        blk = blocks[pos]
        ctype = change.change_type

        if ctype == 'table':
            self._apply_table_diff(body, blocks, pos, change, author, color, get_cid_fn,
                                   extract_dir, copy_path, skip_comment, is_conflict)
        else:
            # 段落修改
            self._apply_para_modify(body, blocks, pos, change, author, color, get_cid_fn, extract_dir, copy_path, skip_comment)

        print(f"      [MOD] pos={pos} type={ctype} author={author} color={color}")

    def _insert_delete_comment(self, body: etree._Element,
                                block: etree._Element,
                                author: str, color: str,
                                get_cid_fn, copy_idx: int) -> None:
        """在删除的块前插入一行带批注的说明段落"""
        # 不改变原块内容，仅在 body 中块之前插入一个批注说明行
        para_text = self._get_para_text(block)[:50]
        cid = get_cid_fn(copy_idx, author, f"[{author} 删除] 内容: {para_text}")
        comment_p = self._make_comment_para(
            f"[{author} 删除] 内容: {para_text}", color, cid)
        idx = list(body).index(block)
        body.insert(idx, comment_p)

    def _mark_deleted(self, body: etree._Element,
                       block: etree._Element,
                       author: str, color: str) -> None:
        """将块内容标记为删除（给所有文字加删除线并应用副本唯一颜色）"""
        for run in block.iter(_wtag('r')):
            rPr = run.find(_wtag('rPr'))
            if rPr is None:
                rPr = etree.SubElement(run, _wtag('rPr'))
            
            # 1. 删除线
            existing = rPr.find(_wtag('strike'))
            if existing is None:
                etree.SubElement(rPr, _wtag('strike'))
            
            # 2. 颜色统一（与批注一致）
            clr = rPr.find(_wtag('color'))
            if clr is None:
                clr = etree.SubElement(rPr, _wtag('color'))
            clr.set(_wtag('val'), color)

    def _assign_copy_local_ids(self, element, copy_index):
        """Assign fresh drawing and revision identifiers to imported markup."""
        for node in element.iter():
            if node.tag == _wptag('docPr') and node.get('id') is not None:
                node.set('id', str(self._next_drawing_id))
                self._next_drawing_id += 1

            old_id = node.get(_wtag('id'))
            if old_id is None or node.tag not in _REVISION_MARKUP:
                continue
            if node.tag in _MOVE_RANGE_TAGS:
                group = node.tag[:-5] if node.tag.endswith('Start') else node.tag[:-3]
                key = (copy_index, group, old_id)
                new_id = self._move_revision_ids.get(key)
                if new_id is not None:
                    node.set(_wtag('id'), new_id)
                    continue
                self._move_revision_ids[key] = str(self._next_revision_id)
            node.set(_wtag('id'), str(self._next_revision_id))
            self._next_revision_id += 1

    def _apply_para_modify(self, body: etree._Element,
                            blocks: List[etree._Element],
                            pos: int,
                            change: DocumentChange,
                            author: str, color: str,
                            get_cid_fn, extract_dir: str = None, copy_path: str = None, skip_comment: bool = False) -> None:
        """用 diff 片段更新段落内容，差异处匹配副本颜色"""
        if pos < 0 or pos >= len(blocks):
            return
        blk = blocks[pos]
        copy_idx = change.copy_index
        cid = None if skip_comment else get_cid_fn(copy_idx, author, f"[{author} 修改]")

        try:
            diff_data = json.loads(change.content)
        except Exception:
            diff_data = None

        if diff_data and isinstance(diff_data, dict) and (diff_data.get('diff_segments') is not None or diff_data.get('copy_para_xml')):
            # 优先使用副本段落 XML（直接标红，不重建 runs）
            copy_para_xml = diff_data.get('copy_para_xml', '')
            if copy_para_xml:
                old_content = diff_data.get('old_content', '')
                new_content = diff_data.get('new_content', '')
                new_para = self._mark_diff_runs_in_copy_para(
                    copy_para_xml, old_content, new_content, color)
                self._assign_copy_local_ids(new_para, change.copy_index)
                if extract_dir and copy_path:
                    self._clone_relationships_and_media(new_para, copy_path, extract_dir)
            else:
                # 备用：用 diff_segments 重建
                copy_pPr_xml = diff_data.get('copy_pPr_xml', '')
                new_para = self._rebuild_para_with_diff(blk, diff_data['diff_segments'], copy_pPr_xml, color)
        else:
            # 纯文本替换
            new_para = self._make_para(change.content)

        # 给新段落加批注
        if cid is not None:
            self._add_comment_to_element(new_para, cid, author)

        # # 💡 核心改进：根据推断级别修正段落属性（样式与纲目级别）
        # actual_level = change.format_info.get('actual_level', 0)
        # self._set_heading_attributes(new_para, actual_level)

        # 冲突并存：如果此块已被他人修改，不要删除，而是级联插入
        # 使用 pos 集合追踪被替换过的段落
        if not hasattr(self, '_replaced_positions'):
            self._replaced_positions = set()
            
        is_already_replaced = pos in self._replaced_positions
        blk_idx = list(body).index(blk)
        
        if is_already_replaced:
            # 插入并列保留提示
            warn_para = self._make_comment_para(
                f"[冲突警告: 此处存在多个副本的变更，{author} 的修改被并列保留]", "FF9900"
            )
            body.insert(blk_idx + 1, warn_para)
            body.insert(blk_idx + 2, new_para)
            blocks[pos] = new_para  # 将追踪指针移向最新的这段，供下一位继续跟随追加
        else:
            # 正常替换
            body.remove(blk)
            body.insert(blk_idx, new_para)
            self._replaced_positions.add(pos)
            blocks[pos] = new_para

    def _rebuild_para_with_diff(self, old_para: etree._Element,
                                 diff_segments: List[dict],
                                 copy_pPr_xml: str = '',
                                 color: str = 'FF0000') -> etree._Element:
        """根据 diff 片段重建段落，保留副本段落样式，差异内容使用作者颜色"""
        new_para = etree.Element(_wtag('p'))
        # 优先使用副本的段落属性（行间距等），其次用原稿的
        if copy_pPr_xml:
            try:
                copy_pPr = etree.fromstring(copy_pPr_xml)
                new_para.append(copy_pPr)
            except Exception:
                pPr_src = old_para.find(_wtag('pPr'))
                if pPr_src is not None:
                    new_para.append(etree.fromstring(etree.tostring(pPr_src)))
        else:
            pPr_src = old_para.find(_wtag('pPr'))
            if pPr_src is not None:
                new_para.append(etree.fromstring(etree.tostring(pPr_src)))

        # 合并连续相同文字的 segment（处理 SequenceMatcher 错位导致的重复）
        merged = []
        for seg in diff_segments:
            if merged and merged[-1]['text'] == seg['text'] and merged[-1]['is_diff'] == seg['is_diff']:
                continue  # 跳过重复的连续相同文字
            merged.append(seg)

        for seg in merged:
            run = etree.SubElement(new_para, _wtag('r'))
            rPr = etree.SubElement(run, _wtag('rPr'))

            attrs = seg.get('attrs', {})
            is_diff = seg.get('is_diff', False)

            fg = attrs.get('color', '000000')
            if is_diff:
                fg = color  # 使用作者颜色标记差异

            # 颜色
            color_el = etree.SubElement(rPr, _wtag('color'))
            color_el.set(_wtag('val'), fg)

            # 字体
            font = attrs.get('font', '')
            if font:
                rFonts = etree.SubElement(rPr, _wtag('rFonts'))
                rFonts.set(_wtag('eastAsia'), font)
                rFonts.set(_wtag('hAnsi'), font)

            # 字号
            fsz = attrs.get('fontSize', '')
            if fsz:
                sz_el = etree.SubElement(rPr, _wtag('sz'))
                sz_el.set(_wtag('val'), fsz)

            # 加粗
            if attrs.get('bold', False):
                etree.SubElement(rPr, _wtag('b'))

            # 斜体
            if attrs.get('italic', False):
                etree.SubElement(rPr, _wtag('i'))

            # 下划线
            ul = attrs.get('underline', '')
            if ul and ul not in ('', 'none', 'false'):
                ul_el = etree.SubElement(rPr, _wtag('u'))
                ul_el.set(_wtag('val'), ul if ul not in ('true', 'single') else 'single')

            t_el = etree.SubElement(run, _wtag('t'))
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t_el.text = seg.get('text', '')

        return new_para

    def _mark_diff_runs_in_copy_para(self, copy_para_xml: str,
                                      old_content: str,
                                      new_content: str,
                                      color: str = 'FF0000') -> etree._Element:
        """
        直接使用副本段落 XML，通过字符级 diff 计算差异区间，
        遍历每个 run，在有差异的 run 上加红色标记。
        这样副本段落的完整结构和样式（runs、数值表等）全部保留。
        """
        try:
            para_el = etree.fromstring(copy_para_xml)
        except Exception:
            # 解析失败，回退到纯文本
            para_el = etree.Element(_wtag('p'))
            run = etree.SubElement(para_el, _wtag('r'))
            t = etree.SubElement(run, _wtag('t'))
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = new_content

        if old_content == new_content or not old_content or not new_content:
            return para_el
        if any(
            node.tag in _REVISION_MARKUP or any(_relationship_attribute(name) for name in node.attrib)
            for node in para_el.iter()
        ):
            return para_el

        # 字符级 diff，找出哪些位置有差异
        diff_ranges = self._char_diff_ranges(old_content, new_content)

        # 收集所有 <w:r> 及其文字范围
        all_runs = para_el.iter(_wtag('r'))
        run_list = list(all_runs)
        if not run_list:
            return para_el

        # 将 run 的文字拼接起来，并记录每个 run 的起止位置
        run_texts = []
        for run in run_list:
            t_els = run.findall(_wtag('t'))
            if t_els:
                run_texts.append(''.join(t.text or '' for t in t_els))
            else:
                run_texts.append('')

        # 计算每个 run 在整段文字中的绝对偏移
        run_offsets = []
        offset = 0
        for rt in run_texts:
            run_offsets.append(offset)
            offset += len(rt)

        # 遍历 run，在 diff 边界处拆分，只对差异字符标色
        for i, (run, rt) in enumerate(zip(run_list, run_texts)):
            r_start = run_offsets[i]
            r_end = r_start + len(rt)
            if r_start >= r_end:
                continue

            # 收集此 run 内与 diff_ranges 有交集的区间
            overlaps = []
            for d_start, d_end in diff_ranges:
                # 取交集
                ov_start = max(d_start, r_start)
                ov_end = min(d_end, r_end)
                if ov_start < ov_end:
                    overlaps.append((ov_start - r_start, ov_end - r_start))  # 转为 run 内偏移

            if not overlaps or any(child.tag not in (_wtag('rPr'), _wtag('t')) for child in run):
                continue

            # 整个 run 都是差异 → 直接标色，无需拆分
            if len(overlaps) == 1 and overlaps[0][0] == 0 and overlaps[0][1] == len(rt):
                rPr = run.find(_wtag('rPr'))
                if rPr is None:
                    rPr = etree.Element(_wtag('rPr'))
                    run.insert(0, rPr)
                color_el = rPr.find(_wtag('color'))
                if color_el is None:
                    color_el = etree.SubElement(rPr, _wtag('color'))
                color_el.set(_wtag('val'), color)
                continue

            # 需要拆分：将当前 run 拆为多个子 run（原文保持 + 差异标色）
            # 构建分段列表：[(local_start, local_end, is_diff), ...]
            segments = []
            pos = 0
            for ov_s, ov_e in sorted(overlaps):
                if pos < ov_s:
                    segments.append((pos, ov_s, False))
                segments.append((ov_s, ov_e, True))
                pos = ov_e
            if pos < len(rt):
                segments.append((pos, len(rt), False))

            # 提取原 run 的 rPr 模板
            orig_rPr = run.find(_wtag('rPr'))
            rPr_xml = etree.tostring(orig_rPr) if orig_rPr is not None else None

            # 获取 run 在父元素中的位置
            parent = run.getparent()
            run_idx = list(parent).index(run)

            # 移除原 run
            parent.remove(run)

            # 在原位置插入拆分后的子 run
            for seg_idx, (seg_s, seg_e, is_diff) in enumerate(segments):
                sub_run = etree.Element(_wtag('r'))
                # 复制 rPr
                if rPr_xml is not None:
                    sub_rPr = etree.fromstring(rPr_xml)
                else:
                    sub_rPr = etree.SubElement(sub_run, _wtag('rPr'))

                if is_diff:
                    # 差异段：设置作者颜色
                    color_el = sub_rPr.find(_wtag('color'))
                    if color_el is None:
                        color_el = etree.SubElement(sub_rPr, _wtag('color'))
                    color_el.set(_wtag('val'), color)

                sub_run.insert(0, sub_rPr)

                # 写入文本
                t_el = etree.SubElement(sub_run, _wtag('t'))
                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t_el.text = rt[seg_s:seg_e]

                parent.insert(run_idx + seg_idx, sub_run)

        return para_el

    def _char_diff_ranges(self, old: str, new: str) -> List[tuple]:
        """
        基于字符级 SequenceMatcher，计算哪些字符位置有差异。
        返回 [(start, end), ...] 列表，坐标基于新字符串（副本）。
        """
        matcher = difflib.SequenceMatcher(None, old, new)
        ranges = []
        for tag, o_s, o_e, n_s, n_e in matcher.get_opcodes():
            if tag != 'equal':
                ranges.append((n_s, n_e))
        return ranges

    def _apply_table_diff(self, body: etree._Element,
                           blocks: List[etree._Element],
                           pos: int,
                           change: DocumentChange,
                           author: str, color: str,
                           get_cid_fn, extract_dir: str = None, copy_path: str = None,
                           skip_comment: bool = False, is_conflict: bool = False) -> None:
        """合并表格单元格变更，保留结构变更和冲突副本。"""
        if pos < 0 or pos >= len(blocks):
            return
        tbl_el = blocks[pos]
        if tbl_el.tag != _wtag('tbl'):
            return

        fmt = change.format_info or {}
        diff_cells = json.loads(change.content)
        copy_table_xml = fmt.get('copy_tbl_xml')
        if copy_table_xml and (fmt.get('structure_changed') or is_conflict):
            copy_table = etree.fromstring(copy_table_xml.encode('utf-8'))
            if extract_dir and copy_path:
                self._clone_relationships_and_media(copy_table, copy_path, extract_dir)

            label = '表格变更冲突，副本待复核' if is_conflict else '修改表格结构'
            first_p = copy_table.find('.//' + _wtag('p'))
            if first_p is not None and not skip_comment:
                cid = get_cid_fn(change.copy_index, author, f"[{author} {label}]")
                self._add_comment_to_element(first_p, cid, author)

            if is_conflict:
                tbl_el.addnext(copy_table)
            else:
                body.replace(tbl_el, copy_table)
                blocks[pos] = copy_table
            return

        for diff in diff_cells:
            r_idx = diff.get('row', 0)
            c_idx = diff.get('col', 0)
            new_text = diff.get('new', '')

            rows = tbl_el.findall(_wtag('tr'))
            if r_idx >= len(rows):
                continue
            row_el = rows[r_idx]
            cells = row_el.findall(_wtag('tc'))
            if c_idx >= len(cells):
                continue
            cell_el = cells[c_idx]

            # 更新单元格文字
            for para in cell_el.findall(_wtag('p')):
                for run in para.findall(_wtag('r')):
                    for t in run.findall(_wtag('t')):
                        t.text = ''
                # 找一个 run 或新建
                runs = [r for r in para.findall(_wtag('r')) if r.find(_wtag('t')) is not None or r.find(_wtag('delText')) is not None]
                if runs:
                    r = runs[0]
                else:
                    r = etree.SubElement(para, _wtag('r'))
                t_el = r.find(_wtag('t'))
                if t_el is None:
                    t_el = etree.SubElement(r, _wtag('t'))
                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t_el.text = new_text

                # 加红色
                rPr = r.find(_wtag('rPr'))
                if rPr is None:
                    rPr = etree.SubElement(r, _wtag('rPr'))
                clr = rPr.find(_wtag('color'))
                if clr is None:
                    clr = etree.SubElement(rPr, _wtag('color'))
                clr.set(_wtag('val'), color)  # 用作者颜色标

        # 不要在表格外插入提示段落，而是把批注加到表格最前面的格子
        if not skip_comment:
            copy_idx = change.copy_index
            comment_text = f"[{author} 修改表格] {len(diff_cells)} 处修改"
            tbl_cid = get_cid_fn(copy_idx, author, comment_text)
            first_p = tbl_el.find('.//' + _wtag('p'))
            if first_p is not None:
                self._add_comment_to_element(first_p, tbl_cid, author)

    def _clone_relationships_and_media(self, elem: etree._Element, copy_path: str, extract_dir: str):
        """
        复制 elem 内部引用的 rId 对应资源，并更新 elem 中的 rId 以匹配 target 的 .rels。
        同时更新相应的 [Content_Types].xml。
        """
        import re
        
        nodes_with_rids = []
        for node in elem.iter():
            for attr_name, attr_val in node.attrib.items():
                if _relationship_attribute(attr_name):
                    nodes_with_rids.append((node, attr_name, attr_val))

        if not nodes_with_rids:
            return

        try:
            with zipfile.ZipFile(copy_path, 'r') as source_zf:
                if 'word/_rels/document.xml.rels' not in source_zf.namelist():
                    raise ValueError("副本缺少正文的资源关系表")
                with source_zf.open('word/_rels/document.xml.rels') as f:
                    source_rels_root = etree.parse(f).getroot()

                with source_zf.open('[Content_Types].xml') as f:
                    source_ct_root = etree.parse(f).getroot()

                target_rels_path = os.path.join(extract_dir, 'word', '_rels', 'document.xml.rels')
                if os.path.exists(target_rels_path):
                    target_rels_tree = etree.parse(target_rels_path)
                else:
                    os.makedirs(os.path.dirname(target_rels_path), exist_ok=True)
                    target_rels_tree = etree.ElementTree(etree.Element(f'{{{REL}}}Relationships'))
                target_rels_root = target_rels_tree.getroot()

                target_ct_path = os.path.join(extract_dir, '[Content_Types].xml')
                target_ct_tree = etree.parse(target_ct_path)
                target_ct_root = target_ct_tree.getroot()
                
                r_tag = lambda n: f'{{{REL}}}{n}'
                ct_tag = lambda n: f'{{{CT_NS}}}{n}'
                source_overrides = {
                    part.get('PartName'): part.get('ContentType')
                    for part in source_ct_root.findall(ct_tag('Override'))
                }
                source_defaults = {
                    part.get('Extension', '').lower(): part.get('ContentType')
                    for part in source_ct_root.findall(ct_tag('Default'))
                }

                # 💡 核心改进：在多副本场景下，由于多次打开/写入 rels 文件，
                # 必须确保基于当前文件内容重新扫描 ID，防止多副本 ID 竞争冲突。
                existing_ids = [el.get('Id', '') for el in target_rels_root]
                max_n = 0
                for rid in existing_ids:
                    # 匹配数字 ID (rId1, rId2...)
                    m = re.findall(r'\d+', rid)
                    if m:
                        try:
                            max_n = max(max_n, int(m[-1]))
                        except: pass
                
                next_rid_num = max_n + 1
                rid_map = {}

                for node, attr_name, old_rid in nodes_with_rids:
                    if old_rid in rid_map:
                        node.set(attr_name, rid_map[old_rid])
                        continue

                    source_rel = None
                    for rel in source_rels_root:
                        if rel.get('Id') == old_rid:
                            source_rel = rel
                            break
                    
                    if source_rel is None:
                        raise ValueError(f"副本缺少资源关系: {old_rid}")
                        
                    target_path_relative = source_rel.get('Target')
                    rel_type = source_rel.get('Type')
                    target_mode = source_rel.get('TargetMode')

                    new_rid = f"rId{next_rid_num}"
                    next_rid_num += 1
                    
                    new_rel = etree.SubElement(target_rels_root, r_tag('Relationship'))
                    new_rel.set('Id', new_rid)
                    new_rel.set('Type', rel_type)
                    if target_mode:
                        new_rel.set('TargetMode', target_mode)
                    
                    if target_mode == 'External':
                        new_rel.set('Target', target_path_relative)
                    else:
                        source_part = posixpath.normpath(posixpath.join('/word', target_path_relative))
                        source_item_path = source_part.lstrip('/')
                        dirname, filename = posixpath.split(source_part)
                        new_part = posixpath.join(dirname, f"{new_rid}_{filename}")
                        dest_path = os.path.abspath(os.path.join(extract_dir, *new_part.lstrip('/').split('/')))
                        if os.path.commonpath((extract_dir, dest_path)) != extract_dir:
                            raise ValueError("DOCX 资源路径超出合稿目录")
                        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                        new_rel.set('Target', posixpath.relpath(new_part, '/word'))
                        
                        with source_zf.open(source_item_path) as source, open(dest_path, 'wb') as target:
                            shutil.copyfileobj(source, target)

                        extension = os.path.splitext(filename)[1].lstrip('.').lower()
                        content_type = source_overrides.get(f"/{source_item_path}") or source_defaults.get(extension)
                        if not content_type:
                            raise ValueError(f"资源缺少内容类型: {source_item_path}")
                        etree.SubElement(target_ct_root, ct_tag('Override'),
                                         PartName=new_part, ContentType=content_type)

                    rid_map[old_rid] = new_rid
                    node.set(attr_name, new_rid)
                
                if rid_map:
                    target_rels_tree.write(target_rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)
                if rid_map:
                    target_ct_tree.write(target_ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)
                     
        except Exception as e:
            raise ValueError(f"无法从 {copy_path} 复制内嵌资源: {e}") from e

    def _ensure_table_borders(self, tbl: etree._Element) -> None:
        """确保表格具有完整的实线边框（单边 0.5pt 黑色）

        OOXML 表格边框结构：
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          </w:tblBorders>
        </w:tblPr>
        """
        # 查找或创建 tblPr
        tblPr = tbl.find(_wtag('tblPr'))
        if tblPr is None:
            tblPr = etree.Element(_wtag('tblPr'))
            tbl.insert(0, tblPr)

        # 查找或创建 tblBorders
        tblBorders = tblPr.find(_wtag('tblBorders'))
        if tblBorders is None:
            tblBorders = etree.SubElement(tblPr, _wtag('tblBorders'))

        # 边框定义：外边框4个 + 内部2个
        border_defs = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']

        for border_type in border_defs:
            existing = tblBorders.find(_wtag(border_type))
            if existing is None:
                new_border = etree.SubElement(tblBorders, _wtag(border_type))
                new_border.set(_wtag('val'), 'single')   # 实线
                new_border.set(_wtag('sz'), '4')       # 0.5pt = 4 half-points
                new_border.set(_wtag('space'), '0')
                new_border.set(_wtag('color'), 'auto')  # 自动黑色

    def _apply_insert(self, body: etree._Element,
                       blocks: List[etree._Element],
                       pos: int,
                       change: DocumentChange,
                       author: str, color: str,
                       get_cid_fn, extract_dir: str = None, copy_path: str = None, skip_comment: bool = False) -> None:
        """在指定位置插入新块（全文标色 + 批注）"""
        fmt = change.format_info or {}
        item_type = fmt.get('item_type', 'paragraph')
        copy_idx = change.copy_index
        
        cid = None
        if not skip_comment:
            if item_type == 'table':
                comment_text = f"[{author} 新增] 表格"
                cid = get_cid_fn(copy_idx, author, comment_text)
            elif item_type == 'image':
                comment_text = f"[{author} 新增] 图片"
                cid = get_cid_fn(copy_idx, author, comment_text)
            else:
                content_preview = change.content[:50]
                comment_text = f"[{author} 新增] 内容: {content_preview}"
                cid = get_cid_fn(copy_idx, author, comment_text)

        xml_str = fmt.get('copy_xml') or fmt.get('copy_tbl_xml') or fmt.get('copy_img_xml')
        new_parsed_el = None
        
        if xml_str:
            new_parsed_el = etree.fromstring(xml_str.encode('utf-8'))
            self._assign_copy_local_ids(new_parsed_el, change.copy_index)
            if extract_dir and copy_path:
                self._clone_relationships_and_media(new_parsed_el, copy_path, extract_dir)

        if new_parsed_el is not None and new_parsed_el.tag == _wtag('tbl'):
            new_tbl = new_parsed_el
            new_para = None
        elif new_parsed_el is not None:
            new_tbl = None
            new_para = new_parsed_el
        else:
            # 纯文本 fallback（针对未携带 XML或提取失败的）
            new_tbl = None
            if item_type == 'paragraph':
                text = change.content if change.content else ""
                new_para = self._make_para(text)
            else:
                new_para = self._make_para(f"[{item_type} 插入]")

        if new_para is not None:
            if cid is not None:
                self._add_comment_to_element(new_para, cid, author)
            
            # # 💡 核心改进：根据推断级别修正段落属性（样式与纲目级别）
            # actual_level = fmt.get('actual_level', 0)
            # self._set_heading_attributes(new_para, actual_level)

        # 不再单独插入外部的批注文字段落，直接操作元素本身
        if item_type == 'table' and new_tbl is not None:
            # 确保表格有默认边框（新增表格可能没有边框定义）
            self._ensure_table_borders(new_tbl)

            # 标红所有内部文字
            for r in new_tbl.findall('.//' + _wtag('r')):
                rPr = r.find(_wtag('rPr'))
                if rPr is None:
                    rPr = etree.Element(_wtag('rPr'))
                    r.insert(0, rPr)
                clr = rPr.find(_wtag('color'))
                if clr is None:
                    clr = etree.SubElement(rPr, _wtag('color'))
                clr.set(_wtag('val'), color)
                
            # 把批注挂载到表格的第一个段落，避免割裂上下文
            if cid is not None:
                first_p = new_tbl.find('.//' + _wtag('p'))
                if first_p is not None:
                    self._add_comment_to_element(first_p, cid, author)
            
            if 0 <= pos <= len(blocks):
                insert_idx = list(body).index(blocks[pos]) if pos < len(blocks) else len(body)
            else:
                insert_idx = len(body)
            body.insert(insert_idx, new_tbl)
        else:
            # 段落或其他元素：将其文字标红，然后直接插入
            if new_para is not None:
                for r in new_para.findall('.//' + _wtag('r')):
                    rPr = r.find(_wtag('rPr'))
                    if rPr is None:
                        rPr = etree.Element(_wtag('rPr'))
                        r.insert(0, rPr)
                    clr = rPr.find(_wtag('color'))
                    if clr is None:
                        clr = etree.SubElement(rPr, _wtag('color'))
                    clr.set(_wtag('val'), color)
                
                if 0 <= pos <= len(blocks):
                    insert_idx = list(body).index(blocks[pos]) if pos < len(blocks) else len(body)
                else:
                    insert_idx = len(body)
                body.insert(insert_idx, new_para)

    # def _set_heading_attributes(self, p_el: etree._Element, level: int) -> None:
    #     """为段落设置标题属性（样式 ID 和纲目级别）
    #     
    #     Args:
    #         p_el:  <w:p> 元素
    #         level: 1-基于的级别 (1, 2, 3...)
    #     """
    #     if level <= 0:
    #         return
    #         
    #     pPr = p_el.find(_wtag('pPr'))
    #     if pPr is None:
    #         pPr = etree.Element(_wtag('pPr'))
    #         p_el.insert(0, pPr)
    #         
    #     # 1. 设置样式 (w:pStyle)
    #     # 在 OOXML 中，标准标题 1-9 的 StyleId 通常就是 '1', '2', '3'...'9'
    #     pStyle = pPr.find(_wtag('pStyle'))
    #     if pStyle is None:
    #         # 插入到 pPr 的最前面
    #         pStyle = etree.Element(_wtag('pStyle'))
    #         pPr.insert(0, pStyle)
    #     pStyle.set(_wtag('val'), str(level))
    #     
    #     # 2. 设置纲目级别 (w:outlineLvl)
    #     # 0 对应 1级标题，1 对应 2级标题...
    #     outline = pPr.find(_wtag('outlineLvl'))
    #     if outline is None:
    #         outline = etree.SubElement(pPr, _wtag('outlineLvl'))
    #     outline.set(_wtag('val'), str(level - 1))

    def _make_para(self, text: str) -> etree._Element:
        para = etree.Element(_wtag('p'))
        run = etree.SubElement(para, _wtag('r'))
        t = etree.SubElement(run, _wtag('t'))
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        return para

    def _make_comment_para(self, text: str, color: str,
                            comment_id: str = None) -> etree._Element:
        """创建说明段落，按需附加批注引用。"""
        para = etree.Element(_wtag('p'))

        r = etree.SubElement(para, _wtag('r'))
        rPr2 = etree.SubElement(r, _wtag('rPr'))
        etree.SubElement(rPr2, _wtag('b'))
        color_el = etree.SubElement(rPr2, _wtag('color'))
        color_el.set(_wtag('val'), color)
        t = etree.SubElement(r, _wtag('t'))
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text

        if comment_id is not None:
            crr = etree.SubElement(para, _wtag('r'))
            cr = etree.SubElement(crr, _wtag('commentReference'))
            cr.set(_wtag('id'), str(comment_id))
        return para

    def _add_comment_to_element(self, elem: etree._Element,
                                 comment_id: str, author: str) -> None:
        """给元素末尾添加标准批注区间和引用"""
        target_p = elem if elem.tag == _wtag('p') else elem.find('.//' + _wtag('p'))
        if target_p is None:
            return
            
        # 必须给一个实体占位符，如果段落完全为空，Word 的解析器在处理游离批注区间时会直接爆炸
        has_text = any((t.text for t in target_p.iter(_wtag('t')) if t.text))
        if not has_text:
            dummy_r = etree.Element(_wtag('r'))
            dummy_t = etree.SubElement(dummy_r, _wtag('t'))
            dummy_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            dummy_t.text = " "
            target_p.append(dummy_r)

        rStart = etree.Element(_wtag('commentRangeStart'))
        rStart.set(_wtag('id'), str(comment_id))
        
        pPr = target_p.find(_wtag('pPr'))
        if pPr is not None:
            idx = list(target_p).index(pPr) + 1
            target_p.insert(idx, rStart)
        else:
            target_p.insert(0, rStart)
        
        rEnd = etree.SubElement(target_p, _wtag('commentRangeEnd'))
        rEnd.set(_wtag('id'), str(comment_id))
        
        run = etree.SubElement(target_p, _wtag('r'))
        cr = etree.SubElement(run, _wtag('commentReference'))
        cr.set(_wtag('id'), str(comment_id))

    def _get_para_text(self, para: etree._Element) -> str:
        parts = []
        for t in para.iter(_wtag('t')):
            if t.text:
                parts.append(t.text)
        return ''.join(parts)

    def _update_comments_xml(self, extract_dir: str,
                              copy_comment_ids: Dict[int, List[Tuple[str, str, str]]],
                              copy_colors: Dict[int, str]) -> None:
        """更新 word/comments.xml，追加批注，确保气泡文字与正文引用色值同步"""
        comments_path = os.path.join(extract_dir, 'word', 'comments.xml')
        if os.path.exists(comments_path):
            c_tree = etree.parse(comments_path)
            c_root = c_tree.getroot()
        else:
            nsmap = {
                'wpc': COMMENTS_NS,
                'w': WNS,
                None: WNS,
            }
            c_root = etree.Element(_wtag('comments'), nsmap=nsmap)

        for copy_idx, cids in copy_comment_ids.items():
            color = copy_colors.get(copy_idx, 'FF6B6B')
            for cid, author, comment_text in cids:
                cmt = etree.SubElement(c_root, _wtag('comment'))
                cmt.set(_wtag('id'), str(cid))
                # 💡 在作者名后标注副本编号，防止 Word UI 强行归类同一作者的颜色
                display_author = f"{author} [副本{int(copy_idx)+1}]"
                cmt.set(_wtag('author'), display_author)
                cmt.set(_wtag('date'), datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ'))
                cmt.set(_wtag('initials'), author[:2] if author else 'XX')

                p_el = etree.SubElement(cmt, _wtag('p'))
                pPr = etree.SubElement(p_el, _wtag('pPr'))
                rStyle = etree.SubElement(pPr, _wtag('pStyle'))
                rStyle.set(_wtag('val'), 'CommentText')

                # 引用（用于关联正文）：必须放到文字的前面，否则空白气泡
                r_ref = etree.SubElement(p_el, _wtag('r'))
                rPr_ref = etree.SubElement(r_ref, _wtag('rPr'))
                rStyle_ref = etree.SubElement(rPr_ref, _wtag('rStyle'))
                rStyle_ref.set(_wtag('val'), 'CommentReference')
                ref_el = etree.SubElement(r_ref, _wtag('annotationRef'))

                # 批注内容：作者 + 操作类型 + 描述
                if comment_text:
                    r_el = etree.SubElement(p_el, _wtag('r'))
                    # 💡 与副本正文颜色强一致
                    rPr_cmt = etree.SubElement(r_el, _wtag('rPr'))
                    clr_cmt = etree.SubElement(rPr_cmt, _wtag('color'))
                    clr_cmt.set(_wtag('val'), color)

                    t_el = etree.SubElement(r_el, _wtag('t'))
                    t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    t_el.text = comment_text

        c_tree2 = etree.ElementTree(c_root)
        c_tree2.write(comments_path, xml_declaration=True,
                      encoding='UTF-8', standalone=True)

        # 更新 [Content_Types].xml 和 .rels
        self._ensure_comments_content_type(extract_dir)

    def _ensure_comments_content_type(self, extract_dir: str) -> None:
        """确保 Content_Types.xml 包含 comments 类型"""
        ct_path = os.path.join(extract_dir, '[Content_Types].xml')
        if os.path.exists(ct_path):
            ct_tree = etree.parse(ct_path)
            ct_root = ct_tree.getroot()
            ct_tag = lambda n: f'{{{CT_NS}}}{n}'
            overrides = ct_root.findall(ct_tag('Override'))
            has_comments = any(
                o.get('PartName', '').lower() == '/word/comments.xml'
                for o in overrides)
            if not has_comments:
                ov = etree.SubElement(ct_root, ct_tag('Override'))
                ov.set('PartName', '/word/comments.xml')
                ov.set('ContentType',
                       'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
            ct_tree2 = etree.ElementTree(ct_root)
            ct_tree2.write(ct_path, xml_declaration=True,
                           encoding='UTF-8', standalone=True)

        # 更新 word/_rels/document.xml.rels
        rels_path = os.path.join(extract_dir, 'word', '_rels', 'document.xml.rels')
        if os.path.exists(rels_path):
            r_tree = etree.parse(rels_path)
        else:
            os.makedirs(os.path.dirname(rels_path), exist_ok=True)
            r_tree = etree.ElementTree(etree.Element(f'{{{REL}}}Relationships'))
        r_root = r_tree.getroot()
        R_NS = REL
        r_tag = lambda n: f'{{{R_NS}}}{n}'
        existing_ids = [el.get('Id', '') for el in r_root]
        next_rid = max(
            (int(i[3:]) for i in existing_ids if i.startswith('rId') and i[3:].isdigit()),
            default=1) + 1

        has_comments = any(
            el.get('Type', '').endswith('comments')
            for el in r_root)
        if not has_comments:
            rel = etree.SubElement(r_root, r_tag('Relationship'))
            rel.set('Id', f'rId{next_rid}')
            rel.set('Type',
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
            rel.set('Target', 'comments.xml')
            r_tree2 = etree.ElementTree(r_root)
            r_tree2.write(rels_path, xml_declaration=True,
                          encoding='UTF-8', standalone=True)

    def _repack_docx(self, extract_dir: str, output_path: str) -> None:
        """将解压后的目录重新打包为 docx"""
        with zipfile.ZipFile(output_path, 'w',
                             compression=zipfile.ZIP_DEFLATED) as zout:
            for dirpath, _, filenames in os.walk(extract_dir):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    arcname = os.path.relpath(filepath, extract_dir)
                    zout.write(filepath, arcname)

    # ── 标题判定与级别推断（供 extract_document_structure 和外部调用） ─────

    def _is_legitimate_heading(self, item, custom_headings: list = None) -> bool:
        """
        判断一个 item 是否是合法的真实标题（评分矩阵）。
        如果传递了 custom_headings，将精准依照其模板正则表达式判断。
        """
        import re
        text = (item.content or '').strip()
        if not text:
            return False

        if custom_headings:
            # 即使是自定义标题，如果段落过长（如 > 120），通常也是摘要或正文提及，而非标题锚点本身
            if len(text) > 120:
                return False
                
            for fmt in custom_headings:
                raw_fmt = fmt.strip()
                # 1. 修正：先转义原始模板中的特殊正则字符（如 . ( ) [ ] 等）
                pattern = re.escape(raw_fmt)
                # 2. 将被转义后的数字片段替换为 \d+
                pattern = re.sub(r'\d+', r'\\d+', pattern)
                # 3. 将中文数字替换为通配模式（支持中英混写）
                pattern = re.sub(r'[一二三四五六七八九十百千]+', r'[一二三四五六七八九十百千\\d]+', pattern)
                
                # 💡 增加标点容错逻辑：
                # 如果模板是以常见标点结尾，允许匹配其他常见标点（解决 1. vs 1、 的个人录入笔误）
                if re.search(r'[\\.\s、:：\)）]$', pattern):
                     # 将末尾的一个或多个标点位替换成正则字符集（点、空格、顿号、冒号、括号）
                     pattern = re.sub(r'[\\\.\s、:：\)）]+$', r'[.\\s、:：\\)）]?', pattern)

                try:
                    if re.match(r'^' + pattern, text):
                        return True
                except re.error:
                    pass
                if text.startswith(raw_fmt):
                    return True
            return False

        style = (item.style_name or '').lower()



        has_heading_style = 'heading' in style or '标题' in style

        # if not has_heading_style:
        #     return False

        if any(text.startswith(p) for p in ['图', '表', 'figure', 'table', 'fig.', '参考文献', 'abstract', '摘要']):
            return False

        # 权重矩阵：前缀(0.5) > 长度(0.2) > 字体格式(0.2) > 样式(0.1)
        prefix_score = 0.0
        
        if re.match(r'^第[一二三四五六七八九十百千\d]+[章节篇节]', text):
            prefix_score = 0.5
        elif re.match(r'^\d{1,3}(?:\.\d{1,3})+[\s\.、]?', text):
            prefix_score = 0.5
        elif re.match(r'^\d+[\s、．.．]', text):
            prefix_score = 0.5
        elif re.match(r'^[一二三四五六七八九十百]+[\s、．.．]', text):
            prefix_score = 0.5
        elif re.match(r'^[（(][一二三四五六七八九十百\d]+[)）]', text):
            prefix_score = 0.5
        elif re.match(r'^[IVXLC]+[\s、．.．]', text):
            prefix_score = 0.5

        # 长文本黑名单
        if len(text) > 60:
            return False

        length_score = 0.2 if len(text) <= 20 else 0.15 if len(text) <= 40 else 0

        # 字体格式评分：检查 run 级别的加粗、标题字体、大字号
        font_score = 0.0
        runs = getattr(item, 'runs', []) or []
        if runs:
            total_chars = sum(len(r.get('text', '')) for r in runs)
            if total_chars > 0:
                # 加粗占比
                bold_chars = sum(len(r.get('text', '')) for r in runs if r.get('bold'))
                bold_ratio = bold_chars / total_chars
                
                # 标题字体检测（黑体系列、Arial Black 等常见标题字体）
                heading_fonts = {'黑体', '微软雅黑', '华文中宋', '华文黑体', '方正小标宋',
                                 '方正黑体', '仿宋_GB2312', 'SimHei', 'Microsoft YaHei',
                                 'Arial Black', 'Impact', 'STHeiti', 'STXihei'}
                heading_font_chars = 0
                for r in runs:
                    font_name = (r.get('font', '') or '').strip()
                    if any(hf.lower() in font_name.lower() for hf in heading_fonts):
                        heading_font_chars += len(r.get('text', ''))
                heading_font_ratio = heading_font_chars / total_chars
                
                # 大字号检测（动态参考基准线）建议基准线至少加4个半点(2pt)
                baseline = getattr(self, '_body_baseline', 21.0)
                large_size_chars = 0
                for r in runs:
                    try:
                        sz = int(r.get('fontSize', '0') or '0')
                        if sz >= baseline + 4: 
                            large_size_chars += len(r.get('text', ''))
                    except: pass
                large_size_ratio = large_size_chars / total_chars

                if bold_ratio >= 0.8:
                    font_score += 0.10
                if heading_font_ratio >= 0.5:
                    font_score += 0.05
                if large_size_ratio >= 0.5:
                    font_score += 0.05
                # 上限 0.2
                font_score = min(font_score, 0.2)

        # 样式评分 (0.1) - 现在有了 styles.xml 解析，这里会更准
        style_score = 0.0
        if has_heading_style:
            style_score = 0.1
        
        # [新增] 确定性评分：如果 Word 官方标记了大纲级别，直接大幅加分
        outline_lvl = item.para_props.get('outlineLvl')
        if outline_lvl is not None:
            # 只要有 outlineLvl 且不是正文级（通常标题是 0-8），就认为是标题候选
            total_score = 1.0  # 官方认证，直接满分通行
        else:
            total_score = prefix_score + length_score + font_score + style_score

        # 将分数存储到 item 上，供后续打印展示
        item.heading_score = round(total_score, 2)
        return total_score >= 0.70

    def _get_heading_level_from_style(self, item) -> str:
        """从样式名中提取标题级别，增强识别能力"""
        style = (item.style_name or '').strip().lower()
        
        # 1. 直接数字样式 (如 <w:pStyle w:val="1" />)
        if style.isdigit() and len(style) == 1 and '1' <= style <= '9':
            return style

        # 2. 匹配 Heading 1, 标题 1, TOC 1 等
        m = re.search(r'(?:heading|标题|toc)\s*([1-9])', style, re.IGNORECASE)
        if m:
            return m.group(1)



        return '0'
 
    def _compute_body_font_baseline(self, items: List[DocItem]) -> float:
        """统计所有段落中最常出现的字体大小作为正文基准（按字符权重）"""
        size_weights = {}
        for item in items:
            if item.type != 'paragraph':
                continue
            for run in item.runs:
                sz = run.get('fontSize')
                if sz:
                    text_len = len(run.get('text', ''))
                    size_weights[sz] = size_weights.get(sz, 0) + text_len
        
        if not size_weights:
            return 21.0
            
        # 找权重最大的字号
        base_sz = max(size_weights, key=lambda k: size_weights[k])
        try:
            return float(base_sz)
        except:
            return 21.0



class MultiCopyMerger:
    """多副本文档合并器 - 累积式三路合并"""
    
    def __init__(self, original_doc: str,
                 diff_mode: str = 'anchor',
                 anchor_threshold: float = 0.9,
                 modify_threshold: float = 0.2,
                 custom_headings: list = None,
                 auto_fix_heading_level: list = None,
                 max_heading_level: int = 0):
        """
        初始化合并器

        Args:
            original_doc:       原始文档路径
            diff_mode:         Diff 算法模式 ('anchor' | 'direct')
            anchor_threshold:  标题锁点最低相似度 [0,1]，默认 0.9
            modify_threshold:  段落判修改/新增的最低相似度 [0,1]，默认 0.2
            custom_headings:   自定义段落标题锚点格式，如果提供则完全依据此进行锚定对齐
            auto_fix_heading_level: 自动修正标题级别的格式列表，例如 ['1.1', '（1）', '1)']
                                   根据列表中的顺序推断正确的级别层次（从第2级开始）
        """
        self.original_doc = os.path.abspath(original_doc)
        self.diff_mode = diff_mode
        self.anchor_threshold = max(0.0, min(1.0, anchor_threshold))
        self.modify_threshold = max(0.0, min(1.0, modify_threshold))
        self.custom_headings = custom_headings
        self.auto_fix_heading_level = auto_fix_heading_level
        self.max_heading_level = max_heading_level
        self.builder = OpenXMLDocumentEngine()

        self.copies: List[Tuple[str, str]] = []  # [(路径, 作者)]
        self.changes_by_author: Dict[str, List[DocumentChange]] = {}
        self.conflicts: List[MergeConflict] = []

        # 中间结果缓存（供各步骤复用）
        self.orig_items: List = []           # 原稿 items
        self.orig_anchors: List = []        # 原稿锚点
        self.copy_items_list: List[List] = []  # 各副本 items 列表
        self.anchor_pairs_list: List[List] = []  # 各副本锚点配对
        self.orphan_orig_list: List[List] = []   # 各副本孤立原稿锚点
        self.orphan_copy_list: List[List] = []   # 各副本孤立副本锚点

        # 根据 auto_fix_heading_level 构建级别映射表
        # 格式列表的顺序对应从第2级开始的级别
        self._heading_level_map: Dict[str, int] = {}  # 格式正则 -> 正确的级别
        if self.auto_fix_heading_level:
            self._build_level_mapping()

        # 提取原始文档结构（新的统一 items 格式）
        print(f"正在分析原始文档: {os.path.basename(original_doc)}...")
        self.original_structure = self.builder.extract_document_structure(original_doc)

        # === 统一锚点基准：预先构建原稿 items 的唯一标识映射 ===
        # 用于在多个副本的变更写回前进行一致性验证
        self._original_item_anchor_map: Dict[str, int] = {}  # anchor_hash -> item index
        self._original_item_by_index: Dict[int, object] = {}   # index -> DocItem

        for item in self.original_structure.get('items', []):
            item_index = item.index
            text = item.content[:50].strip()
            anchor_key = f"i{item_index}:{text}"  # 全局 index + 文本前50字作为唯一键
            self._original_item_anchor_map[anchor_key] = item_index
            self._original_item_by_index[item_index] = item

        # 兼容旧字段（用于遗留代码）
        self._original_para_anchor_map = self._original_item_anchor_map
        self._original_para_by_order = self._original_item_by_index

    def _build_level_mapping(self):
        """
        根据 auto_fix_heading_level 格式列表构建级别映射表。

        格式列表中的每个格式按顺序对应从第2级开始的标题级别。
        例如：['1.1', '（1）', '1)'] 表示：
          - 1.1 格式 → 第2级
          - （1）格式 → 第3级
          - 1) 格式 → 第4级

        同时打印信息提示用户修正了哪些级别。
        """
        if not self.auto_fix_heading_level:
            return

        print(f"\n=== 标题级别自动修正 ===")
        print(f"提供的格式列表: {self.auto_fix_heading_level}")
        print("将根据格式顺序推断正确的级别层次：")

        import re
        for idx, fmt in enumerate(self.auto_fix_heading_level):
            correct_level = idx + 2  # 从第2级开始
            # 尝试将格式转换为正则表达式
            # 保留数字和特殊字符的通配模式
            pattern = self._format_to_pattern(fmt)
            self._heading_level_map[pattern] = correct_level
            print(f"  格式 '{fmt}' → 第{correct_level}级")

        print("=" * 40)

    def _format_to_pattern(self, fmt: str) -> str:
        """
        将用户提供的格式字符串转换为正则表达式模式。

        Args:
            fmt: 格式字符串，例如 '1.1'、'（1）'、'1)' 等

        Returns:
            正则表达式模式字符串
        """
        import re
        # 先尝试直接转换为正则
        # 替换各种编号格式的通用模式
        pattern = fmt
        # 数字模式
        pattern = re.sub(r'\d+', r'\\d+', pattern)
        # 中文数字模式
        pattern = re.sub(r'[一二三四五六七八九十百千]+', r'[一二三四五六七八九十百千\\d]+', pattern)
        # 括号和分隔符保持原样
        # 移除外层空白
        pattern = pattern.strip()
        return pattern

    def _detect_format_in_text(self, text: str) -> str:
        """
        检测文本中的标题格式，匹配已建立的级别映射表。

        Args:
            text: 标题文本

        Returns:
            匹配到的格式对应的正确级别，如果无匹配返回 None
        """
        import re
        for pattern, correct_level in self._heading_level_map.items():
            try:
                # 尝试用正则匹配
                if re.search(pattern, text):
                    return correct_level
            except re.error:
                continue
        return None

    def add_copy(self, copy_path: str, author: str):
        """
        添加一个副本
        
        Args:
            copy_path: 副本文档路径
            author: 编辑者名称
        """
        copy_path = os.path.abspath(copy_path)
        
        if not os.path.exists(copy_path):
            raise FileNotFoundError(f"副本文件不存在: {copy_path}")
        
        self.copies.append((copy_path, author))
        print(f"✓ 已添加副本: {author} - {os.path.basename(copy_path)}")
    
    def merge(self, output_path: str, conflict_report_path: str = None) -> MergeResult:
        """
        执行累积式三路合并
        
        Args:
            output_path: 合并后的文档路径
            conflict_report_path: 冲突报告路径 (HTML)
        
        Returns:
            MergeResult: 合并结果
        """
        print("\n" + "="*60)
        print("开始累积式三路合并...")
        print("="*60)

        # 1. 拆解文档
        print("\n[1/5] 拆解文档...")
        self._disassemble_documents()

        # 2. 标题识别
        print("\n[2/5] 标题识别...")
        self._detect_all_headings()

        # 3. 标题匹配
        print("\n[3/5] 标题匹配...")
        self._match_all_headings()

        # 4. diff 检测
        print("\n[4/5] diff 检测...")
        self._detect_all_changes()

        # 5. 写回操作（包含冲突检测）
        print("\n[5/5] 写回操作...")
        written_count = self._write_changes(output_path)
        conflict_count = sum(len(conflict.changes) for conflict in self.conflicts)

        # 6. 生成冲突报告
        print("\n[6/6] 生成冲突报告...")
        if conflict_report_path is None:
            conflict_report_path = os.path.splitext(output_path)[0] + '_conflicts.html'
        
        self._generate_conflict_report(conflict_report_path)
        
        # 生成结果
        result = MergeResult(
            merged_document=output_path,
            conflicts=self.conflicts,
            auto_merged_count=written_count - conflict_count,
            manual_review_count=len(self.conflicts),
            conflict_report=conflict_report_path
        )
        
        return result

    def _disassemble_documents(self):
        """拆解文档：提取原稿和所有副本的结构"""
        self.original_structure = self.builder.extract_document_structure(self.original_doc)
        self.orig_items = self.original_structure['items']
        print(f"    ✓ 原稿解析完成: {len(self.orig_items)} 个块", flush=True)
        self.copy_items_list.clear()
        for i, (copy_path, author) in enumerate(self.copies):
            print(f"  [{i+1}/{len(self.copies)}] 解析 {author} 的副本...")
            print(f"    正在解析文档结构: {os.path.basename(copy_path)}...", flush=True)

            # 提取副本结构
            copy_structure = self.builder.extract_document_structure(copy_path)
            copy_items = copy_structure.get('items', [])
            print(f"    ✓ 结构解析完成: {len(copy_items)} 个块", flush=True)

            # 缓存副本 items
            self.copy_items_list.append(copy_items)

    def _detect_all_headings(self):
        """标题识别：根据提取逻辑正式提取原稿和所有副本锚点，并仅作打印展示"""
        self.orig_anchors = self._extract_anchors(self.orig_items)
        self.copy_anchors_list = []

        self._print_detected_headings("原稿", self.orig_anchors)

        for i, (copy_path, author) in enumerate(self.copies):
            copy_items = self.copy_items_list[i]
            copy_anchors = self._extract_anchors(copy_items)
            self.copy_anchors_list.append(copy_anchors)
            self._print_detected_headings(author, copy_anchors)

    def _match_all_headings(self):
        """标题配对：复用已提取的锚点并严格执行交叉配对，仅进行结果打印"""
        self.anchor_pairs_list.clear()
        self.orphan_orig_list.clear()
        self.orphan_copy_list.clear()
        for i, (copy_path, author) in enumerate(self.copies):
            copy_anchors = self.copy_anchors_list[i]

            anchor_pairs, orphan_orig, orphan_copy = self._match_anchors_between_docs(
                self.orig_anchors, copy_anchors,
                orig_items=self.orig_items,
                copy_items=self.copy_items_list[i]
            )
            self._print_heading_structure(
                author=author,
                copy_anchors=copy_anchors,
                orig_anchors=self.orig_anchors,
                anchor_pairs=anchor_pairs,
                orphan_orig=orphan_orig,
                orphan_copy=orphan_copy
            )

            # 缓存锚点配对数据供 diff 层调用
            self.anchor_pairs_list.append(anchor_pairs)
            self.orphan_orig_list.append(orphan_orig)
            self.orphan_copy_list.append(orphan_copy)

    def _detect_all_changes(self):
        """diff 检测：提取所有副本相对于原稿的变更"""
        self.changes_by_author.clear()
        for i, (copy_path, author) in enumerate(self.copies):
            print(f"  [{i+1}/{len(self.copies)}] diff 检测 {author} 的副本...")
            self._extract_changes(author, i)

            # 使用唯一键避免同名作者覆盖
            unique_key = f"{author}_{i}"
            print(f"    - 发现 {len(self.changes_by_author.get(unique_key, []))} 处变更")

        # 冲突检测
        self._detect_conflicts()

    def _extract_changes(self, author: str, author_idx: int):
        """
        基于锚点分块的 diff（直接使用成员变量）：
        - 1. 锚点对齐已在 _match_all_headings 完成，结果存储在成员变量
        - 2. 锚点之间用 SequenceMatcher 对齐 items
        - 3. 末尾块也参与 diff（不再直接删除）
        - 4. 不去重，保留所有变更（即使多个副本对同一位置做相同操作）
        """
        copy_items = self.copy_items_list[author_idx]
        anchor_pairs = self.anchor_pairs_list[author_idx]
        orphan_orig = self.orphan_orig_list[author_idx]
        orphan_copy = self.orphan_copy_list[author_idx]

        changes = []
        prev_orig_idx = -1  # 追踪上一个处理过的原稿锚点 index
        prev_copy_idx = -1  # 追踪上一个处理过的副本锚点 index

        # ── [新] 锚点同步与区块分析矩阵 ───────────────────────────────────
        print(f"\n    [锚点同步与区块分析矩阵] (原稿 {len(self.orig_anchors)} / 副本 {len(copy_items)} 块) - 作者: {author}:")
        
        # 定义列宽
        cw = [8, 8, 35, 8, 20, 15] # 状态, 原印, 内容, 副印, 间块规模, 预警
        header = (
            f"    | {self._pad_visual('状态', cw[0])} | "
            f"{self._pad_visual('原索引', cw[1])} | "
            f"{self._pad_visual('同步锚点内容', cw[2])} | "
            f"{self._pad_visual('副索引', cw[3])} | "
            f"{self._pad_visual('间块规模(原vs副)', cw[4])} | "
            f"逻辑预警 |"
        )
        visual_total_w = self._get_visual_width(header)
        sep = f"    {'-' * visual_total_w}"
        
        print(sep)
        print(header)
        print(sep)

        p_o, p_c = -1, -1
        for pair_idx, (o_anchor, c_anchor, text_sim) in enumerate(anchor_pairs):
            o_idx, c_idx = o_anchor['item_index'], c_anchor['item_index']
            
            # 计算间块规模
            c_o = o_idx - p_o - 1
            c_c = c_idx - p_c - 1
            block_info = f"{c_o}项 vs {c_c}项"
            
            # 逻辑健康检查
            alert = "[OK]"
            if c_o < 0 or c_c < 0:
                alert = "[⚠️ 索引回跳]"
            elif c_o != c_c:
                alert = "[内容变更]"
            
            text_disp = o_anchor['content'][:32] + ("..." if len(o_anchor['content']) > 32 else "")
            
            line = (
                f"    | {self._pad_visual('Sync ' + str(pair_idx), cw[0])} | "
                f"{self._pad_visual('#' + str(o_idx), cw[1])} | "
                f"{self._pad_visual(text_disp, cw[2])} | "
                f"{self._pad_visual('#' + str(c_idx), cw[3])} | "
                f"{self._pad_visual(block_info, cw[4])} | "
                f"{alert} |"
            )
            print(line)
            p_o, p_c = o_idx, c_idx

        # 末尾块收尾
        c_o_last = len(self.orig_items) - p_o - 1
        c_c_last = len(copy_items) - p_c - 1
        last_alert = "[OK]" if c_o_last == c_c_last else "[末尾新增/删]"
        
        print(f"    | {self._pad_visual('EndBlock', cw[0])} | "
              f"{self._pad_visual('-', cw[1])} | "
              f"{self._pad_visual('(文档末尾内容)', cw[2])} | "
              f"{self._pad_visual('-', cw[3])} | "
              f"{self._pad_visual(f'{c_o_last}项 vs {c_c_last}项', cw[4])} | "
              f"{last_alert} |")
        print(sep)

        # 孤立点摘要（仅作提示，位置已归入上述间块）
        if orphan_copy:
            print(f"    * 提示: 检测到 {len(orphan_copy)} 个副本孤立标题，已自动归入上述对应的[间块]中进行 Diff 处理。")
        print("")
        
        # 处理配对的锚点
        for pair_idx, (o_anchor, c_anchor, text_sim) in enumerate(anchor_pairs):
            o_idx = o_anchor['item_index']
            c_idx = c_anchor['item_index']

            if self.orig_items[o_idx].paragraph_signature != copy_items[c_idx].paragraph_signature:
                changes.extend(self._make_replace_change(
                    self.orig_items[o_idx], copy_items[c_idx], c_idx, author, author_idx
                ))

            # 处理锚点之间的块：从上一个锚点到当前锚点
            start_orig = prev_orig_idx + 1
            end_orig = o_idx

            start_copy = prev_copy_idx + 1
            end_copy = c_idx

            if start_orig < end_orig or start_copy < end_copy:
                block_changes = self._extract_block_items_diff(
                    self.orig_items, copy_items,
                    start_orig=start_orig, end_orig=end_orig,
                    start_copy=start_copy, end_copy=end_copy,
                    author=author,
                    copy_index=author_idx
                )
                changes.extend(block_changes)

            # 追踪已处理的原稿/副本锚点 index
            prev_orig_idx = o_idx
            prev_copy_idx = c_idx

        # (已删除孤立原稿/副本锚点的特殊处理，它们将作为块内内容由末尾 diff 统一处理)


        # ── 处理原稿末尾块（最后一个锚点之后 → 与副本对应范围 diff）────────
        # 原稿块：从最后一个锚点之后到文档末尾
        start_orig = prev_orig_idx + 1
        end_orig = len(self.orig_items)

        # 副本块：从最后一个副本锚点之后到文档末尾
        start_copy = prev_copy_idx + 1
        end_copy = len(copy_items)

        if start_orig < end_orig or start_copy < end_copy:
            block_changes = self._extract_block_items_diff(
                self.orig_items, copy_items,
                start_orig=start_orig, end_orig=end_orig,
                start_copy=start_copy, end_copy=end_copy,
                author=author,
                copy_index=author_idx
            )
            changes.extend(block_changes)

        # 智能去重（删除去重，新增/修改保留全部）
        changes = self._deduplicate_changes(changes)

        # 存储到成员变量
        unique_key = f"{author}_{author_idx}"
        self.changes_by_author[unique_key] = changes

        # 打印变更详情清单
        self._print_grouped_changes(changes)

    def _print_grouped_changes(self, changes: List[DocumentChange]) -> None:
        """用统一矩阵表格展示变更明细清单 (纯文本标签方案，确保 Windows 下 100% 对齐)"""
        if not changes: return
        
        print(f"\n    [变更执行明细清单] (共 {len(changes)} 项):")
        # [ ID | 全局位置 | 变更类型 | 作者 | 内容快照 ]
        cw = [5, 12, 12, 15, 60]
        header = (
            f"    | {self._pad_visual('ID', cw[0])} | "
            f"{self._pad_visual('位置', cw[1])} | "
            f"{self._pad_visual('类型', cw[2])} | "
            f"{self._pad_visual('作者', cw[3])} | "
            f"{self._pad_visual('变更详情 (原稿 ➔ 副本/新增)', cw[4])} |"
        )
        total_w = self._get_visual_width(header)
        sep = f"    {'-' * total_w}"
        
        print(sep); print(header); print(sep)
        
        for i, change in enumerate(changes):
            # 使用纯文字标签取代 Emoji 以解决对齐崩溃问题
            ctype_map = {
                'insert': '[新增]', 'delete': '[删除]', 'modify': '[修改]',
                'table': '[表格]', 'image': '[图片]', 'table_insert': '[表增]',
                'table_delete': '[表删]', 'image_insert': '[图增]', 'image_delete': '[图删]'
            }
            ctype_disp = ctype_map.get(change.change_type, f"[{change.change_type.upper()}]")
            pos_disp = f"#{getattr(change, 'true_block_index', change.position)}"
            author_disp = change.author if change.author else "-"
            
            # 详情提取
            if change.change_type == 'modify':
                try:
                    data = json.loads(change.content)
                    old = data.get('old_content', '').strip().replace('\n', ' ')[:25]
                    new = data.get('new_content', '').strip().replace('\n', ' ')[:25]
                    detail = f"{old} -> {new}"
                except: detail = str(change.content)[:55].replace('\n', ' ')
            else:
                detail = str(change.content)[:55].replace('\n', ' ')
            
            line = (
                f"    | {self._pad_visual(i+1, cw[0])} | "
                f"{self._pad_visual(pos_disp, cw[1])} | "
                f"{self._pad_visual(ctype_disp, cw[2])} | "
                f"{self._pad_visual(author_disp, cw[3])} | "
                f"{self._pad_visual(detail, cw[4])} |"
            )
            print(line)
        print(sep + "\n")

    def _print_diff_detail(self, change: DocumentChange):
        """打印差异详情，处理各种空值情况"""
        import json
        
        try:
            # 尝试解析 JSON
            if change.content.startswith('{'):
                annotation = json.loads(change.content)
            else:
                annotation = {}
        except (json.JSONDecodeError, TypeError):
            annotation = {}
        
        old_content = annotation.get('old_content', '')
        new_content = annotation.get('new_content', '')
        diff_segments = annotation.get('diff_segments', [])
        
        # 情况1: diff_segments 有内容
        if diff_segments:
            diff_text = self._format_diff_segments(diff_segments)
            
            # 如果 old_content/new_content 为空，从 diff_segments 提取
            if not old_content:
                # 从 diff_segments 提取原文部分
                old_parts = []
                for seg in diff_segments:
                    if seg.get('type') in ('equal', 'delete'):
                        old_parts.append(seg.get('text', ''))
                    elif seg.get('type') == 'replace':
                        old_parts.append(seg.get('old_text', ''))
                old_content = ''.join(old_parts)
            
            if not new_content:
                # 从 diff_segments 提取副本内容部分
                new_parts = []
                for seg in diff_segments:
                    if seg.get('type') in ('equal', 'insert'):
                        new_parts.append(seg.get('text', ''))
                    elif seg.get('type') == 'replace':
                        new_parts.append(seg.get('new_text', text))
                new_content = ''.join(new_parts)
            
            # 截断显示
            old_display = (old_content[:60] + '...') if len(old_content) > 60 else old_content
            new_display = (new_content[:60] + '...') if len(new_content) > 60 else new_content
            
            print(f"       原稿: {old_display}")
            print(f"       副本: {new_display}")
            print(f"       差异: {diff_text}")
        
        # 情况2: old_content 或 new_content 有内容
        elif old_content or new_content:
            old_display = (old_content[:60] + '...') if len(old_content) > 60 else old_content
            new_display = (new_content[:60] + '...') if len(new_content) > 60 else new_content
            
            if old_content and new_content:
                print(f"       原稿: {old_display}")
                print(f"       副本: {new_display}")
            elif old_content:
                print(f"       原稿: {old_display}")
                print(f"       变化: [删除全部]")
            else:
                print(f"       副本: {new_display}")
                print(f"       变化: [新增全部]")
        
        # 情况3: 都为空，显示原始内容
        else:
            content = change.content[:80] if len(change.content) > 80 else change.content
            print(f"       内容: {content}")
            print(f"       [警告: 无法解析差异详情]")

    def _format_diff_segments(self, segments: List[Dict]) -> str:
        """格式化 diff segments，高亮显示差异"""
        import re
        
        result_parts = []
        for seg in segments:
            seg_type = seg.get('type', 'equal')
            text = seg.get('text', '')
            
            if seg_type == 'equal':
                # 原文：截取前20字符
                if len(text) > 20:
                    result_parts.append(f"「{text[:20]}...」")
                else:
                    result_parts.append(f"「{text}」")
            elif seg_type == 'delete':
                # 删除：高亮显示
                result_parts.append(f"[删除:{text}]")
            elif seg_type == 'insert':
                # 新增：高亮显示
                result_parts.append(f"[新增:{text}]")
            elif seg_type == 'replace':
                old_text = seg.get('old_text', '')
                new_text = seg.get('new_text', text)
                result_parts.append(f"[替换:{old_text}→{new_text}]")
        
        return ''.join(result_parts)

    def _extract_changes_direct(self, copy_items: List[DocItem], author: str,
                                copy_index: int) -> List[DocumentChange]:
        """降级方案：无锚点时，直接全局 items diff"""
        orig_items = self.original_structure.get('items', [])

        def item_key(item):
            content = item.table_diff_key() if item.type == 'table' else item.content
            return (item.type, content)

        class _ItemWrapper:
            __slots__ = ('orig', 'sort_key')
            def __init__(self, orig, key):
                self.orig = orig
                self.sort_key = key(orig)
            def __eq__(self, other):
                return self.sort_key == other.sort_key
            def __hash__(self):
                return hash(self.sort_key)

        o_wrapped = [_ItemWrapper(it, item_key) for it in orig_items]
        c_wrapped = [_ItemWrapper(it, item_key) for it in copy_items]
        matcher = difflib.SequenceMatcher(None, o_wrapped, c_wrapped)

        changes = []
        for tag, o_start, o_end, c_start, c_end in matcher.get_opcodes():
            if tag == 'equal':
                for original, copy in zip(orig_items[o_start:o_end], copy_items[c_start:c_end]):
                    if original.paragraph_signature != copy.paragraph_signature:
                        changes.extend(self._make_replace_change(original, copy, original.index, author, copy_index))
            elif tag == 'replace':
                replace_len = min(o_end - o_start, c_end - c_start)
                for i in range(replace_len):
                    o_idx = o_start + i
                    c_idx = c_start + i
                    o_item = o_wrapped[o_idx].orig
                    c_item = c_wrapped[c_idx].orig
                    changes.extend(self._make_replace_change(o_item, c_item, o_idx, author, copy_index))
                if o_end - o_start > c_end - c_start:
                    for i in range(o_end - 1, o_start + replace_len - 1, -1):
                        changes.extend(self._make_delete_change(o_wrapped[i].orig, author, copy_index))
                elif c_end - c_start > o_end - o_start:
                    for i in range(c_end - 1, c_start + replace_len - 1, -1):
                        changes.extend(self._make_insert_change(c_wrapped[i].orig, o_end, author, copy_index))
            elif tag == 'delete':
                for i in range(o_end - 1, o_start - 1, -1):
                    changes.extend(self._make_delete_change(o_wrapped[i].orig, author, copy_index))
            elif tag == 'insert':
                for i in range(c_end - 1, c_start - 1, -1):
                    changes.extend(self._make_insert_change(c_wrapped[i].orig, o_start, author, copy_index))

        return changes

    # ------------------------------------------------------------------
    # 锚点提取与对齐
    # ------------------------------------------------------------------
    def _extract_anchors(self, items: List[DocItem]) -> List[Dict]:
        """从 items 列表中提取锚点（标题段落）"""
        # 第一步：提前计算整篇文档的字号基准
        doc_baseline = self.builder._compute_body_font_baseline(items)
        # 将基准注入 builder 状态，供 _is_legitimate_heading 使用
        self.builder._body_baseline = doc_baseline

        anchors = []
        for i, item in enumerate(items):
            if item.type != 'paragraph':
                continue
            # 使用动态基准进行判定
            if not self.builder._is_legitimate_heading(item, getattr(self, 'custom_headings', None)):
                continue
            
            # 标准化文本
            text_norm = _anchor_norm_text(item.content)
            anchors.append({
                'item_index': i,
                'content': item.content,
                'level': item.level,           # 原判级别
                'detected_level': 0,           # 待推断级别（初始为0）
                'text_norm': text_norm,
                'heading_score': getattr(item, 'heading_score', 0.0),
                'item': item,
            })

        # [纠偏机制] 针对人手录入笔误（如 1. 2、 3. 或 (1) (2] （3））进行逻辑序列与标点兼容处理
        self._fix_heading_punctuation_errors(anchors)

        # 计算视觉权重（作为锚点属性，不参与推理逻辑）
        for anchor in anchors:
            self._compute_visual_weight(anchor, doc_baseline)

        # 第二步：纯粹的树状结构推理
        anchors = self._infer_heading_levels(anchors, items)

        # 第三步：层级深度截断（受 --max-heading-level 控制）
        if self.max_heading_level > 0:
            before_count = len(anchors)
            anchors = [a for a in anchors if a.get('detected_level', 0) <= self.max_heading_level]
            after_count = len(anchors)
            if before_count != after_count:
                print(f"    [max-heading-level={self.max_heading_level}] 截断了 {before_count - after_count} 个超限标题，剩余 {after_count} 个锚点")

        return anchors



    def _compute_visual_weight(self, anchor: Dict, body_baseline: float) -> float:
        """
        计算一个锚点的“视觉重要性”分数。
        越大/越粗/越“标题字体” => 分数越高 => 应当是更高级别的标题。
        """
        item = anchor.get('item')
        if not item: return 0.0
        runs = item.runs
        if not runs: return 0.0

        total_chars = sum(len(r.get('text', '')) for r in runs)
        if total_chars == 0: return 0.0

        # 1. 加粗比 (0.3)
        bold_chars = sum(len(r.get('text', '')) for r in runs if r.get('bold'))
        bold_score = 0.3 if (bold_chars/total_chars) >= 0.5 else 0.0

        # 2. 字号相对得分 (0.7)
        avg_size = 0.0
        for r in runs:
            try:
                sz = float(r.get('fontSize', body_baseline))
                avg_size += sz * len(r.get('text', ''))
            except: avg_size += body_baseline * len(r.get('text', ''))
        avg_size /= total_chars
        
        size_ratio = avg_size / body_baseline
        size_score = min(0.7, max(0.0, (size_ratio - 1.0) * 0.5))

        weight = bold_score + size_score
        anchor['visual_weight'] = round(weight, 2)
        return weight

    def _infer_heading_levels(self, anchors: List[Dict], items: List[DocItem] = None) -> List[Dict]:
        """
        【第二步：纯粹的树状结构推理】
        基于序号前缀形式推断标题层级（不依赖格式样式信息，纯粹基于编号特征和上下文推导）。

        核心规则：
        1. 样式指纹：将编号抽象为 (括号, 数字系统, 分隔符, 嵌套深度)。
        2. 全局记忆：某种样式第一次出现时的深度，即为其“固定身份”。
        3. 动态栈：若 A 样式后出现 B 样式且 B 为新样式，则 B = A.depth + 1（子级）。
        4. 嵌套自述：1.1 等格式自带层级提示。
        """
        if not anchors:
            return []

        # 1. 状态初始化
        global_style_to_level = {} # 样式指纹 -> 确定的逻辑层级
        active_stack = []          # 当前路径下的样式指纹队列

        # 处理 custom_headings (若用户显式指定，则优先遵循)
        if getattr(self, 'custom_headings', None):
            # 如果设置了 max_heading_level，截断 custom_headings 列表
            effective_headings = self.custom_headings
            if self.max_heading_level > 0 and len(self.custom_headings) > self.max_heading_level:
                effective_headings = self.custom_headings[:self.max_heading_level]
                print(f"    [max-heading-level={self.max_heading_level}] custom-headings 截断为前 {self.max_heading_level} 级")
            saved = self.custom_headings
            self.custom_headings = effective_headings
            self._infer_by_custom_patterns(anchors)
            self.custom_headings = saved  # 恢复原始值
            return anchors

        # 2. 迭代推理
        max_l1_value = 0
        for anchor in anchors:
            # 获取样式指纹
            raw_text = anchor['content'].strip()
            # 优先检查纠偏结果 (来自 _fix_heading_punctuation_errors)
            override_fp = anchor.get('format_family_override')
            style_fp = override_fp if override_fp else self._get_strict_style_fingerprint(raw_text)
            
            # 若无任何编号特征且非强制覆盖，则视为正文或默认一级
            if not style_fp or style_fp == "TEXT":
                anchor['detected_level'] = 0
                anchor['parent_level'] = None
                active_stack = [] # 重置栈
                continue

            # A. 确定该样式的层级
            if style_fp in global_style_to_level:
                # 已知样式：使用历史记录，并回跳栈
                level = global_style_to_level[style_fp]
                # 同步路径栈（弹出当前样式之后的所有样式）
                if style_fp in active_stack:
                    idx = active_stack.index(style_fp)
                    active_stack = active_stack[:idx+1]
                else:
                    # 文档结构跳跃或不规范，强制覆盖栈
                    active_stack = [style_fp]
            else:
                # 新样式：推导其深度
                if style_fp.startswith('NESTED_N_'):
                    # 嵌套指纹（如 1.1），包含自述深度 K
                    try:
                        k = int(style_fp.split('_')[-1])
                    except: k = 2
                    
                    # 💡 嵌套补正：尝试寻找其基准样式 (如 1. -> NONE_AR_DOT_1)
                    # 如果基准样式已存在，则层级 = 基准层级 + (K-1)
                    base_fp = "NONE_AR_DOT_1" # 默认基准
                    if base_fp in global_style_to_level:
                        level = global_style_to_level[base_fp] + (k - 1)
                    elif not active_stack:
                        level = k # 无上下文，按自述深度定
                    else:
                        # 随大流：参考当前栈顶深度
                        level = global_style_to_level[active_stack[-1]] + 1
                else:
                    # 单级指纹（如 一、, 1.）
                    if not active_stack:
                        level = 1
                    else:
                        # 紧跟在当前栈顶之后的未知样式，视为子级
                        parent_style = active_stack[-1]
                        parent_level = global_style_to_level[parent_style]
                        level = parent_level + 1
                
                global_style_to_level[style_fp] = level
                active_stack.append(style_fp)

            # [新增] L1 级标题降序/重复限制逻辑
            # 原理：若当前推断为 1 级标题，检查其数值是否小于等于已见过的最大值，若是则判定为正文。
            if level == 1:
                nums = self._extract_numbers(raw_text)
                if nums:
                    val = nums[0]
                    if val <= max_l1_value:
                        # 发现 L1 回跳（如 二、 之后又出现了 一、），强制降为正文
                        level = 0
                        style_fp = "TEXT"
                        active_stack = [] # 回跳通常意味着章节结束，重置栈
                    else:
                        max_l1_value = val

            # B. 写入锚点信息
            anchor['detected_level'] = level
            anchor['style_fingerprint'] = style_fp
            
            # 💡 核心改进：记录物理父节点索引（用于后续章节隔离匹配）
            # 我们需要找到最近的一个级别为 level - 1 的锚点索引
            parent_idx = None
            if level > 0:
                for prev_idx in range(len(anchors) - 1, -1, -1):
                    # 我们只能查找当前锚点之前的元素
                    if anchors[prev_idx] == anchor:
                        # 找到自己了，再往前找
                        continue
                    # 如果这个查找逻辑在迭代中做太慢，可以在后面统一修正
                    pass 
                
                # 优化：使用一个临时查找表
                if not hasattr(self, '_tmp_active_levels'):
                    self._tmp_active_levels = {}
                self._tmp_active_levels[level] = anchors.index(anchor)
                anchor['parent_anchor_idx'] = self._tmp_active_levels.get(level - 1)
                
                # parent_level 参考：栈中前一个元素的级别
                if len(active_stack) >= 2:
                    prev_style = active_stack[-2]
                    anchor['parent_level'] = global_style_to_level.get(prev_style)
                else:
                    anchor['parent_level'] = None
            else:
                anchor['parent_anchor_idx'] = None
                anchor['parent_level'] = None

        # 清理临时状态
        if hasattr(self, '_tmp_active_levels'):
            delattr(self, '_tmp_active_levels')

        # 3. 同步到 DocItem 并打印结果
        print("\n[层级推理报告] 基于纯序列表征推导结果:")
        for anchor in anchors:
            if 'item' in anchor:
                anchor['item'].detected_level = anchor['detected_level']
            
            lv = anchor['detected_level']
            fp = anchor.get('style_fingerprint', 'TEXT')
            text_preview = anchor['content'].strip()[:40]
            print(f"  Level {lv} | Style: {fp.ljust(15)} | {text_preview}")
        print("-" * 50)


        return anchors

    def _get_strict_style_fingerprint(self, text: str) -> str:
        """
        提取标题的“严格样式指纹”。
        不根据数字的大小，只根据数字的类型（阿拉伯/中文/字母/罗马）和修饰符（括号/点/顿号）。
        """
        import re
        text = text.strip()
        if not text: return "TEXT"

        # 归一化全角符号
        norm_text = text.replace('（', '(').replace('）', ')').replace('【', '[').replace('】', ']').replace('．', '.').replace('：', ':').replace('、', 'DOT_CN').replace('。', '.')

        # 1. 优先匹配嵌套数字 (如 1.1, 1.2.3)
        nested_match = re.match(r'^([0-9]+(\.[0-9]+)+)', norm_text)
        if nested_match:
            dots_count = nested_match.group(1).count('.')
            return f"NESTED_N_{dots_count + 1}"

        # 2. 匹配具体样式 (Bracket_System_Separator)
        RE_AR = r'[0-9]+'
        RE_CN = r'[一二三四五六七八九十百千万零〇]+'
        RE_RM = r'[IVXLCDMivxlcdm]+'
        RE_LT = r'[a-zA-Z]'
        
        systems = [('AR', RE_AR), ('CN', RE_CN), ('RM', RE_RM), ('LT', RE_LT)]
        brackets = [
            ('PAREN', r'\((%s)\)'),    # (1)
            ('BRACKET', r'\[(%s)\]'),   # [1]
            ('NONE', r'^(%s)')          # 1.
        ]
        separators = [
            ('DOT_CN', 'DOT_CN'),      # 、
            ('DOT', r'\.'),            # .
            ('COLON', r':'),           # :
            ('B_DOT', r'\)'),          # 1)
            ('NONE', r'\s*')           # 1 (后面没标点)
        ]

        for b_name, b_pat in brackets:
            for s_name, s_pat in systems:
                pattern = b_pat % s_pat
                match = re.search(pattern, norm_text)
                if match:
                    # 检查后缀分隔符
                    suffix = norm_text[match.end():].lstrip()
                    matched_sep = "NONE"
                    for sep_name, sep_pat in separators:
                        if re.match('^' + sep_pat, suffix):
                            matched_sep = sep_name
                            break
                    
                    return f"{b_name}_{s_name}_{matched_sep}_1"

        return "TEXT"

    def _infer_by_custom_patterns(self, anchors: List[Dict]):
        """保留原有的 custom_headings 逻辑支撑"""
        import re
        for anchor in anchors:
            text = anchor['content'].strip()
            matched_level = 0
            for idx, fmt in enumerate(self.custom_headings):
                pattern = self._format_to_pattern(fmt)
                try:
                    if re.match(r'^' + pattern, text):
                        matched_level = idx + 1
                        break
                except:
                    if text.startswith(fmt.strip()):
                        matched_level = idx + 1
                        break
            anchor['detected_level'] = matched_level
            anchor['parent_level'] = None
            if 'item' in anchor:
                anchor['item'].detected_level = matched_level

    def _parse_cn_number(self, s: str) -> int:
        """正确解析中文数字字符串，支持 一~九千九百九十九"""
        cn_map = {'零':0, '一':1, '二':2, '三':3, '四':4, '五':5, '六':6, '七':7, '八':8, '九':9,
                  '〇':0, '壹':1, '贰':2, '叁':3, '肆':4, '伍':5, '陆':6, '柒':7, '捌':8, '玖':9,
                  '貮':2, '两':2}
        units = {'十': 10, '百': 100, '千': 1000, '万': 10000}
        if not s: return 0
        
        # 简单情况：如 "一", "二"
        if len(s) == 1 and s in cn_map:
            return cn_map[s]
            
        res = 0
        unit = 1
        for i in range(len(s) - 1, -1, -1):
            char = s[i]
            if char in units:
                unit = units[char]
                if unit == 10 and i == 0: # 处理 "十", "十二" 这种开头的
                    res += 10
            elif char in cn_map:
                res += cn_map[char] * unit
        return res

    def _extract_numbers(self, text: str) -> List[int]:
        """提取文本中的所有数字（含中文、阿拉伯）"""
        import re
        nums = re.findall(r'[0-9]+|[一二三四五六七八九十百千]+', text[:20])
        res = []
        for n in nums:
            if n.isdigit():
                res.append(int(n))
            else:
                res.append(self._parse_cn_number(n))
        return res

    def _fix_heading_punctuation_errors(self, anchors: List[Dict]):
        """
        [纠偏进阶] 处理标题识别中的"笔误"情况（如 1. -> 2、 -> 3.）。
        逻辑：如果两个标题的数字呈现明显的 N -> N+1 递增关系且基础样式一致，则强制归并为同一指纹族。
        """
        if len(anchors) < 2: return

        def _fp_base_family(fp):
            """提取指纹的基础家族（括号类型+数字系统），如 PAREN_CN_NONE_1 -> PAREN_CN"""
            parts = fp.split('_')
            if len(parts) >= 2:
                return parts[0] + '_' + parts[1]
            return fp

        for i in range(1, len(anchors)):
            curr = anchors[i]
            prev = anchors[i-1]
            
            p_fp = self._get_strict_style_fingerprint(prev['content'])
            c_fp = self._get_strict_style_fingerprint(curr['content'])
            
            if p_fp != c_fp:
                # 安全约束：跨编号家族（如 PAREN_CN vs NONE_AR）绝不允许覆盖
                if _fp_base_family(p_fp) != _fp_base_family(c_fp):
                    continue

                p_nums = self._extract_numbers(prev['content'])
                c_nums = self._extract_numbers(curr['content'])
                
                # 若数字递增且深度一致，尝试修正
                if p_nums and c_nums and len(p_nums) == len(c_nums):
                    if p_nums[:-1] == c_nums[:-1] and c_nums[-1] == p_nums[-1] + 1:
                        # 确定是由于标点差异导致的指纹分叉，强制同步
                        curr['format_family_override'] = p_fp

    def _match_anchors_between_docs(self, orig_anchors: List[Dict],
                                     copy_anchors: List[Dict],
                                     orig_items: List[DocItem] = None,
                                     copy_items: List[DocItem] = None
                                     ) -> Tuple[List, List, List]:
        """
        将原稿锚点与副本锚点对齐（参照参考项目实现）。

        匹配策略：
        - 对每个原稿锚点，按顺序在副本中找最佳匹配
        - 综合评分 = 文本相似度×0.65 + 上下文相似度×0.25 + 位置相似度×0.10
        - 必须满足 anchor_threshold 阈值
        - 顺序约束：副本锚点不能回头匹配（last_c_idx 机制）

        返回：
        - anchor_pairs: [(orig_anchor, copy_anchor), ...] 配对的锚点
        - orphan_orig: 原稿中找不到配对的锚点
        - orphan_copy: 副本中找不到配对的锚点
        """
        if not orig_anchors:
            return [], [], copy_anchors
        if not copy_anchors:
            return [], orig_anchors, []

        import difflib
        n_orig = len(orig_anchors)
        n_copy = len(copy_anchors)

        # ── 上下文提取函数 ────────────────────────────────────────────────
        def _heading_context(paras: List[DocItem], anchor_item_index: int, window: int = 3) -> str:
            """提取标题的上下文（前后各 window 个段落的内容拼接）"""
            if paras is None:
                return ''
            parts = []
            lo = max(0, anchor_item_index - window)
            hi = min(len(paras), anchor_item_index + window + 1)
            for k in range(lo, hi):
                if k != anchor_item_index:
                    t = paras[k].content if hasattr(paras[k], 'content') else ''
                    if t:
                        parts.append(t[:60].strip())
            return ' | '.join(parts)

        # ── 综合评分函数 ──────────────────────────────────────────────────
        def _anchor_score(o_text_norm: str, c_text_norm: str,
                          o_item_idx: int, c_item_idx: int) -> float:
            """
            综合得分 = 文本相似度×0.65 + 上下文相似度×0.25 + 位置相似度×0.10
            """
            # ① 文本相似度
            if self.anchor_threshold >= 1.0:
                text_sim = 1.0 if o_text_norm == c_text_norm else 0.0
            else:
                text_sim = difflib.SequenceMatcher(None, o_text_norm, c_text_norm).ratio()

            # ② 上下文相似度
            o_ctx = _heading_context(orig_items, o_item_idx) if orig_items else ''
            c_ctx = _heading_context(copy_items, c_item_idx) if copy_items else ''
            if o_ctx and c_ctx:
                ctx_sim = difflib.SequenceMatcher(None, o_ctx, c_ctx).ratio()
            else:
                ctx_sim = 0.5

            # ③ 相对位置接近度
            o_rel = o_item_idx / max(n_orig - 1, 1)
            c_rel = c_item_idx / max(n_copy - 1, 1)
            pos_sim = max(0.0, 1.0 - abs(o_rel - c_rel) * 2)

            return 0.65 * text_sim + 0.25 * ctx_sim + 0.10 * pos_sim

        # ── 锚点匹配（实施基于树结构的硬隔离） ───────────────────────────────
        pairs = []
        matched_o_to_c = {} # {orig_anchor_idx: copy_anchor_idx}
        
        # 记录每个父节点下已匹配到的副本锚点位置，确保持续向下的顺序约束
        # key: copy_parent_idx (None for L1), val: last_matched_c_idx_within_this_parent
        last_c_idx_by_parent = {} 

        for oi, o_anchor in enumerate(orig_anchors):
            level = o_anchor.get('detected_level', 0)
            o_text_norm = o_anchor.get('text_norm', '')
            o_item_idx = o_anchor.get('item_index', 0)
            o_parent_idx = o_anchor.get('parent_anchor_idx')
            
            # 💡 第一性原理：章节隔离
            # 确定副本中允许匹配的候选池范围
            can_match_p_c_idx = None
            if o_parent_idx is None:
                # 一级标题：对应的副本父节点也必须是 None
                can_match_p_c_idx = None
            else:
                # 多级标题：只有当原稿父节点已匹配成功，且候选人的父节点正是该匹配项时，才允许匹配
                can_match_p_c_idx = matched_o_to_c.get(o_parent_idx)
                if can_match_p_c_idx is None:
                    # 如果原稿父节点未匹配成功，则其子节点在硬隔离原则下失去匹配基准
                    continue 

            # 过滤符合层级要求且父节点对齐的副本候选人
            copy_candidates = []
            for ci, c_anchor in enumerate(copy_anchors):
                if (c_anchor.get('detected_level') == level and 
                    c_anchor.get('parent_anchor_idx') == can_match_p_c_idx):
                    copy_candidates.append((ci, c_anchor))

            last_limit_idx = last_c_idx_by_parent.get(can_match_p_c_idx, -1)
            
            best_score = -1e9
            best_c_info = None # (copy_idx, text_sim)

            for ci, c_anchor in copy_candidates:
                if ci <= last_limit_idx:
                    continue  # 顺序约束

                c_text_norm = c_anchor.get('text_norm', '')
                c_item_idx = c_anchor.get('item_index', 0)

                # 计算文本相似度
                if self.anchor_threshold >= 1.0:
                    text_sim = 1.0 if o_text_norm == c_text_norm else 0.0
                else:
                    text_sim = difflib.SequenceMatcher(None, o_text_norm, c_text_norm).ratio()

                if text_sim < self.anchor_threshold:
                    continue

                # 既然已经通过了父节点硬隔离，此处的评分仅用于处理同父节点下的重名情况
                score = _anchor_score(o_text_norm, c_text_norm, o_item_idx, c_item_idx)

                if score > best_score:
                    best_score = score
                    best_c_info = (ci, text_sim)

            if best_c_info:
                target_ci, sim = best_c_info
                pairs.append((o_anchor, copy_anchors[target_ci], sim))
                matched_o_to_c[oi] = target_ci
                last_c_idx_by_parent[can_match_p_c_idx] = target_ci

        # 分离出孤立点
        matched_copy_indices = set(matched_o_to_c.values())
        orphan_orig = [a for i, a in enumerate(orig_anchors) if i not in matched_o_to_c]
        orphan_copy = [a for i, a in enumerate(copy_anchors) if i not in matched_copy_indices]

        return pairs, orphan_orig, orphan_copy

    # ------------------------------------------------------------------
    # 锚点间块内容 diff（统一 items，不区分类型）
    # ------------------------------------------------------------------
    def _extract_block_items_diff(self,
                                  orig_items: List[DocItem],
                                  copy_items: List[DocItem],
                                  start_orig: int, end_orig: int,
                                  start_copy: int, end_copy: int,
                                  author: str,
                                  copy_index: int) -> List[DocumentChange]:
        """
        在锚点之间的块范围内，用 SequenceMatcher 对齐原稿和副本的 items。
        统一处理段落、表格、图片，按全局 index 定位。
        
        策略改进：
        - 有内容的段落：按内容匹配
        - 空内容的段落：按相对位置+类型匹配（避免 id 不同导致无法匹配）
        """
        changes = []

        # 提取块内 items
        orig_block = orig_items[start_orig:end_orig]
        copy_block = copy_items[start_copy:end_copy]

        if not orig_block and not copy_block:
            return changes

        # 改进的 item key：空内容按位置匹配，有内容按内容匹配
        def item_key(item: DocItem, offset: int) -> tuple:
            """返回匹配 key：(has_content, type, content/key)"""
            if item.type == 'table':
                return (1, item.type, item.table_diff_key())
            content = item.content.strip() if item.content else ""
            if content:
                return (1, item.type, content)  # 有内容：优先按内容匹配
            else:
                # 空内容：按类型+相对位置匹配
                return (0, item.type, offset)  # offset 是块内相对位置

        class _SmartWrapper:
            __slots__ = ('orig', 'sort_key', 'block_offset', 'item_type')
            def __init__(self, orig, offset: int):
                self.orig = orig
                self.sort_key = item_key(orig, offset)
                self.block_offset = offset
                self.item_type = orig.type
            def __eq__(self, other):
                # 位置接近且类型相同视为匹配
                if self.sort_key[0] == 1 and other.sort_key[0] == 1:
                    # 都有内容：按内容匹配
                    return self.sort_key == other.sort_key
                else:
                    # 有空内容的：类型相同 + 位置接近(差<=2) 
                    type_match = self.item_type == other.item_type
                    pos_close = abs(self.block_offset - other.block_offset) <= 2
                    return type_match and pos_close
            def __hash__(self):
                return hash(self.sort_key)

        o_wrapped = [_SmartWrapper(it, i) for i, it in enumerate(orig_block)]
        c_wrapped = [_SmartWrapper(it, i) for i, it in enumerate(copy_block)]

        matcher = difflib.SequenceMatcher(None, o_wrapped, c_wrapped)
        opcodes = matcher.get_opcodes()
        
        # 收集本块的 diff 统计
        block_changes_summary = []
        for tag, o_s, o_e, c_s, c_e in opcodes:
            if tag == 'equal': continue
            block_changes_summary.append({
                'tag': tag,
                'o_range': f"#{start_orig+o_s}-#{start_orig+o_e}",
                'c_range': f"#{start_copy+c_s}-#{start_copy+c_e}",
                'desc': f"{o_e-o_s}项 ➔ {c_e-c_s}项"
            })
        
        if block_changes_summary:
            print(f"       [区块详述] 发现 {len(block_changes_summary)} 处结构变动:")
            b_cw = [10, 15, 15, 15]
            b_header = f"         | {self._pad_visual('变动', b_cw[0])} | {self._pad_visual('原稿区间', b_cw[1])} | {self._pad_visual('副本区间', b_cw[2])} | {self._pad_visual('规模特征', b_cw[3])} |"
            b_sep = f"         {'-' * self._get_visual_width(b_header)}"
            print(b_sep); print(b_header); print(b_sep)
            for bs in block_changes_summary:
                tag_map = {'replace': '[替换]', 'delete': '[删除]', 'insert': '[插入]'}
                b_line = (
                    f"         | {self._pad_visual(tag_map.get(bs['tag'], bs['tag']), b_cw[0])} | "
                    f"{self._pad_visual(bs['o_range'], b_cw[1])} | "
                    f"{self._pad_visual(bs['c_range'], b_cw[2])} | "
                    f"{self._pad_visual(bs['desc'], b_cw[3])} |"
                )
                print(b_line)
            print(b_sep)
        
        # 逆序处理变更操作：避免 index 漂移
        # priority: 0=删除, 1=修改, 2=新增
        pending_changes = []  # (priority, change_list)
        
        for tag, o_start, o_end, c_start, c_end in opcodes:
            if tag == 'equal':
                for original, copy in zip(orig_block[o_start:o_end], copy_block[c_start:c_end]):
                    if original.paragraph_signature != copy.paragraph_signature:
                        changes.extend(self._make_replace_change(original, copy, original.index, author, copy_index))

            elif tag == 'replace':
                # 记录相匹配的替换区间长度
                replace_len = min(o_end - o_start, c_end - c_start)
                
                # 修改：直接替换（不影响插入顺序）
                for i in range(replace_len):
                    o_idx = o_start + i
                    c_idx = c_start + i
                    o_item = o_wrapped[o_idx].orig
                    c_item = c_wrapped[c_idx].orig
                    # 使用原稿 index 作为变更 position
                    orig_pos = start_orig + o_idx
                    changes.extend(self._make_replace_change(o_item, c_item, orig_pos, author, copy_index))

                # 多删少增：多出的部分当作删除
                if o_end - o_start > c_end - c_start:
                    # 逆序收集删除，确保从下往上删
                    for i in range(o_end - 1, o_start + replace_len - 1, -1):
                        pending_changes.append((0, self._make_delete_change(o_wrapped[i].orig, author, copy_index)))
                
                # 多增少删：多出的部分当作新增
                elif c_end - c_start > o_end - o_start:
                    # 顺序收集新增，在稳定的块指针下自然按序叠加
                    insert_pos = start_orig + o_end
                    for i in range(c_start + replace_len, c_end):
                        pending_changes.append((2, self._make_insert_change(c_wrapped[i].orig, insert_pos, author, copy_index)))

            elif tag == 'delete':
                # 删除：逆序收集（从下往上删）
                for i in range(o_end - 1, o_start - 1, -1):
                    pending_changes.append((0, self._make_delete_change(o_wrapped[i].orig, author, copy_index)))

            elif tag == 'insert':
                # 新增：顺序收集以恢复正确的相对顺序
                insert_pos = start_orig + o_start
                for i in range(c_start, c_end):
                    pending_changes.append((2, self._make_insert_change(c_wrapped[i].orig, insert_pos, author, copy_index)))
        
        # 按 priority 排序并处理（先删后增）
        pending_changes.sort(key=lambda x: x[0])
        for _, change_list in pending_changes:
            changes.extend(change_list)

        return changes

    def _deduplicate_changes(self, changes: List[DocumentChange]) -> List[DocumentChange]:
        """
        智能去重：
        - 删除操作去重：同一位置只保留一个删除（无论来自哪个作者）
        - 新增和修改不去重：都保留（即使同一位置）
        
        原因：多个副本对同一位置都做删除，效果等价，只保留一个
        而新增和修改可能来自不同的修改基点（如副本1基于原稿改，副本2基于副本1改），
        需要全部保留到最终文档
        """
        # 删除操作去重：(position) -> first delete change
        delete_seen = {}  # position -> delete change
        
        # 新增和修改保留所有
        insert_changes = []
        replace_changes = []
        other_changes = []
        
        for ch in changes:
            if ch.change_type in ('insert', 'paragraph_insert') or '_insert' in ch.change_type:
                insert_changes.append(ch)
            elif ch.change_type in ('replace', 'paragraph_replace') or '_replace' in ch.change_type:
                replace_changes.append(ch)
            elif ch.change_type in ('delete', 'paragraph_delete') or '_delete' in ch.change_type:
                # 删除去重：同一位置只保留一个
                pos = ch.position
                if pos not in delete_seen:
                    delete_seen[pos] = ch
            else:
                other_changes.append(ch)
        
        # 合并所有结果
        result = (
            list(delete_seen.values()) +
            insert_changes +
            replace_changes +
            other_changes
        )
        # 按 position 排序
        result.sort(key=lambda c: c.position)
        return result

    # ------------------------------------------------------------------
    # 变更构建方法
    # ------------------------------------------------------------------
    def _make_delete_change(self, item: DocItem, author: str, copy_index: int) -> List[DocumentChange]:
        """生成删除变更"""
        return [DocumentChange(
            change_id=f"{author}_{copy_index}_del_{item.index}",
            change_type=f'{item.type}_delete' if item.type != 'paragraph' else 'delete',
            position=item.index,
            content=item.content,
            author=author,
            timestamp=datetime.now(),
            copy_index=copy_index,
            page_number=item.page,
            paragraph_index=item.index,
            para_order=item.index,
            format_info={'item_type': item.type},
        )]

    def _make_insert_change(self, item: DocItem, position: int, author: str, copy_index: int) -> List[DocumentChange]:
        """生成新增变更"""
        format_info = {'item_type': item.type, 'runs': item.runs}
        
        # 将原汁原味的 XML 结构全部保留传下去（不限于表格，段落等照样提取）
        if item.type in ('table', 'paragraph', 'sdt') and item.xml is not None:
            format_info['copy_xml'] = etree.tostring(item.xml, encoding='unicode')
            if item.type == 'table':
                # 兼容旧字段
                format_info['copy_tbl_xml'] = format_info['copy_xml']

        # # 💡 核心改进：在 format_info 中携带推断出的正确层级
        # if hasattr(item, 'detected_level') and item.detected_level > 0:
        #     format_info['actual_level'] = item.detected_level

        return [DocumentChange(
            change_id=f"{author}_{copy_index}_ins_{item.index}",
            change_type=f'{item.type}_insert' if item.type != 'paragraph' else 'insert',
            position=position,
            content=item.content,
            author=author,
            timestamp=datetime.now(),
            copy_index=copy_index,
            page_number=item.page,
            paragraph_index=position,
            para_order=position,
            format_info=format_info,
        )]

    def _make_replace_change(self, o_item: DocItem, c_item: DocItem,
                              copy_position: int, author: str,
                              copy_index: int) -> List[DocumentChange]:
        """
        生成替换变更：
        - 段落：runs diff，生成 diff_segments，content = JSON(diff_segments)
        - 表格：单元格 diff，content = JSON(diff_cells)
        - 图片：content = "图片被替换"
        """
        changes = []

        if o_item.type == 'paragraph' and c_item.type == 'paragraph':
            # runs diff（仅用于日志展示，不再用于重建段落）
            diff_segments = self._diff_runs(o_item.runs, c_item.runs,
                                             o_item.content, c_item.content)
            # 序列化副本的完整段落 XML（用于直接写回，原样保留）
            copy_para_xml = ''
            copy_pPr_xml = ''
            if c_item.xml is not None:
                copy_para_xml = etree.tostring(c_item.xml, encoding='unicode')
                c_pPr = c_item.xml.find(_wtag('pPr'))
                if c_pPr is not None:
                    copy_pPr_xml = etree.tostring(c_pPr, encoding='unicode')
            annotation = {
                'type': 'paragraph',
                'old_content': o_item.content,
                'new_content': c_item.content,
                'diff_segments': diff_segments,
                'copy_para_xml': copy_para_xml,
                'copy_pPr_xml': copy_pPr_xml,
                'format_info': {
                    'item_type': 'paragraph',
                    'runs': c_item.runs,
                    'diff_segments': diff_segments,
                    'copy_para_xml': copy_para_xml,
                    'run_format_diffs': self._extract_run_format_diffs(o_item.runs, c_item.runs),
                },
            }
            # # 💡 核心改进：在修改中也携带推断出的层级
            # if hasattr(c_item, 'detected_level') and c_item.detected_level > 0:
            #     annotation['format_info']['actual_level'] = c_item.detected_level

            changes.append(DocumentChange(
                change_id=f"{author}_{copy_index}_mod_{o_item.index}",
                change_type='modify',
                position=o_item.index,
                content=json.dumps(annotation, ensure_ascii=False),
                author=author,
                timestamp=datetime.now(),
                copy_index=copy_index,
                page_number=o_item.page,
                paragraph_index=o_item.index,
                para_order=o_item.index,
                format_info=annotation['format_info'],
            ))

        elif o_item.type == 'table' and c_item.type == 'table':
            # 表格单元格 diff
            diff_cells = self._diff_table_cells(o_item.rows, c_item.rows)
            format_info = {
                'item_type': 'table',
                'diff_cells': diff_cells,
                'structure_changed': o_item.table_shape() != c_item.table_shape(),
                'copy_tbl_xml': etree.tostring(c_item.xml, encoding='unicode'),
            }
            changes.append(DocumentChange(
                change_id=f"{author}_{copy_index}_tbl_{o_item.index}",
                change_type='table',
                position=o_item.index,
                content=json.dumps(diff_cells, ensure_ascii=False),
                author=author,
                timestamp=datetime.now(),
                copy_index=copy_index,
                page_number=o_item.page,
                paragraph_index=o_item.index,
                para_order=o_item.index,
                format_info=format_info,
            ))

        elif o_item.type == 'image':
            changes.append(DocumentChange(
                change_id=f"{author}_{copy_index}_img_mod_{o_item.index}",
                change_type='image',
                position=o_item.index,
                content=f"图片被替换: {o_item.item_id} → {c_item.item_id}",
                author=author,
                timestamp=datetime.now(),
                copy_index=copy_index,
                page_number=o_item.page,
                paragraph_index=o_item.index,
                para_order=o_item.index,
                format_info={'item_type': 'image'},
            ))

        else:
            # 类型不同：按删除+新增处理
            changes.extend(self._make_delete_change(o_item, author, copy_index))
            changes.extend(self._make_insert_change(c_item, copy_position, author, copy_index))

        return changes

    def _diff_runs(self, o_runs: List[dict], c_runs: List[dict],
                    o_text: str, c_text: str) -> List[dict]:
        """
        对比两个段落的 runs，生成 diff_segments。
        统一背景=无、字体=黑色，差异文字标红色。
        """
        def run_text(r: dict) -> str:
            return r.get('text', '')

        # Python 3.8 兼容：使用 wrapper 替代 key= 参数
        class _RunWrapper:
            __slots__ = ('orig', 'sort_key')
            def __init__(self, orig, key):
                self.orig = orig
                self.sort_key = key(orig)
            def __eq__(self, other):
                return self.sort_key == other.sort_key
            def __hash__(self):
                return hash(self.sort_key)

        o_wrapped = [_RunWrapper(r, run_text) for r in o_runs]
        c_wrapped = [_RunWrapper(r, run_text) for r in c_runs]
        matcher = difflib.SequenceMatcher(None, o_wrapped, c_wrapped)
        segments = []

        for tag, o_s, o_e, c_s, c_e in matcher.get_opcodes():
            if tag == 'equal':
                # 文字相同，但输出副本的样式（run 边界可能不同）
                for i in range(c_s, c_e):
                    r = c_wrapped[i].orig
                    segments.append({
                        'text': r.get('text', ''),
                        'is_diff': False,
                        'attrs': {
                            'bold': r.get('bold', False),
                            'italic': r.get('italic', False),
                            'underline': r.get('underline', ''),
                            'color': r.get('color', '000000'),
                            'background': r.get('background', ''),
                            'font': r.get('font', ''),
                            'fontSize': r.get('fontSize', ''),
                        },
                    })
            else:
                # replace / insert / delete: 保留副本样式，但差异标红
                for i in range(o_s, o_e):
                    r = o_wrapped[i].orig
                    segments.append({
                        'text': r.get('text', ''),
                        'is_diff': True,
                        'attrs': {
                            'bold': r.get('bold', False),
                            'italic': r.get('italic', False),
                            'underline': r.get('underline', ''),
                            'color': 'FF0000',
                            'background': '',
                            'font': r.get('font', ''),
                            'fontSize': r.get('fontSize', ''),
                        },
                    })
                for i in range(c_s, c_e):
                    r = c_wrapped[i].orig
                    segments.append({
                        'text': r.get('text', ''),
                        'is_diff': True,
                        'attrs': {
                            'bold': r.get('bold', False),
                            'italic': r.get('italic', False),
                            'underline': r.get('underline', ''),
                            'color': 'FF0000',
                            'background': '',
                            'font': r.get('font', ''),
                            'fontSize': r.get('fontSize', ''),
                        },
                    })

        return segments

    def _extract_run_format_diffs(self, o_runs: List[dict], c_runs: List[dict]) -> List[dict]:
        """检测文字相同但格式不同的 runs（用于打印展示）        """
        def run_text(r: dict) -> str:
            return r.get('text', '')

        class _RunWrapper:
            __slots__ = ('orig', 'sort_key')
            def __init__(self, orig, key):
                self.orig = orig
                self.sort_key = key(orig)
            def __eq__(self, other):
                return self.sort_key == other.sort_key
            def __hash__(self):
                return hash(self.sort_key)

        diffs = []
        o_wrapped = [_RunWrapper(r, run_text) for r in o_runs]
        c_wrapped = [_RunWrapper(r, run_text) for r in c_runs]
        matcher = difflib.SequenceMatcher(None, o_wrapped, c_wrapped)

        for tag, o_s, o_e, c_s, c_e in matcher.get_opcodes():
            if tag == 'equal':
                for oi, ci in zip(range(o_s, o_e), range(c_s, c_e)):
                    o_r, c_r = o_wrapped[oi].orig, c_wrapped[ci].orig
                    fmt_changes = {}
                    for attr in ('bold', 'italic', 'font', 'fontSize', 'underline', 'color'):
                        o_val = o_r.get(attr, '')
                        c_val = c_r.get(attr, '')
                        if o_val != c_val:
                            fmt_changes[attr] = {'from': o_val, 'to': c_val}
                    if fmt_changes:
                        diffs.append({
                            'text': c_r.get('text', '')[:20],
                            'changes': fmt_changes,
                            'text_also_changed': False,
                        })

        return diffs

    def _diff_table_cells(self, o_rows: List[List[dict]], c_rows: List[List[dict]]) -> List[dict]:
        """
        对比两个表格的单元格，生成 diff_cells。
        每人固定一个颜色（已由 author_colors 在写回时处理）。
        """
        diff_cells = []
        max_rows = max(len(o_rows), len(c_rows))
        max_cols = max(
            max(len(row) for row in o_rows) if o_rows else 0,
            max(len(row) for row in c_rows) if c_rows else 0,
        )

        for r in range(max_rows):
            for c in range(max_cols):
                o_cell = (o_rows[r][c] if r < len(o_rows) and c < len(o_rows[r]) else None)
                c_cell = (c_rows[r][c] if r < len(c_rows) and c < len(c_rows[r]) else None)

                o_text = (o_cell.get('text', '') or '' if isinstance(o_cell, dict) else '').strip()
                c_text = (c_cell.get('text', '') or '' if isinstance(c_cell, dict) else '').strip()

                if o_text != c_text:
                    diff_cells.append({
                        'row': r,
                        'col': c,
                        'old': o_text,
                        'new': c_text,
                    })

        return diff_cells


    # =====================================================================
    def _get_visual_width(self, s: str) -> int:
        """计算字符串在终端中的视觉宽度（全角/半角处理）"""
        width = 0
        for char in s:
            width += 2 if ord(char) > 0x7F else 1
        return width

    def _pad_visual(self, text: str, width: int) -> str:
        """视觉宽度对齐填充"""
        text = str(text)
        curr_w = self._get_visual_width(text)
        return text + ' ' * max(0, width - curr_w)

    def _print_detected_headings(self, label: str, anchors: List[Dict]):
        """
        纯视觉展示：直接输出已评估的矩阵列表，确保中文对齐。
        """
        print(f"\n  ═══ [{label} 识别的标题大纲] ({len(anchors)} 个) ═══")
        
        # 定义各列视觉宽度
        col_w = [4, 10, 10, 6, 6, 14, 6] # #, 原判, 修正, 评分, 视觉, 样式, 索引
        
        header = (
            f"  {self._pad_visual('#', col_w[0])} | "
            f"{self._pad_visual('原判级别', col_w[1])} | "
            f"{self._pad_visual('修正级别', col_w[2])} | "
            f"{self._pad_visual('评分', col_w[3])} | "
            f"{self._pad_visual('视觉分', col_w[4])} | "
            f"{self._pad_visual('样式名称', col_w[5])} | "
            f"{self._pad_visual('索引', col_w[6])} | 标题内容"
        )
        print(header)
        print(f"  {'-' * (sum(col_w) + 30)}")

        if not anchors:
            print(f"  (无标题)\n")
            return

        for i, anchor in enumerate(anchors):
            item = anchor.get('item')
            orig_level = f"L{anchor.get('level', '?')}"
            det_level = f"L{anchor.get('detected_level', 0)}"
            score = f"{anchor.get('heading_score', 0.0):.2f}"
            vw = f"{anchor.get('visual_weight', 0.0):.2f}"
            style = (item.style_name or 'Normal')[:12] if item else "Unknown"
            idx = str(anchor.get('item_index', '?'))
            
            # 推断级别缩进
            level_num = anchor.get('detected_level', 0)
            indent = '  ' * (level_num - 1) if level_num > 1 else ''
            
            text = anchor.get('content', '')
            display_text = (text[:55] + '...') if len(text) > 55 else text
            
            line = (
                f"  {self._pad_visual(i+1, col_w[0])} | "
                f"{self._pad_visual(orig_level, col_w[1])} | "
                f"{self._pad_visual(det_level, col_w[2])} | "
                f"{self._pad_visual(score, col_w[3])} | "
                f"{self._pad_visual(vw, col_w[4])} | "
                f"{self._pad_visual(style, col_w[5])} | "
                f"#{self._pad_visual(idx, col_w[6]-1)} | "
                f"{indent}{display_text}"
            )
            print(line)

        print(f"  {'-' * (sum(col_w) + 30)}\n")

    def _print_heading_structure(self, author, copy_anchors, orig_anchors=None, anchor_pairs=None, orphan_orig=None, orphan_copy=None):
        """
        严格降维后的纯视觉展示函数：直接打印底层的锚点配对关系，去除自身任何二次配对或推理逻辑。
        """
        if not copy_anchors:
            print(f"\n    [标题对齐表] (0 个标题) - 副本作者: {author}:")
            print(f"    *(未检测到标题样式)*\n")
            return

        if orig_anchors is not None:
            # 定义列宽
            col_id_w = 10
            col_txt_w = 45
            
            header = (
                f"    | {self._pad_visual('原索引', col_id_w)} | "
                f"{self._pad_visual('原稿标题', col_txt_w)} | "
                f"{self._pad_visual('副索引', col_id_w)} | "
                f"{self._pad_visual('副本标题', col_txt_w)} | "
                f"相似度 |"
            )
            # 关键修复：使用视觉宽度计算分割线长度
            visual_len = self._get_visual_width(header)
            sep = f"    {'-' * visual_len}"
            
            print(f"\n    [标题对齐表] (原稿 {len(orig_anchors)} / 副本 {len(copy_anchors)}) - 副本作者: {author}:")
            print(sep)
            print(header)
            print(sep)

            # 1. 打印成功配对
            if anchor_pairs:
                for (o_anchor, c_anchor, sim) in anchor_pairs:
                    o_l = f"[L{o_anchor.get('detected_level', 0)}]"
                    c_l = f"[L{c_anchor.get('detected_level', 0)}]"
                    o_text = f"{o_l} {o_anchor.get('content', '')}"
                    c_text = f"{c_l} {c_anchor.get('content', '')}"
                    
                    # 截断以防溢出
                    if self._get_visual_width(o_text) > col_txt_w:
                        o_text = o_text[:col_txt_w-3] + "..."
                    if self._get_visual_width(c_text) > col_txt_w:
                        c_text = c_text[:col_txt_w-3] + "..."
                    
                    line = (
                        f"    | {self._pad_visual('#' + str(o_anchor['item_index']), col_id_w)} | "
                        f"{self._pad_visual(o_text, col_txt_w)} | "
                        f"{self._pad_visual('#' + str(c_anchor['item_index']), col_id_w)} | "
                        f"{self._pad_visual(c_text, col_txt_w)} | "
                        f"{sim:.2f}   |"
                    )
                    print(line)

            # 2. 打印未配对的原稿（在副本中已消失）
            if orphan_orig:
                for o_anchor in orphan_orig:
                    o_l = f"[L{o_anchor.get('detected_level', 0)}]"
                    o_text = f"{o_l} {o_anchor.get('content', '')}"
                    if self._get_visual_width(o_text) > col_txt_w:
                        o_text = o_text[:col_txt_w-3] + "..."
                    
                    line = (
                        f"    | {self._pad_visual('#' + str(o_anchor['item_index']), col_id_w)} | "
                        f"{self._pad_visual(o_text, col_txt_w)} | "
                        f"{self._pad_visual('-', col_id_w)} | "
                        f"{self._pad_visual('<删除>', col_txt_w)} | "
                        f"-      |"
                    )
                    print(line)

            # 3. 打印未配对的副本（在副本中新增加）
            if orphan_copy:
                for c_anchor in orphan_copy:
                    c_l = f"[L{c_anchor.get('detected_level', 0)}]"
                    c_text = f"{c_l} {c_anchor.get('content', '')}"
                    if self._get_visual_width(c_text) > col_txt_w:
                        c_text = c_text[:col_txt_w-3] + "..."
                    
                    line = (
                        f"    | {self._pad_visual('-', col_id_w)} | "
                        f"{self._pad_visual('<新增>', col_txt_w)} | "
                        f"{self._pad_visual('#' + str(c_anchor['item_index']), col_id_w)} | "
                        f"{self._pad_visual(c_text, col_txt_w)} | "
                        f"-      |"
                    )
                    print(line)

            print(sep + "\n")
        else:
            # 仅显示独立副本的标题状态
            print(f"\n    [标题树概览] ({len(copy_anchors)} 个锚点) - 副本作者: {author}:")
            print(f"    {'-' * 100}")
            print(f"    | {'#':<4} | {'级别':<8} | {'样式':<15} | {'标题文本':<40} |")
            print(f"    {'-' * 100}")

            for i, anchor in enumerate(copy_anchors):
                text = anchor.get('content', '')[:38] if len(anchor.get('content', '')) > 38 else anchor.get('content', '')
                level = anchor.get('actual_level', '?')
                item = anchor.get('item')
                style = (item.style_name or 'Normal')[:13] if item else 'Unknown'
                print(f"    | {i+1:<4} | L{level:<7} | {style:<15} | {text:<40} |")

            print(f"    {'-' * 100}\n")

    def _get_section_path(self, item_index: int, items: List) -> List[str]:
        """获取 item 所在的章节路径"""
        section_path = []
        for i in range(item_index, -1, -1):
            item = items[i]
            if getattr(item, 'detected_level', 0) > 0:
                section_path.insert(0, item.content.strip())
                if len(section_path) >= 3:
                    break
        return section_path if section_path else ['(未分类)']

    def _estimate_page_number(self, item_index: int, items: List) -> int:
        """估算 item 所在页码"""
        if items and 0 <= item_index < len(items):
            item = items[item_index]
            if hasattr(item, 'page') and item.page:
                return item.page
        return (item_index // 30) + 1
    def _is_conflict(self, first: DocumentChange, second: DocumentChange) -> bool:
        if first.copy_index == second.copy_index or first.position != second.position:
            return False
        if any(change.change_type.endswith('insert') for change in (first, second)):
            return False
        if all(change.change_type.endswith('delete') for change in (first, second)):
            return False

        if first.change_type == second.change_type == 'table':
            first_format = first.format_info or {}
            second_format = second.format_info or {}
            if first_format.get('structure_changed') or second_format.get('structure_changed'):
                return True
            first_cells = {(cell['row'], cell['col']) for cell in first_format.get('diff_cells', [])}
            second_cells = {(cell['row'], cell['col']) for cell in second_format.get('diff_cells', [])}
            return bool(first_cells & second_cells)
        return True

    def _detect_conflicts(self):
        """检测多个副本对同一位置的冲突。"""
        all_changes = [change for changes in self.changes_by_author.values() for change in changes]
        all_changes.sort(key=lambda change: change.position)
        self.conflicts.clear()
        for position, group in groupby(all_changes, key=lambda change: change.position):
            same_position = list(group)
            conflict_ids = {
                change.change_id
                for first, second in combinations(same_position, 2)
                if self._is_conflict(first, second)
                for change in (first, second)
            }
            if not conflict_ids:
                continue

            changes = [change for change in same_position if change.change_id in conflict_ids]
            conflict = MergeConflict(
                conflict_id=f"conflict_{len(self.conflicts) + 1}",
                position=position,
                changes=changes,
                severity=self._assess_severity(changes),
                description=self._describe_conflict(changes),
            )
            self.conflicts.append(conflict)
            print(f"    - 检测到冲突: {conflict.conflict_id} ({len(changes)} 个变更)")

        if not self.conflicts:
            print("    - 未检测到冲突")

    def _merge_positions(self, positions: List[int]) -> int:
        """获取冲突位置的代表值"""
        if not positions:
            return 0
        return positions[0]
    
    def _assess_severity(self, changes: List[DocumentChange]) -> str:
        """评估冲突严重程度"""
        if len(changes) <= 2:
            return 'low'
        elif len(changes) <= 4:
            return 'medium'
        else:
            return 'high'
    
    def _describe_conflict(self, changes: List[DocumentChange]) -> str:
        """描述冲突"""
        authors = [c.author for c in changes]
        types = set(c.change_type for c in changes)
        
        return f"{len(changes)} 个编辑者 ({', '.join(authors)}) 在同一位置进行了 {', '.join(types)} 操作"

    def _write_changes(self, output_path: str) -> int:
        """写回操作：执行合并写入"""
        all_changes = []
        for changes in self.changes_by_author.values():
            all_changes.extend(changes)

        # ── 关键映射：items 索引 → body 直接子元素索引 ──────────────────────
        items_idx_to_block_idx: Dict[int, int] = {}
        current_block_idx = -1
        for i, item in enumerate(self.orig_items):
            if item.type in ('paragraph', 'table', 'sdt'):
                current_block_idx += 1
            items_idx_to_block_idx[i] = current_block_idx

        # 给每个 change 附加 true_block_index
        for change in all_changes:
            if change.position in items_idx_to_block_idx:
                setattr(change, 'true_block_index', items_idx_to_block_idx[change.position])
            else:
                setattr(change, 'true_block_index', change.position)

        # 同一位置先插入，再标记原文删除，最后写入修改版本。
        def write_order(change):
            operation = change.change_type.rsplit('_', 1)[-1]
            priority = {'insert': 2, 'delete': 1}.get(operation, 0)
            return change.true_block_index, priority, change.copy_index

        all_changes_sorted = sorted(all_changes, key=write_order, reverse=True)
        print(f"    - 待写入变更: {len(all_changes_sorted)} 条")

        # 执行写回
        self.builder.merge_documents_with_revisions(
            self.original_doc,
            self.copies,
            output_path,
            changes=all_changes_sorted,
            original_para_by_order=self._original_para_by_order,
            original_para_anchor_map=self._original_para_anchor_map,
            conflicting_changes_by_para={conflict.position: conflict.changes for conflict in self.conflicts}
        )
        print(f"    ✓ 写回完成: {output_path}")
        return len(all_changes_sorted)

    def _generate_conflict_report(self, report_path: str):
        """生成冲突报告 (TXT 格式)"""
        self._generate_conflict_report_txt(report_path)
        print(f"    冲突报告已生成: {report_path}")

    def _generate_conflict_report_txt(self, report_path: str):
        """生成 TXT 格式的冲突报告"""
        try:
            # 收集冲突变更 ID
            conflict_change_ids = set()
            for c in self.conflicts:
                for change in c.changes:
                    conflict_change_ids.add(change.change_id)

            all_changes = []
            for changes in self.changes_by_author.values():
                all_changes.extend(changes)
            all_changes.sort(key=lambda c: (c.page_number, c.paragraph_index))

            with open(report_path, 'w', encoding='utf-8') as f:
                # 标题
                f.write("=" * 70 + "\n")
                f.write("           文档合并冲突报告\n")
                f.write("=" * 70 + "\n")
                f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"原稿文件: {os.path.basename(self.original_doc)}\n")
                f.write(f"副本数量: {len(self.copies)}\n")
                f.write(f"冲突数量: {len(self.conflicts)}\n")
                f.write(f"冲突变更总数: {sum(len(c.changes) for c in self.conflicts)}\n")
                f.write(f"自动合并变更数: {len(all_changes) - sum(len(c.changes) for c in self.conflicts)}\n")
                f.write("\n")

                # 1. 冲突详情
                f.write("\n" + "=" * 70 + "\n")
                f.write("1. 冲突详情\n")
                f.write("=" * 70 + "\n\n")

                if not self.conflicts:
                    f.write("未发现冲突，所有变更已自动合并。\n")
                else:
                    for conflict in self.conflicts:
                        f.write(f"[冲突 ID] {conflict.conflict_id} | 严重程度: {conflict.severity}\n")
                        f.write(f"描述: {conflict.description}\n")
                        f.write("-" * 50 + "\n")

                        for change in conflict.changes:
                            loc = f"P{change.page_number}/段落{change.paragraph_index + 1}"
                            section = ' > '.join(change.section_path) if change.section_path else '未分类'
                            content_preview = change.content[:200].replace('\n', ' ').replace('\r', '')
                            f.write(f"  * {change.author} ({change.change_type}) @ {loc} [{section}]\n")
                            f.write(f"    内容: {content_preview}...\n")
                        f.write("\n")

                # 2. 自动合并记录
                f.write("\n" + "=" * 70 + "\n")
                f.write("2. 自动合并记录\n")
                f.write("=" * 70 + "\n\n")

                count = 0
                for change in all_changes:
                    if change.change_id not in conflict_change_ids:
                        count += 1
                        section = ' > '.join(change.section_path) if change.section_path else '未分类'
                        content_preview = change.content[:200].replace('\n', ' ').replace('\r', '')
                        type_label_map = {
                            'delete': '删除', 'insert': '新增',
                            'modify': '修改', 'table': '表格修改',
                            'table_insert': '表格新增', 'table_delete': '表格删除',
                        }
                        type_label = type_label_map.get(change.change_type, change.change_type.upper())
                        f.write(f"[{count}] {change.author} | {type_label} | "
                                f"P{change.page_number}/段落{change.paragraph_index + 1} | {section}\n")
                        f.write(f"    内容: {content_preview}\n\n")

                if count == 0:
                    f.write("无额外的自动合并记录。\n")

            print(f"✓ 冲突报告已生成: {report_path}")

        except Exception as e:
            raise RuntimeError(f"无法写入冲突报告 {report_path}: {e}") from e
    
    def _create_conflict_report_html(self) -> str:
        """创建冲突报告 HTML"""
        html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>文档合并冲突报告</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 { font-size: 32px; margin-bottom: 10px; }
        .header p { font-size: 16px; opacity: 0.9; }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
        }
        .stat-number { font-size: 36px; font-weight: bold; color: #667eea; margin-bottom: 5px; }
        .stat-label { color: #6c757d; font-size: 14px; }
        .conflicts { padding: 30px; }
        .conflict {
            background: white;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            margin-bottom: 20px;
            overflow: hidden;
            transition: all 0.3s;
        }
        .conflict:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-color: #667eea; }
        .conflict-header {
            background: #f8f9fa;
            padding: 15px 20px;
            border-bottom: 1px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .conflict-id { font-weight: bold; color: #495057; }
        .severity {
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: bold;
            text-transform: uppercase;
        }
        .severity-low { background: #d4edda; color: #155724; }
        .severity-medium { background: #fff3cd; color: #856404; }
        .severity-high { background: #f8d7da; color: #721c24; }
        .conflict-body { padding: 20px; }
        .change {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 4px;
        }
        .change-header { display: flex; justify-content: space-between; margin-bottom: 10px; }
        .author { font-weight: bold; color: #667eea; }
        .change-type {
            background: #667eea;
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 12px;
        }
        .location { color: #6c757d; font-size: 14px; margin-bottom: 10px; }
        .content {
            background: white;
            padding: 10px;
            border-radius: 4px;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        .breadcrumb { color: #6c757d; font-size: 13px; margin-top: 5px; }
        .breadcrumb span { margin: 0 5px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 文档合并冲突报告</h1>
            <p>生成时间: """ + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + """</p>
        </div>
        <div class="summary">
            <div class="stat-card">
                <div class="stat-number">""" + str(len(self.conflicts)) + """</div>
                <div class="stat-label">冲突数量</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">""" + str(len(self.copies)) + """</div>
                <div class="stat-label">副本数量</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">""" + str(sum(len(c.changes) for c in self.conflicts)) + """</div>
                <div class="stat-label">冲突变更</div>
            </div>
        </div>
        <div class="conflicts">
"""
        
        if not self.conflicts:
            html += """
            <div style="text-align: center; padding: 40px; color: #6c757d;">
                <h2>🎉 没有发现冲突!</h2>
                <p>所有变更已自动合并</p>
            </div>
"""
        else:
            for conflict in self.conflicts:
                html += self._create_conflict_html(conflict)
        
        html += """
        </div>
        
        <div class="header" style="margin-top: 40px; background: linear-gradient(135deg, #28a745 0%, #20c997 100%);">
            <h1>✅ 自动合并记录</h1>
            <p>以下变更是安全的，已自动应用到最终文档</p>
        </div>
        <div class="conflicts">
"""
        
        # 收集所有参与冲突的变更ID
        conflict_change_ids = set()
        for c in self.conflicts:
            for change in c.changes:
                conflict_change_ids.add(change.change_id)
        
        # 收集所有变更并排序
        all_changes = []
        for changes in self.changes_by_author.values():
            all_changes.extend(changes)
        all_changes.sort(key=lambda c: (c.page_number, c.paragraph_index))
        
        auto_merged_count = 0
        for change in all_changes:
            if change.change_id not in conflict_change_ids:
                auto_merged_count += 1
                section_path = ' → '.join(change.section_path) if change.section_path else '(未分类)'
                clean_content = self._escape_html(change.content)
                if len(clean_content) > 100:
                    clean_content = clean_content[:100] + "..."
                
                html += f"""
            <div class="conflict" style="border-left: 4px solid #28a745;">
                <div class="conflict-header">
                    <span class="conflict-id">自动合并</span>
                    <span class="severity severity-low">SUCCESS</span>
                </div>
                <div class="conflict-body" style="padding: 15px;">
                     <div class="change-header">
                        <span class="author">👤 {change.author}</span>
                        <span class="change-type" style="background: #28a745;">{change.change_type}</span>
                    </div>
                    <div class="location">
                        📍 第 {change.page_number} 页, 段落 {change.paragraph_index + 1} | 📂 {section_path}
                    </div>
                    <div class="content">{clean_content}</div>
                </div>
            </div>
"""

        if auto_merged_count == 0:
             html += """
            <div style="text-align: center; padding: 20px; color: #6c757d;">
                <p>没有其他自动合并的变更</p>
            </div>
"""

        html += """
        </div>
    </div>
</body>
</html>
"""
        return html
    
    def _create_conflict_html(self, conflict: MergeConflict) -> str:
        """创建单个冲突的 HTML"""
        html = f"""
        <div class="conflict">
            <div class="conflict-header">
                <span class="conflict-id">{conflict.conflict_id}</span>
                <span class="severity severity-{conflict.severity}">{conflict.severity}</span>
            </div>
            <div class="conflict-body">
                <p style="margin-bottom: 15px; color: #495057;">{conflict.description}</p>
"""
        
        for change in conflict.changes:
            section_path = ' → '.join(change.section_path) if change.section_path else '(未分类)'
            
            html += f"""
                <div class="change">
                    <div class="change-header">
                        <span class="author">👤 {change.author}</span>
                        <span class="change-type">{change.change_type}</span>
                    </div>
                    <div class="location">
                        📍 位置: 第 {change.page_number} 页, 段落 {change.paragraph_index + 1}
                    </div>
                    <div class="breadcrumb">
                        📂 章节: {section_path}
                    </div>
                    <div class="content">{self._escape_html(change.content[:200])}...</div>
"""
            
            if change.table_info:
                html += f"""
                    <div style="margin-top: 10px; padding: 10px; background: #e7f3ff; border-radius: 4px;">
                        📊 表格变更: {change.table_info}
                    </div>
"""
            
            html += """
                </div>
"""
        
        html += """
            </div>
        </div>
"""
        return html
    
    def _escape_html(self, text: str) -> str:
        """转义 HTML 特殊字符"""
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#39;'))



