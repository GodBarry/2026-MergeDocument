#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文档自动合并脚本
基于 MultiCopyMerger 系统，支持命令行参数配置。

用法示例:
    python run_merge.py original.docx ./copies/ -o merged.docx --authors "Alice" "None" "Bob"
    
如果不指定 --authors，或者对应位置为 None，脚本将尝试从 .docx 文件的元数据中提取作者名。
"""

import argparse
import os
import io
import sys

# 解决 Windows 控制台乱码问题
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import glob
import zipfile
import subprocess
import shutil
import xml.etree.ElementTree as ET
from typing import Optional, List
from datetime import datetime
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from .multi_copy_merger import MultiCopyMerger


# 安全 XML 解析：处理 BOM 头和编码声明不一致的情况
def safe_parse_xml(xml_path: str):
    """
    安全解析 XML 文件，处理以下常见问题：
    1. 动态注册所有 xmlns，避免 ET.parse 重写成 ns0 前缀导致 Word 崩溃报错
    2. UTF-8 BOM 头 (EF BB BF) 导致 ET.parse 失败
    3. 编码声明与实际编码不一致
    4. 文件头包含垃圾字节
    """
    import re
    try:
        with open(xml_path, 'r', encoding='utf-8', errors='ignore') as fh:
            head_content = fh.read(8192)
        for m in re.finditer(r'xmlns:([a-zA-Z0-9_\-]+)="([^"]+)"', head_content):
            try:
                ET.register_namespace(m.group(1), m.group(2))
            except ValueError:
                pass
    except Exception:
        pass

    try:
        return ET.parse(xml_path)
    except ET.ParseError:
        # 回退：读取内容，去除 BOM，再解析
        try:
            with open(xml_path, 'rb') as f:
                raw = f.read()
            # 去除 UTF-8 BOM
            if raw.startswith(b'\xef\xbb\xbf'):
                raw = raw[3:]
            # 去除 UTF-16 LE/BE BOM
            if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
                text = raw.decode('utf-16')
                raw = text.encode('utf-8')
            # 修正编码声明
            content = raw.decode('utf-8', errors='replace')
            # 替换可能的错误编码声明
            content = re.sub(
                r'<\?xml[^>]*encoding="[^"]*"',
                '<?xml version="1.0" encoding="UTF-8"',
                content,
                count=1
            )
            root = ET.fromstring(content)
            # 创建 ElementTree 并返回
            tree = ET.ElementTree(root)
            return tree
        except Exception:
            raise


def inject_comments_to_docx(docx_path: str, all_changes: list, context_window: int = 3):
    """
    将修改者信息以标准 Word 批注的形式注入到合并后的 DOCX 文件中。
    直接操作 OOXML，不依赖 ONLYOFFICE JS API。
    通过精确的段落文本匹配或者后置位索引兜底，确保批注插入在正确的段落位置。
    """
    import tempfile, shutil, re, time
    from datetime import datetime

    print(f"\n正在注入修改批注 ({len(all_changes)} 条): {os.path.basename(docx_path)}")

    if not all_changes:
        print("  (无变更)")
        return
    if not zipfile.is_zipfile(docx_path):
        print("  ! 文件不是有效 DOCX，跳过")
        return

    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    wtag = lambda t: f'{{{W}}}{t}'

    temp_dir = tempfile.mkdtemp()
    temp_output = docx_path + '.comment.tmp'

    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            zf.extractall(temp_dir)

        word_dir = os.path.join(temp_dir, 'word')
        doc_path = os.path.join(word_dir, 'document.xml')
        comm_path = os.path.join(word_dir, 'comments.xml')
        rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')
        ctype_path = os.path.join(temp_dir, '[Content_Types].xml')

        if not os.path.exists(doc_path):
            return

        ET.register_namespace('w', W)
        ns_map = {
            'w': W,
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        }
        for prefix, uri in ns_map.items():
            ET.register_namespace(prefix, uri)

        tree = safe_parse_xml(doc_path)
        root = tree.getroot()

        # 找所有段落
        all_paras = list(root.iter(wtag('p')))

        # 预构建所有段落的纯文本索引
        para_text_map = {}
        para_norm_text_map = {}

        def _norm(t):
            return re.sub(r'\s+', '', t)

        for p in all_paras:
            runs_in_p = p.findall(wtag('r'))
            if not runs_in_p:
                continue
            p_text = ''
            for r in runs_in_p:
                t_el = r.find(wtag('t'))
                if t_el is not None and t_el.text:
                    p_text += t_el.text
            if p_text:
                if p_text not in para_text_map:
                    para_text_map[p_text] = []
                para_text_map[p_text].append(p)

                norm_text = _norm(p_text)
                if norm_text:
                    if norm_text not in para_norm_text_map:
                        para_norm_text_map[norm_text] = []
                    para_norm_text_map[norm_text].append(p)

        now_iso = datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
        existing_max_id = 0

        if os.path.exists(comm_path):
            try:
                comm_tree_for_id = safe_parse_xml(comm_path)
                comm_root_for_id = comm_tree_for_id.getroot()
                for c in list(comm_root_for_id):
                    cid = c.get(f'{{{W}}}id') or c.get('w:id') or c.get('id')
                    if cid is not None and str(cid).isdigit():
                        existing_max_id = max(existing_max_id, int(cid))
            except Exception:
                try:
                    with open(comm_path, 'r', encoding='utf-8', errors='replace') as f:
                        old_comm = f.read()
                    for mid in re.findall(r'w:id="(\d+)"', old_comm):
                        existing_max_id = max(existing_max_id, int(mid))
                except Exception:
                    pass

        next_id = existing_max_id + 1
        comment_entries = []

        for change in all_changes:
            ctype = getattr(change, 'change_type', '')
            content = getattr(change, 'content', '')
            para_idx = getattr(change, 'paragraph_index', None)
            author = getattr(change, 'author', '未知')

            ctype_desc_map = {
                'insert': ('新增', '此段落为副本新增内容'),
                'delete': ('删除', '原稿中此段落已被副本删除'),
                'modify': ('修改', '此段落在副本中被修改'),
                'table_insert': ('新增表格', '此表格为副本新增'),
                'table_replace': ('表格修改', '此表格在副本中被修改'),
                'table_delete': ('删除表格', '原稿中此表格已被副本删除'),
            }
            ctype_label, ctype_note = ctype_desc_map.get(ctype, (ctype, ''))
            label = f"{author} [{ctype_label}]"
            if ctype_note:
                label += f" - {ctype_note}"

            diff_text = ''
            if ctype == 'insert':
                diff_text = content
            elif ctype == 'modify':
                try:
                    import json
                    diff_parts = json.loads(content)
                    diff_text = ''.join([p.get('text', '') for p in diff_parts if p.get('op') != 'delete'])
                except (json.JSONDecodeError, TypeError, AttributeError):
                    diff_text = ''

            target_p = None
            match_method = ''

            if diff_text:
                norm_diff = _norm(diff_text)
                matches = para_text_map.get(diff_text)
                if matches:
                    target_p = matches.pop(0)
                    if not matches:
                        del para_text_map[diff_text]
                    if norm_diff in para_norm_text_map and target_p in para_norm_text_map[norm_diff]:
                        para_norm_text_map[norm_diff].remove(target_p)
                    match_method = '精确匹配'

                if target_p is None and norm_diff:
                    matches = para_norm_text_map.get(norm_diff)
                    if matches:
                        target_p = matches.pop(0)
                        if not matches:
                            del para_norm_text_map[norm_diff]
                        match_method = '忽略空白匹配'

            if target_p is None:
                if para_idx is not None and 0 <= para_idx < len(all_paras):
                    target_p = all_paras[para_idx]
                    match_method = f'索引兜底(idx={para_idx})'

            if target_p is None:
                print(f"    ⚠ 批注跳过: [{ctype}] 无法定位段落 (para_idx={para_idx})")
                continue

            if next_id <= existing_max_id + 10 or (next_id - existing_max_id) % 25 == 0:
                print(f"    批注#{next_id} [{ctype}] -> {match_method}")
            comment_entries.append((next_id, target_p, author, label))
            next_id += 1

        if not comment_entries:
            print("  (无需注入批注)")
            return

        # 在 XML 中插入注释元素
        for c_id, target_p, author, label in comment_entries:
            pPr_idx = -1
            for i, child in enumerate(target_p):
                if child.tag == wtag('pPr'):
                    pPr_idx = i
                    break
            insert_idx = pPr_idx + 1 if pPr_idx >= 0 else 0

            c_start = ET.Element(wtag('commentRangeStart'), {f'{{{W}}}id': str(c_id)})
            target_p.insert(insert_idx, c_start)

            c_end = ET.Element(wtag('commentRangeEnd'), {f'{{{W}}}id': str(c_id)})
            target_p.append(c_end)

            r_el = ET.Element(wtag('r'))
            rPr_el = ET.SubElement(r_el, wtag('rPr'))
            ET.SubElement(rPr_el, wtag('rStyle'), {f'{{{W}}}val': 'CommentReference'})
            ET.SubElement(r_el, wtag('commentReference'), {f'{{{W}}}id': str(c_id)})
            target_p.append(r_el)

        tree.write(doc_path, encoding='utf-8', xml_declaration=True)

        # 构建 comments.xml
        WPC = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
        MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
        W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
        ET.register_namespace('wpc', WPC)
        ET.register_namespace('mc', MC)
        ET.register_namespace('r', R)
        ET.register_namespace('m', M)
        ET.register_namespace('w14', W14)

        if os.path.exists(comm_path):
            try:
                comm_tree = safe_parse_xml(comm_path)
                comm_root = comm_tree.getroot()
            except Exception:
                comm_root = ET.Element(wtag('comments'))
                comm_tree = ET.ElementTree(comm_root)
        else:
            comm_root = ET.Element(wtag('comments'))
            comm_tree = ET.ElementTree(comm_root)

        ign_key = f'{{{MC}}}Ignorable'
        if ign_key in comm_root.attrib:
            try:
                del comm_root.attrib[ign_key]
            except Exception:
                pass

        def _clean_xml_text(s: str) -> str:
            if s is None:
                return ''
            return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', str(s))

        for c_id, target_p, author, label in comment_entries:
            cmt = ET.SubElement(comm_root, wtag('comment'))
            cmt.set(f'{{{W}}}id', str(c_id))
            cmt.set(f'{{{W}}}author', _clean_xml_text(author))
            cmt.set(f'{{{W}}}date', now_iso)

            p = ET.SubElement(cmt, wtag('p'))
            pPr = ET.SubElement(p, wtag('pPr'))
            ET.SubElement(pPr, wtag('spacing'), {f'{{{W}}}line': '240', f'{{{W}}}after': '0', f'{{{W}}}lineRule': 'auto', f'{{{W}}}before': '0'})
            r_el2 = ET.SubElement(p, wtag('r'))
            rPr2 = ET.SubElement(r_el2, wtag('rPr'))
            ET.SubElement(rPr2, wtag('rStyle'), {f'{{{W}}}val': 'CommentText'})
            t_el2 = ET.SubElement(r_el2, wtag('t'))
            t_el2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t_el2.text = _clean_xml_text(label)

        comm_tree.write(comm_path, encoding='utf-8', xml_declaration=True)

        # 更新 document.xml.rels
        if os.path.exists(rels_path):
            try:
                RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
                rtag = lambda t: f'{{{RELNS}}}{t}'
                ET.register_namespace('', RELNS)
                rel_tree = safe_parse_xml(rels_path)
                rel_root = rel_tree.getroot()
                has_comments_rel = False
                used_ids = set()
                for rel in rel_root.findall(rtag('Relationship')):
                    rid = rel.get('Id', '')
                    if rid:
                        used_ids.add(rid)
                    typ = rel.get('Type', '')
                    tgt = rel.get('Target', '')
                    if (typ.endswith('/comments') or 'comments' in typ) and tgt == 'comments.xml':
                        has_comments_rel = True
                if not has_comments_rel:
                    max_n = 0
                    for rid in used_ids:
                        m = re.match(r'rId(\d+)$', rid)
                        if m:
                            max_n = max(max_n, int(m.group(1)))
                    new_rid = f"rId{max_n + 1}" if max_n > 0 else "rIdComments"
                    while new_rid in used_ids:
                        max_n += 1
                        new_rid = f"rId{max_n + 1}"
                    rel_el = ET.Element(rtag('Relationship'))
                    rel_el.set('Id', new_rid)
                    rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
                    rel_el.set('Target', 'comments.xml')
                    rel_root.append(rel_el)
                    rel_tree.write(rels_path, encoding='utf-8', xml_declaration=True)
            except Exception as _rels_e:
                print(f"  ⚠️ 更新 document.xml.rels 失败: {_rels_e}")

        # 更新 [Content_Types].xml
        if os.path.exists(ctype_path):
            try:
                CTNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
                cttag = lambda t: f'{{{CTNS}}}{t}'
                ET.register_namespace('', CTNS)
                ct_tree = safe_parse_xml(ctype_path)
                ct_root = ct_tree.getroot()
                has_override = False
                for ov in ct_root.findall(cttag('Override')):
                    if ov.get('PartName') == '/word/comments.xml':
                        has_override = True
                        ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
                        break
                if not has_override:
                    ov = ET.Element(cttag('Override'))
                    ov.set('PartName', '/word/comments.xml')
                    ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
                    ct_root.append(ov)
                ct_tree.write(ctype_path, encoding='utf-8', xml_declaration=True)
            except Exception as _ct_e:
                print(f"  ⚠️ 更新 [Content_Types].xml 失败: {_ct_e}")

        # 重新打包
        with zipfile.ZipFile(temp_output, 'w', zipfile.ZIP_DEFLATED) as zf:
            mimetype_path = os.path.join(temp_dir, 'mimetype')
            if os.path.exists(mimetype_path):
                zf.write(mimetype_path, 'mimetype', compress_type=zipfile.ZIP_STORED)
            for root_dir, _, files in os.walk(temp_dir):
                for fname in files:
                    if fname == 'mimetype':
                        continue
                    fp = os.path.join(root_dir, fname)
                    arc = os.path.relpath(fp, temp_dir)
                    zf.write(fp, arc, compress_type=zipfile.ZIP_DEFLATED)

        replaced = False
        last_err = None
        for _ in range(6):
            try:
                os.replace(temp_output, docx_path)
                replaced = True
                break
            except Exception as _re:
                last_err = _re
                time.sleep(0.2)
        if not replaced:
            raise last_err
        print(f"  ✓ 批注注入完成: 共 {len(comment_entries)} 条批注")

    except Exception as e:
        print(f"  ⚠️ 批注注入失败: {e}")
        import traceback; traceback.print_exc()
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if os.path.exists(temp_output):
            try: os.remove(temp_output)
            except Exception: pass


def inject_corrections_to_docx(docx_path: str, corrections: list):
    """
    将文本校对结果以批注形式注入到合并后的 DOCX 文件中（纯 Python XML 方式）。
    不依赖 ONLYOFFICE，通过直接操作 OOXML 实现。

    标注策略：对于每个有修正的段落，在段落后方插入一条批注，
    说明原文错误内容和修正后的内容。
    """
    import tempfile, re, time
    from datetime import datetime

    print(f"\n正在注入校对标注 ({len(corrections)} 条): {os.path.basename(docx_path)}")

    if not corrections:
        return
    if not zipfile.is_zipfile(docx_path):
        print("  ! 文件不是有效 DOCX，跳过")
        return

    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    wtag = lambda t: f'{{{W}}}{t}'

    temp_dir = tempfile.mkdtemp()
    temp_output = docx_path + '.correction.tmp'

    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            zf.extractall(temp_dir)

        word_dir = os.path.join(temp_dir, 'word')
        doc_path = os.path.join(word_dir, 'document.xml')
        comm_path = os.path.join(word_dir, 'comments.xml')
        rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')
        ctype_path = os.path.join(temp_dir, '[Content_Types].xml')

        if not os.path.exists(doc_path):
            print("  ! document.xml 不存在，跳过")
            return

        ET.register_namespace('w', W)
        ns_map = {
            'w': W,
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        }
        for prefix, uri in ns_map.items():
            ET.register_namespace(prefix, uri)

        tree = safe_parse_xml(doc_path)
        root = tree.getroot()
        body = root.find(wtag('body'))
        if body is None:
            print("  ! 未找到 body，跳过")
            return

        # 收集所有段落元素（按 body 中的顺序）
        block_tags = {wtag('p'), wtag('tbl'), wtag('sdt')}
        body_children = [c for c in body if c.tag in block_tags]

        now_iso = datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')

        # 建立段落文本 -> 元素索引的映射（按原始顺序）
        para_text_to_idx = {}
        for i, child in enumerate(body_children):
            if child.tag == wtag('p'):
                texts = [t.text for t in child.iter(wtag('t')) if t.text]
                text_key = ''.join(texts).strip()
                if text_key:
                    if text_key not in para_text_to_idx:
                        para_text_to_idx[text_key] = []
                    para_text_to_idx[text_key].append(i)

        # 处理 corrections：将修正注入到对应段落
        matched_corrections = []
        used_keys = set()  # 已匹配的段落文本

        for corr in corrections:
            orig_text = corr.get('original', '').strip()
            corrected_text = corr.get('corrected', '').strip()
            details = corr.get('details', [])
            page = corr.get('page', 0)
            idx = corr.get('index', -1)

            # 尝试通过原文精确定位段落
            matched_para_idx = -1
            match_method = ''

            if orig_text in para_text_to_idx and orig_text not in used_keys:
                # 精确匹配
                matched_para_idx = para_text_to_idx[orig_text][0]
                used_keys.add(orig_text)
                match_method = '精确文本匹配'
            elif idx >= 0 and idx < len(body_children):
                # 索引兜底
                matched_para_idx = idx
                match_method = f'索引兜底(idx={idx})'

            if matched_para_idx >= 0:
                # 检查是否已经匹配过
                if matched_para_idx in [m['para_idx'] for m in matched_corrections]:
                    # 如果同一段落被多次修正，合并到同一条批注
                    for m in matched_corrections:
                        if m['para_idx'] == matched_para_idx:
                            # 向已有批注追加内容
                            detail_lines = []
                            for d in details:
                                detail_lines.append(f"  - [{d.get('type', '未知')}] "
                                    f"「{d.get('original', '')}」→「{d.get('corrected', '')}」"
                                    f" ({d.get('explanation', '')})")
                            m['detail_lines'].extend(detail_lines)
                            break
                else:
                    detail_lines = []
                    for d in details:
                        detail_lines.append(f"  - [{d.get('type', '未知')}] "
                            f"「{d.get('original', '')}」→「{d.get('corrected', '')}」"
                            f" ({d.get('explanation', '')})")
                    matched_corrections.append({
                        'para_idx': matched_para_idx,
                        'match_method': match_method,
                        'original': orig_text,
                        'corrected': corrected_text,
                        'detail_lines': detail_lines,
                        'page': page,
                    })

        if not matched_corrections:
            print("  ! 无法将校对标注匹配到文档段落，跳过")
            return

        print(f"  成功匹配 {len(matched_corrections)} 处标注")

        # 读取或创建 comments.xml
        if os.path.exists(comm_path):
            try:
                comm_tree = safe_parse_xml(comm_path)
                comm_root = comm_tree.getroot()
            except Exception:
                comm_root = ET.Element(wtag('comments'))
                comm_tree = ET.ElementTree(comm_root)
        else:
            comm_root = ET.Element(wtag('comments'))
            comm_tree = ET.ElementTree(comm_root)
            ET.register_namespace('w', W)
            ET.register_namespace('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas')
            ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')

        # 查找现有最大 comment id
        existing_max_id = 0
        for c_el in comm_root.findall(wtag('comment')):
            cid_str = c_el.get(f'{{{W}}}id') or c_el.get('w:id') or c_el.get('id')
            if cid_str and cid_str.isdigit():
                existing_max_id = max(existing_max_id, int(cid_str))

        def _clean_xml_text(s: str) -> str:
            if s is None:
                return ''
            # 移除 XML 非法字符
            s = str(s)
            s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', s)
            return s

        comment_entries = []

        for i_corr, corr_item in enumerate(matched_corrections):
            para_idx = corr_item['para_idx']
            if para_idx < 0 or para_idx >= len(body_children):
                print(f"    ⚠ 标注跳过: para_idx={para_idx} 超出范围")
                continue

            target_para = body_children[para_idx]
            next_id = existing_max_id + 1 + i_corr

            # 构建批注文本
            label_lines = [
                f"[文本校对] 第 {corr_item['page']} 页",
                f"原文: {corr_item['original']}",
                f"修正: {corr_item['corrected']}",
            ]
            label_lines.extend(corr_item['detail_lines'])

            # 在段落末尾插入 commentRangeStart / End
            pPr_idx = -1
            for j, child in enumerate(target_para):
                if child.tag == wtag('pPr'):
                    pPr_idx = j
                    break
            insert_idx = pPr_idx + 1 if pPr_idx >= 0 else len(target_para)

            c_start = ET.Element(wtag('commentRangeStart'), {f'{{{W}}}id': str(next_id)})
            target_para.insert(insert_idx, c_start)

            c_end = ET.Element(wtag('commentRangeEnd'), {f'{{{W}}}id': str(next_id)})
            target_para.append(c_end)

            r_el = ET.Element(wtag('r'))
            rPr_el = ET.SubElement(r_el, wtag('rPr'))
            ET.SubElement(rPr_el, wtag('rStyle'), {f'{{{W}}}val': 'CommentReference'})
            ET.SubElement(r_el, wtag('commentReference'), {f'{{{W}}}id': str(next_id)})
            target_para.append(r_el)

            # 创建 comments.xml 条目
            cmt = ET.SubElement(comm_root, wtag('comment'))
            cmt.set(f'{{{W}}}id', str(next_id))
            cmt.set(f'{{{W}}}author', '文本校对')
            cmt.set(f'{{{W}}}date', now_iso)
            cmt.set(f'{{{W}}}initials', 'TC')

            for line in label_lines:
                p_el = ET.SubElement(cmt, wtag('p'))
                pPr_el = ET.SubElement(p_el, wtag('pPr'))
                ET.SubElement(pPr_el, wtag('spacing'),
                              {f'{{{W}}}line': '240', f'{{{W}}}after': '0',
                               f'{{{W}}}lineRule': 'auto', f'{{{W}}}before': '0'})
                r_el2 = ET.SubElement(p_el, wtag('r'))
                rPr2 = ET.SubElement(r_el2, wtag('rPr'))
                ET.SubElement(rPr2, wtag('rStyle'), {f'{{{W}}}val': 'CommentText'})
                t_el2 = ET.SubElement(r_el2, wtag('t'))
                t_el2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t_el2.text = _clean_xml_text(line)

            comment_entries.append((next_id, corr_item['match_method']))

        # 写回 document.xml
        tree.write(doc_path, encoding='utf-8', xml_declaration=True)

        # 写回 comments.xml
        comm_tree.write(comm_path, encoding='utf-8', xml_declaration=True)

        # 更新 document.xml.rels
        if os.path.exists(rels_path):
            try:
                RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
                rtag = lambda t: f'{{{RELNS}}}{t}'
                ET.register_namespace('', RELNS)
                rel_tree = safe_parse_xml(rels_path)
                rel_root = rel_tree.getroot()
                has_comments_rel = False
                used_ids = set()
                for rel in rel_root.findall(rtag('Relationship')):
                    rid = rel.get('Id', '')
                    if rid:
                        used_ids.add(rid)
                    typ = rel.get('Type', '')
                    tgt = rel.get('Target', '')
                    if (typ.endswith('/comments') or 'comments' in typ) and tgt == 'comments.xml':
                        has_comments_rel = True
                if not has_comments_rel:
                    max_n = 0
                    for rid in used_ids:
                        m = re.match(r'rId(\d+)$', rid)
                        if m:
                            max_n = max(max_n, int(m.group(1)))
                    new_rid = f"rId{max_n + 1}"
                    rel_el = ET.Element(rtag('Relationship'))
                    rel_el.set('Id', new_rid)
                    rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments')
                    rel_el.set('Target', 'comments.xml')
                    rel_root.append(rel_el)
                    rel_tree.write(rels_path, encoding='utf-8', xml_declaration=True)
            except Exception as _rels_e:
                print(f"  ⚠ 更新 document.xml.rels 失败: {_rels_e}")

        # 更新 [Content_Types].xml
        if os.path.exists(ctype_path):
            try:
                CTNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
                cttag = lambda t: f'{{{CTNS}}}{t}'
                ET.register_namespace('', CTNS)
                ct_tree = safe_parse_xml(ctype_path)
                ct_root = ct_tree.getroot()
                has_override = False
                for ov in ct_root.findall(cttag('Override')):
                    if ov.get('PartName') == '/word/comments.xml':
                        has_override = True
                        ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
                        break
                if not has_override:
                    ov = ET.Element(cttag('Override'))
                    ov.set('PartName', '/word/comments.xml')
                    ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
                    ct_root.append(ov)
                ct_tree.write(ctype_path, encoding='utf-8', xml_declaration=True)
            except Exception as _ct_e:
                print(f"  ⚠ 更新 [Content_Types].xml 失败: {_ct_e}")

        # 重新打包
        with zipfile.ZipFile(temp_output, 'w', zipfile.ZIP_DEFLATED) as zf:
            mimetype_path = os.path.join(temp_dir, 'mimetype')
            if os.path.exists(mimetype_path):
                zf.write(mimetype_path, 'mimetype', compress_type=zipfile.ZIP_STORED)
            for root_dir, _, files in os.walk(temp_dir):
                for fname in files:
                    if fname == 'mimetype':
                        continue
                    fp = os.path.join(root_dir, fname)
                    arc = os.path.relpath(fp, temp_dir)
                    zf.write(fp, arc, compress_type=zipfile.ZIP_DEFLATED)

        replaced = False
        last_err = None
        for _ in range(6):
            try:
                os.replace(temp_output, docx_path)
                replaced = True
                break
            except Exception as _re:
                last_err = _re
                time.sleep(0.2)
        if not replaced:
            raise last_err

        for c_id, method in comment_entries:
            print(f"    ✓ 批注#{c_id} -> {method}")

        print(f"  ✓ 校对标注注入完成: 共 {len(comment_entries)} 条批注")

    except Exception as e:
        print(f"  ⚠️ 校对标注注入失败: {e}")
        import traceback; traceback.print_exc()
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if os.path.exists(temp_output):
            try: os.remove(temp_output)
            except Exception: pass


def get_docx_author(file_path: str) -> str:
    """
    不依赖 DocBuilder，直接解析 docx (zip) 中的 core.xml 获取最后修改者或创建者。
    """
    try:
        if not zipfile.is_zipfile(file_path):
            return "Unknown"
            
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 尝试读取 core-properties
            if 'docProps/core.xml' in zf.namelist():
                with zf.open('docProps/core.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    
                    # 定义命名空间 (Office Open XML core-properties)
                    namespaces = {
                        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                        'dc': 'http://purl.org/dc/elements/1.1/'
                    }
                    
                    # 优先查找最后修改者
                    last_mod = root.find('.//cp:lastModifiedBy', namespaces)
                    if last_mod is not None and last_mod.text:
                        return last_mod.text
                        
                    # 其次查找创建者
                    creator = root.find('.//dc:creator', namespaces)
                    if creator is not None and creator.text:
                        return creator.text
                        
    except Exception as e:
        # print(f"Warning: Failed to extract metadata from {os.path.basename(file_path)}: {e}")
        pass
        
    # 如果提取失败，返回 None，这就需要调用者决定用文件名还是什么
    return "Unknown"


def main():
    parser = argparse.ArgumentParser(
        description="基于 ONLYOFFICE 的多副本 Word 文档智能合并工具",
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    parser.add_argument('original', metavar='ORIGINAL_DOC', nargs='?', default='对比原稿.docx',
                        help="原稿 Word 文档路径 (.docx) [默认: BY1617126.docx]")
    
    parser.add_argument('copies_dir', metavar='COPIES_DIR', nargs='?', default='copies1',
                        help="包含所有副本的文件夹路径 [默认: . (当前目录)]")
    
    parser.add_argument('-o', '--output', required=False, default='results1',
                        help="合并后的输出文件夹路径 [默认: results]")
    
    parser.add_argument('--authors', nargs='+', default=None,
                        help="作者名字列表，顺序应与文件夹内按文件名排序后的副本一致。\n"
                             "如果某个位置填写 'None' 或 'Auto'，将自动从文档属性中提取作者名。\n"
                             "如果未提供此参数，所有副本都将自动提取作者名。\n"
                             "[默认: 自动提取]")
    parser.add_argument('--no-hierarchy-fix', action='store_true',
                        help="不修正标题层次结构")
    
    # 核心合并选项
    parser.add_argument('--diff-mode', choices=['anchor', 'direct'], default='anchor',
                        help="Diff 算法模式:\n"
                             "  anchor: 基于标题锚点分块对比（推荐，处理长文档结构更准确）\n"
                             "  direct: 直接全文线性对比（适合短文档或结构差异巨大的文档）")

    parser.add_argument('--anchor-threshold', type=float, default=0.9, metavar='0~1',
                        help="标题锁点的最低相似度阈値（归一化后）。\n"
                             "两个标题的相似度必须 ≥ 此阈値才被采用为对齐锁点。\n"
                             "标题一般改动少，高阈値能更准确区分「标题修改」和「原文不变」。\n"
                             "[0~1, 默认: 0.9]，建议范围: 0.7~1.0")

    parser.add_argument('--modify-threshold', type=float, default=0.2, metavar='0~1',
                        help="段落匹配时判定「修改」 vs 「新增」的最低相似度阈値。\n"
                             "副本段落与吟稿段落的文字相似度 ≥ 此値 → 判定为「修改」；\n"
                             "小于此値且各种兑底条件也不满足 → 判定为「新增」。\n"
                             "[0~1, 默认: 0.2]，建议范围: 0.1~0.5")

    # 标题处理选项
    parser.add_argument('--auto-fix-headings', action='store_true',
                        help="合并后自动重排标题编号（Post-processing）。\n"
                             "注意：标题识别（包含自动编号属性识别和文本正则识别）在合并过程中始终默认开启，\n"
                             "此选项仅控制是否在合稿后修正编号顺序（如将 1.1, 1.1 改为 1.1, 1.2）。")
    
    parser.add_argument('--heading-style', choices=['auto', 'numeric', 'chinese', 'mixed'],
                        default='auto',
                        help="标题编号风格：auto（自动）/numeric（1.1）/chinese（一、）/mixed（混合）")
    
    parser.add_argument('--custom-headings', nargs='*', default= None,
                        help="自定义标题锚点配置项（例如：'第一章' '1.1'）。\n"
                             "如果提供此参数，锚点匹配将不再使用全局隐式正则，而是精确按照所提供的 n 级标题层级内容及格式匹配。\n"
                             "如果留空（或不加此参数），将依然按照之前的自动标题检测方法进行框架锚定。")
    
    parser.add_argument('--max-heading-level', type=int, default=0,
                        help="限制标题识别的最大层级深度。例如设为2则仅识别L1和L2，L3及以下视为普通段落参与diff。0表示不限制。")
    
    parser.add_argument('--no-numbering-fix', action='store_true',
                        help="不修正标题编号")
    
    # 图片序号处理选项
    parser.add_argument('--auto-fix-figures', action='store_true',
                        help="合并后自动处理图片序号（识别图注、修正编号）")
    
    parser.add_argument('--figure-style', choices=['auto', 'simple', 'chapter', 'section'],
                        default='auto',
                        help="图片编号风格：auto（自动）/simple（图1）/chapter（图1.1）/section（图1-1）")
    
    # 图注语言选项已移除，默认仅支持中文（图1）
    
    # 文本校对选项
    parser.add_argument('--auto-correct-text', action='store_true',
                        default=False,
                        help="合并后自动校对文本（修正错别字、标点、语法）。默认开启。")
    
    parser.add_argument('--text-context', 
                        default='通用文档',
                        help="文本校对的上下文背景。用于提示 AI 了解文档的领域和风格，以提供更准确的建议。\n"
                             "例如：'学术论文'（正式、严谨）、'技术报告'（专业术语）、'法律文书'（规范性）、'小说'（文学性）。")
    
    # AI 辅助选项（默认不启用，提供 API 密钥后自动启用）
    parser.add_argument('--ai-api-key', 
                        help="AI API 密钥（提供后自动启用 AI 辅助）")
    
    parser.add_argument('--ai-provider', 
                        choices=['openai', 'anthropic', 'gemini', 'deepseek', 'qwen', 'local'],
                        default='deepseek',
                        help="AI 提供商（默认: deepseek）")
    
    parser.add_argument('--ai-model',
                        help="AI 模型名称（留空使用默认模型）")
    
    parser.add_argument('--ai-base-url',
                        help="自定义 AI API 地址")

    args = parser.parse_args()

    # 1. 验证原稿
    if not os.path.exists(args.original):
        print(f"错误: 原稿文件不存在: {args.original}")
        sys.exit(1)
        
    original_abs = os.path.abspath(args.original)

    # 2. 扫描副本文件夹
    if not os.path.isdir(args.copies_dir):
        print(f"错误: 副本文件夹不存在: {args.copies_dir}")
        sys.exit(1)
        
    copy_files = []
    
    # 扫描目录全量
    if not copy_files:
        copy_files = sorted([
            os.path.abspath(os.path.join(args.copies_dir, f))
            for f in os.listdir(args.copies_dir) 
            if f.lower().endswith('.docx') and not f.startswith('~$')
        ])
    
    # 过滤掉原稿文件（如果原稿也在这个目录里）
    copy_files = [f for f in copy_files if os.path.abspath(f) != original_abs]
    
    if not copy_files:
        print(f"错误: 在 {args.copies_dir} 中未找到有效的 .docx 副本文件。")
        sys.exit(1)
        
    print(f"找到 {len(copy_files)} 个副本文件:")
    
    # 3. 准备作者列表
    final_authors = []
    
    # 获取用户输入的 authors 列表 (如果存在)
    manual_authors = args.authors if args.authors else []
    
    for i, file_path in enumerate(copy_files):
        filename = os.path.basename(file_path)
        assigned_author = None
        
        # 尝试从参数获取
        if i < len(manual_authors):
            val = manual_authors[i]
            if val.lower() not in ('none', 'auto', 'null'):
                assigned_author = val
        
        # 如果没有指定，或者指定为 None -> 自动提取
        if not assigned_author:
            detected = get_docx_author(file_path)
            if detected and detected != "Unknown":
                assigned_author = detected
            else:
                # 实在提取不到，用文件名的一部分
                # 假设格式: DocName-Author.docx 或 JustName.docx
                name_part = os.path.splitext(filename)[0]
                # 简单的策略：如果文件名包含 '-' 或 '_', 取最后一部分
                if '-' in name_part:
                    assigned_author = name_part.split('-')[-1]
                elif '_' in name_part:
                    assigned_author = name_part.split('_')[-1]
                else:
                    assigned_author = f"Modifier_{i+1}"
        
        final_authors.append(assigned_author)
        print(f"  [{i+1}] {filename} -> 作者: {assigned_author}")

    # 4. 执行合并
    print("\n" + "="*50)
    print("正在初始化合并系统...")
    
    try:
        merger = MultiCopyMerger(
            original_abs,
            diff_mode=args.diff_mode,
            anchor_threshold=args.anchor_threshold,
            modify_threshold=args.modify_threshold,
            custom_headings=args.custom_headings,
            max_heading_level=args.max_heading_level,
        )
        
        for file_path, author in zip(copy_files, final_authors):
            print(f"正在添加副本: {author} ({os.path.basename(file_path)})")
            merger.add_copy(file_path, author=author)
            
        print("\n开始合并处理...")
        
        # 确定输出目录和文件名
        output_arg = os.path.abspath(args.output)
        
        if output_arg.lower().endswith('.docx'):
            # -o 指定了一个 .docx 文件路径，直接使用
            output_abs = output_arg
            output_dir = os.path.dirname(output_abs)
        else:
            # -o 指定了一个目录，自动生成文件名
            output_dir = output_arg
            original_basename = os.path.basename(original_abs)
            output_filename = os.path.splitext(original_basename)[0] + "_merged.docx"
            output_abs = os.path.join(output_dir, output_filename)
        
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"创建输出目录: {output_dir}")
        
        print(f"输出文件: {output_abs}")
        
        # 定义冲突报告路径 (改为 TXT)
        conflict_report_txt = os.path.splitext(output_abs)[0] + "_conflicts.txt"
        
        result = merger.merge(output_abs, conflict_report_path=conflict_report_txt)
        
        print("\n" + "="*50)
        print("处理完成!")
        print(f"生成文件: {result.merged_document}")
        if os.path.exists(result.conflict_report):
            print(f"冲突报告: {result.conflict_report}")
        
        # 准备全量日志收集器
        full_log_sections = []
        
        # 1. 全量变更日志
        changes_log_content = []
        changes_log_content.append(f"合并变更全量日志")
        changes_log_content.append(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        changes_log_content.append(f"="*50 + "\n")
        
        # 使用 merger 内部数据
        all_changes = []
        for author, changes in merger.changes_by_author.items():
            all_changes.extend(changes)
        
        # 按位置排序
        all_changes.sort(key=lambda c: (c.page_number, c.paragraph_index))
        
        # 准备原稿索引映射以便查找合稿前内容
        # original_structure 返回 'items'（DocItem 列表），不是 'paragraphs'
        orig_items_list = merger.original_structure.get('items', [])
        dom_to_list_idx = {p.index: i for i, p in enumerate(orig_items_list)}

        # 获取文档默认字体/字号（作为兜底）
        _doc_default_font = ''
        _doc_default_size = ''
        try:
            style_map = merger.builder._load_styles(merger.original_doc)
            # 先取默认段落样式（Normal）的字体
            default_sid = style_map.get('_default_para_style_')
            if default_sid:
                default_style = style_map.get(default_sid, {})
                _doc_default_font = default_style.get('east_asia', '')
                _doc_default_size = default_style.get('sz', '')
            # 再取 docDefaults 兜底
            defaults = style_map.get('_defaults_', {})
            if not _doc_default_font:
                _doc_default_font = defaults.get('east_asia', '')
            if not _doc_default_size:
                _doc_default_size = defaults.get('sz', '')
        except Exception:
            pass

        # 变更类型中文映射
        _TYPE_CN = {
            'insert': '插入段落',
            'delete': '删除段落',
            'modify': '修改段落',
            'table': '修改表格',
            'table_insert': '插入表格',
            'table_delete': '删除表格',
            'paragraph_insert': '插入段落',
            'paragraph_delete': '删除段落',
            'paragraph_replace': '替换段落',
            'image': '修改图片',
            'image_insert': '插入图片',
            'image_delete': '删除图片',
            'sdt_insert': '插入SDT',
            'sdt_delete': '删除SDT',
        }

        for c in all_changes:
            status = "✅ [自动合并]"
            # 检查是否在冲突中
            is_conflict = False
            for conflict in result.conflicts:
                if any(cc.change_id == c.change_id for cc in conflict.changes):
                    status = "❌ [冲突待定]"
                    is_conflict = True
                    break
            
            type_cn = _TYPE_CN.get(c.change_type, c.change_type)
            changes_log_content.append(f"{status} 第{c.page_number}页 (段落 {c.paragraph_index+1}):")
            changes_log_content.append(f"  用户: {c.author}")
            changes_log_content.append(f"  类型: {type_cn}")
            if c.section_path:
                changes_log_content.append(f"  位置: {' > '.join(c.section_path)}")
            
            # 处理显示内容
            import json
            orig_text = "(无法获取原稿内容)"
            new_text = c.content
            font_orig, size_orig = "未知", "未知"
            font_copy_log = "未知"
            size_copy_log = "未知"

            try:
                parsed = json.loads(c.content)
                if isinstance(parsed, dict):
                    if parsed.get('type') == 'paragraph' or 'old_content' in parsed:
                        # modify 类型：完整的 JSON 注解
                        orig_text = parsed.get('old_content', '(无法获取原稿内容)')
                        new_text = parsed.get('new_content', '')
                    elif parsed.get('type') == 'table':
                        # 表格 diff
                        new_text = ', '.join(
                            f"[{d.get('row','?')},{d.get('col','?')}]:{d.get('new','')}"
                            for d in parsed.get('diff_cells', [])
                        )
                elif isinstance(parsed, list):
                    full_text_parts = []
                    for item in parsed:
                        if 'text' in item:
                            full_text_parts.append(item['text'])
                        elif 'new' in item:
                            r, col = item.get('row', '?'), item.get('col', '?')
                            full_text_parts.append(f"[R{r},C{col}]: {item['new']}")
                    if full_text_parts:
                        new_text = " | ".join(full_text_parts)
            except Exception:
                pass

            # 空内容的友好显示
            if not new_text or not new_text.strip():
                if 'table' in c.change_type:
                    new_text = "(表格)"
                elif 'image' in c.change_type:
                    new_text = "(图片)"
                else:
                    new_text = "(空段落)"

            fmt_copy = c.format_info or {}
            # 优先从顶层 format_info 取字体
            font_copy_log = fmt_copy.get('font', '') or fmt_copy.get('font_name', '')
            size_copy_log = fmt_copy.get('fontSize', '') or fmt_copy.get('font_size', '')

            # insert/delete/modify 类型：字体在 runs 数组里
            if not font_copy_log and 'runs' in fmt_copy:
                for r in fmt_copy['runs']:
                    f_ = r.get('font', '')
                    s_ = r.get('fontSize', '')
                    if f_:
                        font_copy_log = f_
                    if s_:
                        size_copy_log = s_
                    if f_:
                        break

            # modify/insert 类型：从原始 XML 字符串中解析字体
            # etree.tostring 输出的命名空间属性格式是 {ns}attr，而非 w:attr
            def _xml_font_size(xml_str: str):
                """从 etree.tostring 产生的 XML 文本中提取字体和字号"""
                import re as _re
                f_, s_ = '', ''
                # 匹配 {ns}eastAsia="值" 或 w:eastAsia="值" 两种格式
                ea = _re.search(r'(?:{[^}]+}|w:)eastAsia="([^"]+)"', xml_str)
                ha = _re.search(r'(?:{[^}]+}|w:)hAnsi="([^"]+)"', xml_str)
                # <w:sz w:val="28"> 或 <{ns}sz {ns}val="28">
                sz = _re.search(r'(?:{[^}]+}|w:)sz[^>]*?(?:{[^}]+}|w:)val="([^"]+)"', xml_str)
                if ea: f_ = ea.group(1)
                elif ha: f_ = ha.group(1)
                if sz: s_ = sz.group(1)
                return f_, s_

            if not font_copy_log:
                for xml_key in ('copy_para_xml', 'copy_xml', 'copy_tbl_xml'):
                    xml_str = fmt_copy.get(xml_key, '')
                    if xml_str:
                        f_, s_ = _xml_font_size(xml_str)
                        if f_:
                            font_copy_log = f_
                            if s_: size_copy_log = s_
                            break

            if not font_copy_log:
                font_copy_log = _doc_default_font or '未知'
            if not size_copy_log:
                size_copy_log = _doc_default_size or '未知'

            # 字号转换：OpenXML 的 sz 是半磅值，转为更友好的磅值显示
            def _sz_display(sz_val):
                if not sz_val or sz_val == '未知':
                    return '未知'
                try:
                    half_pt = int(sz_val)
                    pt = half_pt / 2
                    if pt == int(pt):
                        return f"{int(pt)}pt"
                    return f"{pt}pt"
                except (ValueError, TypeError):
                    return sz_val
            
            size_copy_display = _sz_display(size_copy_log)

            if c.change_type == 'modify':
                # 尝试从原稿的 DocItem 获取原稿字体（orig_items_list 是 DocItem 列表）
                list_idx = dom_to_list_idx.get(c.paragraph_index)
                if list_idx is not None and 0 <= list_idx < len(orig_items_list):
                    orig_p = orig_items_list[list_idx]
                    # DocItem 中字体存在 runs 里
                    o_runs = getattr(orig_p, 'runs', [])
                    for _r in o_runs:
                        if _r.get('font'):
                            font_orig = _r['font']
                            size_orig = _r.get('fontSize', '') or ''
                            break
                # 兜底：文档默认字体
                if font_orig == '未知':
                    font_orig = _doc_default_font or '未知'
                if size_orig == '未知':
                    size_orig = _doc_default_size or '未知'

                size_orig_display = _sz_display(size_orig)
                changes_log_content.append(f"  [合稿前] (字体: {font_orig}, 字号: {size_orig_display}): {orig_text}")
                changes_log_content.append(f"  [合稿后] (字体: {font_copy_log}, 字号: {size_copy_display}): {new_text}")
            else:
                changes_log_content.append(f"  [合稿内容] (字体: {font_copy_log}, 字号: {size_copy_display}): {new_text}")


            changes_log_content.append("-" * 30)
            
        full_log_sections.append({
            "title": "1. 全量变更清单",
            "content": "\n".join(changes_log_content)
        })

        print(f"自动合并变更数: {result.auto_merged_count}")
        print(f"待处理冲突数:   {result.manual_review_count}")

        # ==========================================
        # 后续处理流程 (Heading -> Text -> Figure)
        # ==========================================

        # 5. 智能标题处理 (Heading)
        if args.auto_fix_headings:
            print(f"\n[5/7] 智能标题处理...")
            try:
                from .process_headings import extract_paragraphs_from_document, apply_headings_to_document, HeadingProcessor
                # 读取文档
                paragraphs = extract_paragraphs_from_document(output_abs)
                
                # 初始化处理器
                processor = HeadingProcessor(numbering_style=args.heading_style)
                
                # 执行处理
                headings, heading_report = processor.process_document(
                    paragraphs,
                    detect_headings=True,
                    infer_levels=True,
                    correct_numbering=not args.no_numbering_fix,
                    correct_hierarchy=not args.no_hierarchy_fix
                )
                
                # 获取报告文本
                temp_report = os.path.splitext(output_abs)[0] + "_heading_temp.txt"
                processor.generate_report(temp_report)
                with open(temp_report, 'r', encoding='utf-8') as f:
                    heading_report_text = f.read()
                os.remove(temp_report)
                
                full_log_sections.append({
                    "title": "2. 标题层级调整日志",
                    "content": heading_report_text
                })
                
                # 应用更改
                apply_headings_to_document(output_abs, headings, output_abs)
                
                print(f"✓ 标题处理完成 (修正: {heading_report.corrected_numbering} 处)")
                
            except Exception as e:
                print(f"⚠ 标题处理失败: {e}")
                import traceback; traceback.print_exc()
        else:
             print(f"\n[5/7] 智能标题处理 (跳过)")

        # 6. 文本校对 (Text Correction)
        if args.auto_correct_text:
            print(f"\n[6/7] 文本校对...")
            
            try:
                from ai_heading_analyzer import create_ai_corrector
                from .process_headings import extract_paragraphs_from_document
                
                # 尝试创建校对器 (如果没key会自动尝试 pycorrector)
                corrector = create_ai_corrector(
                    api_key=args.ai_api_key,
                    provider=args.ai_provider,
                    model=args.ai_model,
                    base_url=args.ai_base_url
                )
                
                if not corrector:
                    print("⚠ 错误: 未提供 --ai-api-key 且未检测到本地 pycorrector 库。")
                    print("         若要使用 AI 校对请提供 Key；若要使用离线校对请运行 `pip install pycorrector`")
                else:
                    # 获取实际使用的 Provider 名称
                    provider_name = corrector.config.provider
                    print(f"✓ 使用校对引擎: {provider_name}")
                    
                    # 重新读取文档
                    paragraphs = extract_paragraphs_from_document(output_abs)
                    
                    corrections = []
                    total_paras = len(paragraphs)
                    
                    print(f"正在分析 {total_paras} 个段落...")
                    
                    for i, para in enumerate(paragraphs):
                        if para.get('isHeading') or len(para.get('text', '').strip()) < 5:
                            continue
                        if i % 10 == 0:
                            print(f"\r  进度: {i}/{total_paras}", end="")
                        
                        result = corrector.correct_paragraph(para.get('text'), context=args.text_context)
                        
                        if result.get('has_errors'):
                            # 估算页码 (参考 MultiCopyMerger 的 30 段/页 逻辑)
                            est_page = (para.get('index', i) // 30) + 1
                            corrections.append({
                                'index': i,
                                'page': est_page,
                                'original': para.get('text'),
                                'corrected': result.get('corrected_text'),
                                'details': result.get('corrections')
                            })
                    
                    print(f"\r  进度: {total_paras}/{total_paras} (完成)")
                    
                    # 生成报告文本
                    correction_report_text = f"文本校对报告\n====================\n发现错误段落数: {len(corrections)}\n\n"
                    for c in corrections:
                        correction_report_text += f"第 {c['page']} 页 (段落 #{c['index']}):\n"
                        correction_report_text += f"❌ 原文: {c['original']}\n"
                        correction_report_text += f"✅ 修正: {c['corrected']}\n"
                        for d in c['details']:
                            correction_report_text += f"   - [{d.get('type')}] {d.get('original')} -> {d.get('corrected')} ({d.get('explanation')})\n"
                        correction_report_text += "\n--------------------\n\n"
                            
                    full_log_sections.append({
                        "title": "3. 文本校对日志",
                        "type": "text_correction",
                        "content": correction_report_text,
                        "data": corrections # 存储结构化数据供后续生成红字
                    })
                            
                    # 生成 TXT 格式的校对报告
                    correction_report_txt = os.path.splitext(output_abs)[0] + "_text_correction.txt"
                    with open(correction_report_txt, 'w', encoding='utf-8') as f:
                        f.write(correction_report_text)
                    print(f"✓ 校对报告已生成: {correction_report_txt}")

                    # 重新合成带标注的文档内容（纯 Python XML 方式）
                    if corrections:
                        print(f"正在向文档应用 {len(corrections)} 处标注...")
                        try:
                            # 直接通过 XML 操作注入校对批注，不依赖 docbuilder
                            inject_corrections_to_docx(output_abs, corrections)
                            print(f"✓ 已在文档中成功标注 {len(corrections)} 处修正")
                        except Exception as e:
                            print(f"⚠️ 应用标注失败: {e}")
                            import traceback; traceback.print_exc()

                    print(f"\n✓ 文本校对完成 (发现问题: {len(corrections)} 处)")

            except Exception as e:
                print(f"⚠ 文本校对失败: {e}")
                import traceback; traceback.print_exc()
        else:
            print(f"\n[6/7] 文本校对 (跳过)")

        # 7. 图片序号处理 (Figure)
        if args.auto_fix_figures:
            print(f"\n[7/7] 图片序号处理...")
            try:
                from .figure_numbering import process_document_figures
                from .process_headings import extract_paragraphs_from_document
                from ai_heading_analyzer import create_ai_analyzer
                
                paragraphs = extract_paragraphs_from_document(output_abs)
                document_structure = {'paragraphs': paragraphs}
                
                # 初始化 AI
                ai_analyzer = None
                if args.ai_api_key:
                    ai_analyzer = create_ai_analyzer(
                        api_key=args.ai_api_key,
                        provider=args.ai_provider,
                        model=args.ai_model,
                        base_url=args.ai_base_url
                    )
                
                figures, figure_report = process_document_figures(
                    document_structure,
                    numbering_style=args.figure_style,
                    language='chinese', # 默认中文
                    ai_analyzer=ai_analyzer,
                    use_ai=bool(ai_analyzer)
                )
                
                full_log_sections.append({
                    "title": "4. 图片编号变更日志",
                    "content": figure_report
                })
                
                print(f"✓ 图片处理完成 (检测: {len(figures)} 个)")
                
            except Exception as e:
                print(f"⚠ 图片处理失败: {e}")
                import traceback; traceback.print_exc()
        else:
            print(f"\n[7/7] 图片序号处理 (跳过)")

        # === 最终生成合并日志 TXT ===
        if full_log_sections:
            merged_log_txt = os.path.splitext(output_abs)[0] + "_merged_changes_log.txt"
            print(f"\n正在生成全量日志文档: {merged_log_txt} ...")
            try:
                with open(merged_log_txt, 'w', encoding='utf-8') as f:
                    # 总标题
                    f.write("=" * 70 + "\n")
                    f.write("          文档合并全量综合日志\n")
                    f.write("=" * 70 + "\n")
                    f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"原稿文件: {os.path.basename(original_abs)}\n")
                    f.write(f"输出文件: {os.path.basename(output_abs)}\n")
                    f.write(f"副本数量: {len(copy_files)}\n")
                    f.write(f"自动合并变更数: {result.auto_merged_count}\n")
                    f.write(f"待处理冲突数:   {result.manual_review_count}\n")
                    f.write("\n")

                    for section in full_log_sections:
                        # 章节标题
                        f.write("\n" + "=" * 70 + "\n")
                        f.write(section['title'] + "\n")
                        f.write("=" * 70 + "\n\n")

                        if section.get('type') == 'text_correction' and section.get('data'):
                            # 专门处理文本校对日志，用文字方式呈现修正差异
                            import difflib
                            for c in section['data']:
                                f.write(f"第 {c['page']} 页 (段落 #{c['index']}):\n")
                                f.write(f"  ❌ 原文: {c['original']}\n")
                                f.write(f"  ✅ 修正: ")

                                # 用文字方式标记差异部分
                                orig_text = c['original']
                                corr_text = c['corrected']
                                s = difflib.SequenceMatcher(None, orig_text, corr_text)
                                diff_parts = []
                                for tag, i1, i2, j1, j2 in s.get_opcodes():
                                    part_text = corr_text[j1:j2]
                                    if not part_text:
                                        continue
                                    if tag == 'equal':
                                        diff_parts.append(part_text)
                                    else:
                                        # 差异部分用方括号和下划线标记
                                        diff_parts.append(f"[_{part_text}_]")
                                f.write("".join(diff_parts) + "\n")

                                # 详情
                                for d in c['details']:
                                    f.write(f"     - [{d.get('type', '未知')}] "
                                            f"「{d.get('original', '')}」→「{d.get('corrected', '')}」"
                                            f" ({d.get('explanation', '')})\n")
                                f.write("\n" + "-" * 30 + "\n\n")
                        else:
                            # 内容 (直接写入)
                            f.write(section['content'] + "\n")

                        # 分隔符
                        f.write("\n")

                    # 附录：副本作者信息
                    f.write("\n" + "=" * 70 + "\n")
                    f.write("附录：副本作者信息\n")
                    f.write("=" * 70 + "\n")
                    for i, (fp, author) in enumerate(zip(copy_files, final_authors)):
                        f.write(f"  [{i+1}] {os.path.basename(fp)} -> 作者: {author}\n")
                    f.write("\n")

                print(f"✓ 全量日志文档生成完毕: {merged_log_txt}")

            except Exception as e:
                 print(f"⚠️ 生成日志文档失败: {e}")
                 import traceback; traceback.print_exc()

        # === 最终步骤: 生成日志文档 ===
        # if os.path.exists(output_abs) and merger and hasattr(merger, 'changes_by_author'):
        #     try:
        #         all_changes = []
        #         for changes in merger.changes_by_author.values():
        #             all_changes.extend(changes)
        #         if all_changes:
        #             inject_comments_to_docx(output_abs, all_changes)
        #     except Exception as e:
        #         print(f"⚠️ 批注注入失败: {e}")

    except Exception as e:
        print(f"\n错误: 合并过程中发生异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
