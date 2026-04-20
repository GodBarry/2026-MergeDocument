#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
混合锚点定位引擎
Hybrid Anchor Locator Engine

实现三层定位策略：
1. 书签精确匹配（Bookmark）
2. 标题锚点定位（Heading Anchor）
3. 内容指纹匹配（Content Fingerprint）

优先级：书签 > 标题 > 指纹 > 模糊匹配
"""

import re
import hashlib
from typing import Optional, Tuple, List, Dict
from dataclasses import dataclass


@dataclass
class LocateResult:
    """定位结果"""
    index: int                    # 原稿段落索引
    confidence: float             # 置信度 [0.0, 1.0]
    method: str                   # 定位方法
    details: str                  # 详细信息


class ContentFingerprint:
    """内容指纹生成器（SimHash实现）"""
    
    @staticmethod
    def normalize_text(text: str) -> str:
        """文本归一化"""
        if not text:
            return ""
        
        # 1. 全角转半角
        fullwidth_map = str.maketrans({
            '，': ',', '。': '.', '；': ';', '：': ':', '！': '!', '？': '?',
            '（': '(', '）': ')', '【': '[', '】': ']', '《': '<', '》': '>',
            '、': ',', '．': '.', '\u201c': '"', '\u201d': '"',
            '\u3000': ' ', '—': '-', '～': '~',
        })
        text = text.translate(fullwidth_map)
        
        # 2. 移除标点和空白
        text = re.sub(r'[^\w\s]', '', text)
        text = re.sub(r'\s+', ' ', text)
        
        # 3. 转小写
        return text.strip().lower()
    
    @staticmethod
    def extract_features(text: str, n: int = 3) -> List[str]:
        """提取N-gram特征"""
        normalized = ContentFingerprint.normalize_text(text)
        if len(normalized) < n:
            return [normalized] if normalized else []
        
        # 生成字符级N-gram
        features = []
        for i in range(len(normalized) - n + 1):
            features.append(normalized[i:i+n])
        
        return features
    
    @staticmethod
    def hash_feature(feature: str) -> int:
        """特征哈希（64位）"""
        h = hashlib.md5(feature.encode('utf-8')).hexdigest()
        return int(h[:16], 16)  # 取前64位
    
    @staticmethod
    def simhash(text: str, bits: int = 64) -> int:
        """生成SimHash指纹"""
        if not text or len(text.strip()) < 3:
            return 0
        
        features = ContentFingerprint.extract_features(text)
        if not features:
            return 0
        
        # 初始化向量
        v = [0] * bits
        
        # 累加特征哈希
        for feature in features:
            h = ContentFingerprint.hash_feature(feature)
            for i in range(bits):
                if h & (1 << i):
                    v[i] += 1
                else:
                    v[i] -= 1
        
        # 生成最终指纹
        fingerprint = 0
        for i in range(bits):
            if v[i] > 0:
                fingerprint |= (1 << i)
        
        return fingerprint
    
    @staticmethod
    def hamming_distance(fp1: int, fp2: int) -> int:
        """计算汉明距离"""
        xor = fp1 ^ fp2
        distance = 0
        while xor:
            distance += 1
            xor &= xor - 1  # 清除最低位的1
        return distance
    
    @staticmethod
    def similarity(fp1: int, fp2: int, bits: int = 64) -> float:
        """计算相似度 [0.0, 1.0]"""
        distance = ContentFingerprint.hamming_distance(fp1, fp2)
        return 1.0 - (distance / bits)


class BookmarkManager:
    """书签管理器"""
    
    # 书签前缀
    BOOKMARK_PREFIX = "_kiro_anchor_"
    
    @staticmethod
    def generate_bookmark_id(index: int, content_hash: str = "") -> str:
        """生成书签ID"""
        if content_hash:
            return f"{BookmarkManager.BOOKMARK_PREFIX}{index}_{content_hash[:8]}"
        return f"{BookmarkManager.BOOKMARK_PREFIX}{index}"
    
    @staticmethod
    def extract_bookmark(para: Dict) -> Optional[str]:
        """从段落中提取书签ID"""
        # 方法1: 从format_info中提取
        fmt = para.get('format_info', {}) or {}
        bookmark = fmt.get('bookmark_id')
        if bookmark and bookmark.startswith(BookmarkManager.BOOKMARK_PREFIX):
            return bookmark
        
        # 方法2: 从文本中提取隐藏书签标记
        text = para.get('text', '')
        match = re.search(rf'{BookmarkManager.BOOKMARK_PREFIX}(\d+)(?:_([a-f0-9]{{8}}))?', text)
        if match:
            return match.group(0)
        
        return None
    
    @staticmethod
    def parse_bookmark_index(bookmark_id: str) -> Optional[int]:
        """从书签ID中解析索引"""
        if not bookmark_id or not bookmark_id.startswith(BookmarkManager.BOOKMARK_PREFIX):
            return None
        
        match = re.match(rf'{BookmarkManager.BOOKMARK_PREFIX}(\d+)', bookmark_id)
        if match:
            return int(match.group(1))
        return None


class HybridAnchorLocator:
    """混合锚点定位器"""
    
    def __init__(self, original_paras: List[Dict], enable_bookmark: bool = True):
        """
        初始化定位器
        
        Args:
            original_paras: 原稿段落列表
            enable_bookmark: 是否启用书签模式
        """
        self.original_paras = original_paras
        self.enable_bookmark = enable_bookmark
        
        # 预计算指纹索引
        self._build_fingerprint_index()
        
        # 构建书签索引
        if enable_bookmark:
            self._build_bookmark_index()
    
    def _build_fingerprint_index(self):
        """构建内容指纹索引"""
        self.fingerprints = {}
        for i, para in enumerate(self.original_paras):
            text = para.get('text', '').strip()
            if text and len(text) >= 10:  # 只为足够长的文本生成指纹
                fp = ContentFingerprint.simhash(text)
                self.fingerprints[i] = fp
    
    def _build_bookmark_index(self):
        """构建书签索引"""
        self.bookmark_index = {}
        for i, para in enumerate(self.original_paras):
            bookmark = BookmarkManager.extract_bookmark(para)
            if bookmark:
                self.bookmark_index[bookmark] = i
    
    def locate(self, copy_para: Dict, section_range: Tuple[int, int] = None) -> Optional[LocateResult]:
        """
        定位段落
        
        Args:
            copy_para: 副本段落
            section_range: 章节范围 (start_idx, end_idx)，用于限制搜索范围
        
        Returns:
            LocateResult 或 None
        """
        # 确定搜索范围
        if section_range:
            start_idx, end_idx = section_range
        else:
            start_idx, end_idx = 0, len(self.original_paras)
        
        # 第1层：书签精确匹配
        if self.enable_bookmark:
            result = self._locate_by_bookmark(copy_para, start_idx, end_idx)
            if result:
                return result
        
        # 第2层：内容指纹匹配
        result = self._locate_by_fingerprint(copy_para, start_idx, end_idx)
        if result:
            return result
        
        # 第3层：模糊文本匹配（兜底）
        result = self._locate_by_fuzzy_text(copy_para, start_idx, end_idx)
        if result:
            return result
        
        return None
    
    def _locate_by_bookmark(self, copy_para: Dict, start_idx: int, end_idx: int) -> Optional[LocateResult]:
        """通过书签定位"""
        bookmark = BookmarkManager.extract_bookmark(copy_para)
        if not bookmark:
            return None
        
        # 查找书签
        orig_idx = self.bookmark_index.get(bookmark)
        if orig_idx is not None and start_idx <= orig_idx < end_idx:
            return LocateResult(
                index=orig_idx,
                confidence=1.0,
                method="bookmark",
                details=f"精确书签匹配: {bookmark}"
            )
        
        return None
    
    def _locate_by_fingerprint(self, copy_para: Dict, start_idx: int, end_idx: int) -> Optional[LocateResult]:
        """通过内容指纹定位"""
        copy_text = copy_para.get('text', '').strip()
        if not copy_text or len(copy_text) < 10:
            return None
        
        copy_fp = ContentFingerprint.simhash(copy_text)
        if copy_fp == 0:
            return None
        
        # 在范围内查找最相似的段落
        best_match = None
        best_similarity = 0.0
        
        for i in range(start_idx, end_idx):
            if i not in self.fingerprints:
                continue
            
            orig_fp = self.fingerprints[i]
            similarity = ContentFingerprint.similarity(copy_fp, orig_fp)
            
            if similarity > best_similarity:
                best_similarity = similarity
                best_match = i
        
        # 相似度阈值：0.75（允许25%差异，更宽松）
        if best_match is not None and best_similarity >= 0.75:
            distance = ContentFingerprint.hamming_distance(copy_fp, self.fingerprints[best_match])
            return LocateResult(
                index=best_match,
                confidence=0.90 * best_similarity,  # 基础置信度0.90
                method="fingerprint",
                details=f"指纹匹配: 相似度={best_similarity:.3f}, 汉明距离={distance}"
            )
        
        return None
    
    def _locate_by_fuzzy_text(self, copy_para: Dict, start_idx: int, end_idx: int) -> Optional[LocateResult]:
        """通过模糊文本匹配定位（兜底策略）"""
        copy_text = copy_para.get('text', '').strip()
        if not copy_text:
            return None
        
        copy_normalized = ContentFingerprint.normalize_text(copy_text)
        if not copy_normalized:
            return None
        
        # 使用SequenceMatcher进行模糊匹配
        import difflib
        best_match = None
        best_ratio = 0.0
        
        for i in range(start_idx, end_idx):
            orig_text = self.original_paras[i].get('text', '').strip()
            if not orig_text:
                continue
            
            orig_normalized = ContentFingerprint.normalize_text(orig_text)
            if not orig_normalized:
                continue
            
            ratio = difflib.SequenceMatcher(None, copy_normalized, orig_normalized).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = i
        
        # 相似度阈值：0.70
        if best_match is not None and best_ratio >= 0.70:
            return LocateResult(
                index=best_match,
                confidence=0.70 * best_ratio,
                method="fuzzy_text",
                details=f"模糊匹配: 相似度={best_ratio:.3f}"
            )
        
        return None


class TableFingerprint:
    """表格指纹生成器"""
    
    @staticmethod
    def generate(table_data: Dict) -> str:
        """生成表格指纹（结构+内容）"""
        rows = table_data.get('rows', [])
        if not rows:
            return ""
        
        # 1. 结构指纹：行数×列数
        n_rows = len(rows)
        n_cols = len(rows[0]) if rows else 0
        structure = f"{n_rows}x{n_cols}"
        
        # 2. 内容指纹：首行+首列
        first_row = []
        first_col = []
        
        if rows:
            # 首行
            for cell in rows[0]:
                text = cell.get('text', '') if isinstance(cell, dict) else str(cell)
                first_row.append(ContentFingerprint.normalize_text(text))
            
            # 首列
            for row in rows:
                if row:
                    cell = row[0]
                    text = cell.get('text', '') if isinstance(cell, dict) else str(cell)
                    first_col.append(ContentFingerprint.normalize_text(text))
        
        # 3. 组合指纹
        content = f"{structure}|{'|'.join(first_row[:3])}|{'|'.join(first_col[:3])}"
        
        # 4. 生成哈希
        return hashlib.md5(content.encode('utf-8')).hexdigest()[:16]
    
    @staticmethod
    def similarity(fp1: str, fp2: str) -> float:
        """计算表格指纹相似度"""
        if not fp1 or not fp2:
            return 0.0
        if fp1 == fp2:
            return 1.0
        
        # 简单的字符串相似度
        import difflib
        return difflib.SequenceMatcher(None, fp1, fp2).ratio()


def inject_bookmarks_to_document(doc_path: str, output_path: str = None) -> str:
    """
    为文档注入书签（预处理步骤）
    
    Args:
        doc_path: 输入文档路径
        output_path: 输出文档路径（None则覆盖原文件）
    
    Returns:
        输出文档路径
    """
    # TODO: 实现书签注入逻辑
    # 这需要使用python-docx或ONLYOFFICE API
    pass


if __name__ == "__main__":
    # 测试代码
    print("混合锚点定位引擎 - 测试")
    
    # 测试内容指纹
    text1 = "第一章 项目概述\n本项目旨在开发一个智能文档合并系统。"
    text2 = "第一章 项目概述\n本项目旨在开发一个智能文档合并系统，支持多人协作。"
    
    fp1 = ContentFingerprint.simhash(text1)
    fp2 = ContentFingerprint.simhash(text2)
    
    print(f"文本1指纹: {fp1:016x}")
    print(f"文本2指纹: {fp2:016x}")
    print(f"汉明距离: {ContentFingerprint.hamming_distance(fp1, fp2)}")
    print(f"相似度: {ContentFingerprint.similarity(fp1, fp2):.3f}")
