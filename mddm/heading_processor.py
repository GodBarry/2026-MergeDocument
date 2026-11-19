"""
智能标题处理模块
功能：
1. 智能识别标题（不依赖样式）
2. 推断真实标题级别
3. 自动修正序号错误
4. 补全缺失的层次结构
5. 生成详细的处理报告

适用场景：
- 单个文档的标题修正
- 合并过程中的标题处理
"""

import re
import os
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class HeadingInfo:
    """标题信息"""
    index: int                          # 在文档中的索引
    text: str                           # 标题文本
    style_level: Optional[int] = None   # 样式声明的级别
    true_level: Optional[int] = None    # 推断的真实级别
    confidence: float = 0.0             # 置信度
    numbering: Optional[str] = None     # 编号部分
    content: Optional[str] = None       # 内容部分
    numbering_depth: Optional[int] = None  # 编号深度
    font_size: int = 12                 # 字体大小
    is_bold: bool = False               # 是否加粗
    is_centered: bool = False           # 是否居中
    spacing_before: int = 0             # 段前间距
    spacing_after: int = 0              # 段后间距
    ilvl: Optional[int] = None          # 自动编号级别 (0-8)
    num_id: Optional[int] = None        # 自动编号ID
    style_name: Optional[str] = None    # 样式名称
    is_placeholder: bool = False        # 是否为占位标题
    evidence: Dict = field(default_factory=dict)  # 证据链
    original_text: Optional[str] = None  # 修正前的原始文本


@dataclass
class ProcessingReport:
    """处理报告"""
    total_headings: int = 0
    detected_headings: int = 0
    corrected_numbering: int = 0
    inserted_placeholders: int = 0
    style_mismatches: int = 0
    errors_fixed: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    details: List[str] = field(default_factory=list)
    timestamp: datetime = field(default_factory=datetime.now)


class HeadingNumberingPatterns:
    """
    标题编号模式识别器
    支持中文、英文、数字、罗马数字等各种编号风格
    """
    
    # 一级标题模式
    LEVEL_1_PATTERNS = [
        (r'^第[一二三四五六七八九十百千万壹贰叁肆伍陆柒捌玖拾佰仟]+章', '中文章'),
        (r'^第[一二三四五六七八九十百千万壹贰叁肆伍陆柒捌玖拾佰仟]+篇', '中文篇'),
        (r'^第[一二三四五六七八九十百千万壹贰叁肆伍陆柒捌玖拾佰仟]+部分', '中文部分'),
        (r'^第[一二三四五六七八九十百千万壹贰叁肆伍陆柒捌玖拾佰仟]+编', '中文编'),
        (r'^[一二三四五六七八九十百]+[、\s\.]', '中文数字符号'),
        (r'^[壹贰叁肆伍陆柒捌玖拾佰]+[、\s\.]', '大写中文数字符号'),
        (r'^\d{1,3}[、\s\.]', '数字符号'),
        (r'^[IVXLCDM]+[\.、\s]', '大写罗马数字'),
        (r'^[Ⅰ Ⅱ Ⅲ Ⅳ Ⅴ Ⅵ Ⅶ Ⅷ Ⅸ Ⅹ]+', '罗马数字符号'),
        (r'^Chapter\s+\d+', '英文Chapter'),
        (r'^CHAPTER\s+\d+', '英文CHAPTER'),
        (r'^Part\s+\d+', '英文Part'),
        (r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', '带圈数字'),
        (r'^[A-Z]\s*[\.、]', '大写字母'),
    ]
    
    # 二级标题模式
    LEVEL_2_PATTERNS = [
        (r'^\d{1,3}\.\d{1,3}[、\s\.]', '两级数字'),
        (r'^\d{1,3}\.\d{1,3}\s*$', '两级数字纯'),
        (r'^第[一二三四五六七八九十百千]+节', '中文节'),
        (r'^第[一二三四五六七八九十百千]+条', '中文条'),
        (r'^\([一二三四五六七八九十]+\)', '括号中文'),
        (r'^\（[一二三四五六七八九十]+\）', '全角括号中文'),
        (r'^\(\d{1,3}\)', '括号数字'),
        (r'^\（\d{1,3}\）', '全角括号数字'),
        (r'^[A-Z]\.\d{1,3}', '字母数字组合'),
        (r'^\([A-Z]\)', '括号大写字母'),
        (r'^[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽]', '带括号数字符号'),
    ]
    
    # 三级标题模式
    LEVEL_3_PATTERNS = [
        (r'^\d{1,3}\.\d{1,3}\.\d{1,3}[、\s\.]', '三级数字'),
        (r'^\d{1,3}\.\d{1,3}\.\d{1,3}\s*$', '三级数字纯'),
        (r'^[a-z]\.[、\s]', '小写字母点'),
        (r'^\([a-z]\)', '括号小写字母'),
        (r'^\d+\)', '数字右括号'),
        (r'^[①②③④⑤⑥⑦⑧⑨⑩]+', '带圈数字小'),
        (r'^[●○◆◇■□▪▫]', '几何符号'),
    ]
    
    # 四级标题模式
    LEVEL_4_PATTERNS = [
        (r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', '四级数字'),
        (r'^[a-z]\.\d+', '字母数字'),
        (r'^[ivxlcdm]+\)', '小写罗马右括号'),
    ]
    
    # 五级标题模式
    LEVEL_5_PATTERNS = [
        (r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', '五级数字'),
        (r'^[-–—]', '短横线'),
        (r'^[•·]', '圆点'),
        (r'^[*＊]', '星号'),
    ]
    
    # 六级标题模式
    LEVEL_6_PATTERNS = [
        (r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', '六级数字'),
    ]
    
    @classmethod
    def detect_level(cls, text: str) -> Tuple[Optional[int], Optional[str]]:
        """
        检测文本的标题级别
        
        返回: (level, pattern_name) 或 (None, None)
        """
        text = text.strip()
        
        for level, patterns in [
            (1, cls.LEVEL_1_PATTERNS),
            (2, cls.LEVEL_2_PATTERNS),
            (3, cls.LEVEL_3_PATTERNS),
            (4, cls.LEVEL_4_PATTERNS),
            (5, cls.LEVEL_5_PATTERNS),
            (6, cls.LEVEL_6_PATTERNS),
        ]:
            for pattern, name in patterns:
                if re.match(pattern, text, re.IGNORECASE):
                    return (level, name)
        
        return (None, None)
    
    @classmethod
    def extract_numbering(cls, text: str) -> Tuple[Optional[str], str]:
        """
        提取编号部分
        
        返回: (numbering, content)
        """
        text = text.strip()
        
        all_patterns = (
            cls.LEVEL_1_PATTERNS + 
            cls.LEVEL_2_PATTERNS + 
            cls.LEVEL_3_PATTERNS + 
            cls.LEVEL_4_PATTERNS +
            cls.LEVEL_5_PATTERNS +
            cls.LEVEL_6_PATTERNS
        )
        
        for pattern, _ in all_patterns:
            match = re.match(pattern, text)
            if match:
                numbering = match.group(0)
                content = text[len(numbering):].strip()
                return (numbering, content)
        
        return (None, text)
    
    @classmethod
    def get_numbering_depth(cls, text: str) -> Optional[int]:
        """
        获取编号的层级深度
        
        例如: "1" -> 1, "1.1" -> 2, "1.1.1" -> 3
        """
        match = re.match(r'^(\d+(?:\.\d+)*)', text)
        if match:
            numbering = match.group(1)
            return numbering.count('.') + 1
        
        if re.match(r'^第.+章', text):
            return 1
        if re.match(r'^第.+节', text):
            return 2
        
        return None


class HeadingDetector:
    """
    智能标题检测器
    使用多维度综合评分识别标题
    """
    
    def __init__(self):
        self.pattern_detector = HeadingNumberingPatterns()
        self.document_base_font_size = 12  # 默认正文字体大小
    
    def detect_headings(self, paragraphs: List[Dict]) -> List[HeadingInfo]:
        """
        检测文档中的所有标题
        
        参数:
            paragraphs: 段落列表
            
        返回:
            标题信息列表
        """
        # 第一遍：提取特征
        para_features = []
        for i, para in enumerate(paragraphs):
            features = self._extract_features(para, i)
            para_features.append(features)
        
        # 估算文档基准字体大小
        self.document_base_font_size = self._estimate_base_font_size(para_features)
        
        # 第二遍：识别标题
        headings = []
        for features in para_features:
            if self._is_heading(features):
                heading = self._create_heading_info(features)
                headings.append(heading)
        
        return headings
    
    def _extract_features(self, para: Dict, index: int) -> Dict:
        """提取段落特征"""
        text = para.get('text', '').strip()
        
        # 提取编号
        numbering, content = self.pattern_detector.extract_numbering(text)
        numbering_depth = self.pattern_detector.get_numbering_depth(text)
        
        # 检测编号级别
        pattern_level, pattern_name = self.pattern_detector.detect_level(text)
        
        return {
            'index': index,
            'text': text,
            'numbering': numbering,
            'content': content,
            'numbering_depth': numbering_depth,
            'pattern_level': pattern_level,
            'pattern_name': pattern_name,
            'font_size': para.get('font_size', 12),
            'is_bold': para.get('is_bold', False),
            'is_centered': para.get('is_centered', False),
            'spacing_before': para.get('spacing_before', 0),
            'spacing_after': para.get('spacing_after', 0),
            'text_length': len(text),
            'ilvl': para.get('ilvl'),
            'num_id': para.get('num_id'),
            'style_name': para.get('style_name')
        }
    
    def _estimate_base_font_size(self, features_list: List[Dict]) -> int:
        """估算文档的基准字体大小（正文字体）"""
        # 排除可能是标题的（太大的字体）
        # 简单 heuristic: 排除 style 含有 "Heading" / "标题" 的
        valid_sizes = []
        for f in features_list:
             if f['font_size'] <= 0: continue
             name = (f.get('style_name') or "").lower()
             if "heading" in name or "标题" in name:
                 continue
             valid_sizes.append(f['font_size'])
             
        if not valid_sizes:
            # 如果全被排除，说明都是正常或者是空文档，这就fallback
            valid_sizes = [f['font_size'] for f in features_list if f['font_size'] > 0]

        if not valid_sizes:
            return 12
        
        # 使用众数作为基准字体大小
        from collections import Counter
        counter = Counter(valid_sizes)
        base_size = counter.most_common(1)[0][0]
        return base_size
    
    def _is_heading(self, features: Dict) -> bool:
        """
        判断是否为标题
        使用综合评分系统
        """
        score = 0
        
        # 样式名称 (权重 100) - 绝对信号
        style_name = (features.get('style_name') or "").lower()
        if "heading" in style_name or "标题" in style_name:
             # 确保不是“标题 9”这种太深的
             # 通常 Heading 1-9
             import re
             if re.search(r'(heading|标题)\s*[1-6]', style_name):
                 return True # 直接判定为真

        # 自动编号 (权重 30) - 辅助信号 
        # (加强：如果是有层级的编号，且文本短，大概率是标题)
        if features.get('num_id') is not None and features.get('ilvl') is not None:
             if features['text_length'] < 100:
                score += 30

        # 编号模式（权重 40）
        if features['pattern_level']:
            score += 40
        
        # 字体大小（权重 30）
        font_diff = features['font_size'] - self.document_base_font_size
        if font_diff >= 10:
            score += 30
        elif font_diff >= 6:
            score += 25
        elif font_diff >= 4:
            score += 20
        elif font_diff >= 2:
            score += 15
        
        # 加粗（权重 15）
        if features['is_bold']:
            score += 15
        
        # 文本长度（权重 10）
        if features['text_length'] <= 50:
            score += 10
        elif features['text_length'] <= 100:
            score += 5
        
        # 段落间距（权重 5）
        if features['spacing_before'] > 200 or features['spacing_after'] > 200:
            score += 5
        
        # 阈值判断
        HEADING_THRESHOLD = 40
        return score >= HEADING_THRESHOLD
    
    def _create_heading_info(self, features: Dict) -> HeadingInfo:
        """创建标题信息对象"""
        return HeadingInfo(
            index=features['index'],
            text=features['text'],
            numbering=features['numbering'],
            content=features['content'],
            numbering_depth=features['numbering_depth'],
            font_size=features['font_size'],
            is_bold=features['is_bold'],
            is_centered=features['is_centered'],
            spacing_before=features['spacing_before'],
            spacing_after=features['spacing_after'],
            ilvl=features.get('ilvl'),
            num_id=features.get('num_id'),
            style_name=features.get('style_name')
        )


class RelativeFormatLevelInferenceEngine:
    """
    基于格式相对关系的级别推断引擎
    
    核心原则：格式出现顺序 = 级别
    - 第一个遇到的格式 → 1级
    - 第二个新格式 → 2级
    - 第n个新格式 → n级
    - 已知格式 → 使用该格式首次出现时确定的级别
    """
    
    def __init__(self):
        self.pattern_detector = HeadingNumberingPatterns()
        self.format_stack: List[str] = []  # 已知格式栈，按首次出现顺序
    
    def infer_levels(self, headings: List[HeadingInfo]) -> List[HeadingInfo]:
        """
        推断所有标题的真实级别
        
        Args:
            headings: 标题列表（按文档顺序排列）
            
        Returns:
            更新后的标题列表
        """
        self.format_stack = []
        
        for heading in headings:
            self._infer_single_level(heading)
        
        return headings
    
    def _get_format_key(self, heading: HeadingInfo) -> str:
        """
        提取标题的格式特征键
        
        格式键由多个特征组合：
        - 编号深度（1.1 → 深度2，1.1.1 → 深度3）
        - 编号前缀（第、1.、1) 等）
        - 编号符号（.、、- 等）
        """
        # 1. 获取编号深度
        depth = heading.numbering_depth or 1
        
        # 2. 提取编号前缀（用于区分"1."和"第一章"等）
        numbering = heading.numbering or ""
        prefix = ""
        if numbering:
            # 提取前缀模式
            if re.match(r'^第.+章', numbering):
                prefix = "第X章"
            elif re.match(r'^第.+节', numbering):
                prefix = "第X节"
            elif re.match(r'^\d+\.', numbering):
                prefix = "N."
            elif re.match(r'^\d+\)', numbering):
                prefix = "N)"
            elif re.match(r'^\([一二三四五六七八九十\d]+\)', numbering):
                prefix = "(N)"
            elif re.match(r'^\([A-Z]\)', numbering):
                prefix = "(X)"
            else:
                prefix = numbering[:3]  # 取前3个字符作为前缀
        
        # 3. 组合格式键：深度_前缀
        format_key = f"{depth}_{prefix}"
        return format_key
    
    def _infer_single_level(self, heading: HeadingInfo):
        """推断单个标题的级别"""
        format_key = self._get_format_key(heading)
        
        if format_key in self.format_stack:
            # 已知格式：使用该格式首次出现时的级别
            level = self.format_stack.index(format_key) + 1
            heading.true_level = level
            heading.confidence = 1.0
            heading.evidence = {'method': 'relative_format', 'format_key': format_key, 'level_from_stack': True}
        else:
            # 新格式：分配新级别
            level = len(self.format_stack) + 1
            heading.true_level = level
            heading.confidence = 0.8  # 新格式置信度稍低
            heading.evidence = {'method': 'relative_format', 'format_key': format_key, 'new_format': True}
            # 入栈
            self.format_stack.append(format_key)


class HeadingNumberingCorrector:
    """
    标题序号自动修正器
    """
    
    def __init__(self, numbering_style='auto'):
        self.style = numbering_style
        self.corrections = []
    
    def correct_numbering(self, headings: List[HeadingInfo]) -> List[HeadingInfo]:
        """
        修正所有标题的编号
        """
        # 检测主要编号风格和具体格式
        if self.style == 'auto':
            self.style, self.level_formats = self._detect_detailed_styles(headings)
        else:
            # 如果指定了风格，则使用默认的层级格式
            self.level_formats = {i: '.' for i in range(1, 7)} 
        
        # 重新编号
        level_counters = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
        
        for heading in headings:
            level = heading.true_level
            
            # 更新计数器
            level_counters[level] += 1
            for deeper in range(level + 1, 7):
                level_counters[deeper] = 0
            
            # 生成正确编号
            correct_numbering = self._generate_numbering(level_counters, level)
            
            # 替换编号 (传入当前 level 以应用对应格式)
            original_text = heading.text
            corrected_text = self._replace_numbering(original_text, correct_numbering, heading.content or "", level)
            
            if corrected_text != original_text:
                heading.original_text = original_text
                heading.text = corrected_text
                heading.numbering = correct_numbering
                self.corrections.append(heading)
        
        return headings
    
    def _detect_detailed_styles(self, headings: List[HeadingInfo]) -> Tuple[str, Dict[int, str]]:
        """检测主要编号风格以及每一级的多数派分隔符"""
        style_counts = {'numeric': 0, 'chinese': 0, 'chapter': 0}
        level_sep_counts = {i: {} for i in range(1, 7)}
        
        for heading in headings:
            num = heading.numbering
            if not num: continue
            
            # 1. 统计大类风格
            if re.match(r'^\d+', num):
                style_counts['numeric'] += 1
            elif re.match(r'^[一二三四五六七八九十]+', num):
                style_counts['chinese'] += 1
            elif re.match(r'^第.+[章节]', num):
                style_counts['chapter'] += 1
            
            # 2. 统计具体分隔符 (例如 "1." vs "1、" vs "1 ")
            # 匹配数字或中文数字后的最后一个非文字/非数字字符
            match = re.search(r'([\.、\s]+)$', num)
            if match:
                sep = match.group(1)
                level = heading.true_level
                level_sep_counts[level][sep] = level_sep_counts[level].get(sep, 0) + 1

        # 确定主风格
        main_style = max(style_counts, key=style_counts.get) if any(style_counts.values()) else 'numeric'
        
        # 确定每一级的多数派分隔符
        majority_formats = {}
        for level, counts in level_sep_counts.items():
            if counts:
                # 选取数量最多的作为格式
                majority_formats[level] = max(counts, key=counts.get)
            else:
                # 默认值
                if main_style == 'chinese' and level == 1:
                    majority_formats[level] = '、'
                else:
                    majority_formats[level] = '.' 
        
        return main_style, majority_formats

    def _detect_style(self, headings: List[HeadingInfo]) -> str:
        # 已过时，被 _detect_detailed_styles 替代
        style, _ = self._detect_detailed_styles(headings)
        return style
    
    def _generate_numbering(self, counters: Dict[int, int], level: int) -> str:
        """生成编号"""
        chinese_nums = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十', 
                        '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十']
        
        if self.style == 'chinese' and level == 1:
            num = counters[1]
            return chinese_nums[num] if num < len(chinese_nums) else str(num)
            
        elif self.style == 'chapter' and level == 1:
            num = counters[1]
            num_str = chinese_nums[num] if num < len(chinese_nums) else str(num)
            return f'第{num_str}章'
            
        else:
            # 默认数字型: 1.1.1
            parts = [str(counters[i]) for i in range(1, level + 1)]
            return '.'.join(parts)
    
    def _replace_numbering(self, original: str, new_numbering: str, content: str, level: int) -> str:
        """替换编号并应用多数派格式"""
        if not content:
            # 移除原有编号
            # 匹配: 数字/中文数字开头，后接符号或第...章/节
            text_without_numbering = re.sub(
                r'^(?:\d+(?:\.\d+)*[\.、\s\：\:]+|[一二三四五六七八九十百]+[\.、\s\：\:]+|第.+?[章节]\s*|\([一二三四五六七八九十\d]+\)[\.、\s\：\:]*|（[一二三四五六七八九十\d]+）[\.、\s\：\:]*)',
                '',
                original
            ).strip()
            content = text_without_numbering
        
        # 获取该等级的多数派分隔符
        sep = self.level_formats.get(level, '. ')
        if not sep.strip() and not sep: # 如果是空的
            sep = ' '
            
        # 拼接
        if self.style == 'chapter' and level == 1:
            return f'{new_numbering} {content}'
        
        return f'{new_numbering}{sep}{content}'


class HeadingHierarchyCorrector:
    """
    标题层次结构修正器
    自动补全缺失的中间层级
    """
    
    def __init__(self, numbering_style='auto'):
        self.style = numbering_style
        self.inserted_headings = []
    
    def correct_hierarchy(self, headings: List[HeadingInfo]) -> List[HeadingInfo]:
        """
        修正层次结构，补全缺失层级
        """
        corrected = []
        level_counters = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
        
        for heading in headings:
            current_level = heading.true_level
            
            # 检查是否需要插入中间层级
            if corrected:
                last_level = corrected[-1].true_level
                
                if current_level > last_level + 1:
                    # 插入缺失层级
                    for missing_level in range(last_level + 1, current_level):
                        placeholder = self._create_placeholder(
                            missing_level,
                            level_counters,
                            heading
                        )
                        corrected.append(placeholder)
                        self.inserted_headings.append(placeholder)
            
            corrected.append(heading)
            
            # 更新计数器
            level_counters[current_level] += 1
            for deeper in range(current_level + 1, 7):
                level_counters[deeper] = 0
        
        return corrected
    
    def _create_placeholder(self, level: int, counters: Dict, reference: HeadingInfo) -> HeadingInfo:
        """创建占位标题"""
        counters[level] += 1
        for deeper in range(level + 1, 7):
            counters[deeper] = 0
        
        # 生成编号
        if self.style == 'numeric':
            parts = [str(counters[i]) for i in range(1, level + 1)]
            numbering = '.'.join(parts)
        else:
            parts = [str(counters[i]) for i in range(1, level + 1)]
            numbering = '.'.join(parts)
        
        # 生成占位文本
        content = reference.content or "内容"
        level_names = {1: "章", 2: "节", 3: "小节", 4: "条目", 5: "项", 6: "款"}
        level_name = level_names.get(level, "部分")
        
        text = f'{numbering} {content}相关{level_name}'
        
        return HeadingInfo(
            index=-1,  # 占位符
            text=text,
            true_level=level,
            numbering=numbering,
            content=f'{content}相关{level_name}',
            is_placeholder=True,
            font_size={1: 22, 2: 18, 3: 16, 4: 14, 5: 12, 6: 12}.get(level, 12),
            is_bold=True,
        )


class HeadingProcessor:
    """
    智能标题处理器（主入口）
    整合所有功能模块
    """
    
    def __init__(self, numbering_style='auto'):
        """
        初始化标题处理器
        
        参数:
            numbering_style: 编号风格 ('auto', 'numeric', 'chinese', 'mixed')
        """
        self.numbering_style = numbering_style
        self.detector = HeadingDetector()
        self.level_engine = RelativeFormatLevelInferenceEngine()
        self.numbering_corrector = HeadingNumberingCorrector(numbering_style)
        self.hierarchy_corrector = HeadingHierarchyCorrector(numbering_style)
        self.report = ProcessingReport()
    
    def process_document(self, paragraphs: List[Dict], 
                        detect_headings: bool = True,
                        infer_levels: bool = True,
                        correct_numbering: bool = True,
                        correct_hierarchy: bool = True) -> Tuple[List[HeadingInfo], ProcessingReport]:
        """
        处理文档中的标题
        
        参数:
            paragraphs: 段落列表
            detect_headings: 是否检测标题
            infer_levels: 是否推断真实级别
            correct_numbering: 是否修正序号
            correct_hierarchy: 是否修正层次结构
            
        返回:
            (处理后的标题列表, 处理报告)
        """
        print("\n" + "="*60)
        print("智能标题处理")
        print("="*60)
        
        self.report.total_headings = len(paragraphs)
        
        # 步骤1: 检测标题
        if detect_headings:
            print("\n[1/4] 检测标题...")
            headings = self.detector.detect_headings(paragraphs)
            self.report.detected_headings = len(headings)
            print(f"  ✓ 检测到 {len(headings)} 个标题")
        else:
            # 使用已有的标题信息
            headings = [HeadingInfo(index=i, text=p['text']) 
                       for i, p in enumerate(paragraphs) if p.get('isHeading')]
            self.report.detected_headings = len(headings)
        
        # 步骤2: 推断真实级别
        if infer_levels:
            print("\n[2/4] 推断真实级别...")
            headings = self.level_engine.infer_levels(headings)
            
            # 统计样式错误
            style_mismatches = sum(1 for h in headings 
                                  if h.style_level and abs(h.true_level - h.style_level) >= 2)
            self.report.style_mismatches = style_mismatches
            if style_mismatches > 0:
                print(f"  ⚠ 发现 {style_mismatches} 处样式设置错误")
        
        # 步骤3: 修正序号
        if correct_numbering:
            print("\n[3/4] 修正标题序号...")
            headings = self.numbering_corrector.correct_numbering(headings)
            self.report.corrected_numbering = len(self.numbering_corrector.corrections)
            if self.report.corrected_numbering > 0:
                print(f"  ✓ 修正了 {self.report.corrected_numbering} 处序号错误")
        
        # 步骤4: 修正层次结构
        if correct_hierarchy:
            print("\n[4/4] 修正层次结构...")
            headings = self.hierarchy_corrector.correct_hierarchy(headings)
            self.report.inserted_placeholders = len(self.hierarchy_corrector.inserted_headings)
            if self.report.inserted_placeholders > 0:
                print(f"  ✓ 插入了 {self.report.inserted_placeholders} 个占位标题")
        
        print("\n" + "="*60)
        print("✓ 标题处理完成")
        print("="*60)
        
        return headings, self.report
    
    def generate_report(self, output_path: str = None) -> str:
        """
        生成详细的处理报告
        
        返回: 报告文本
        """
        report = []
        report.append("="*60)
        report.append("智能标题处理报告")
        report.append("="*60)
        report.append(f"处理时间: {self.report.timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"总段落数: {self.report.total_headings}")
        report.append(f"检测到标题: {self.report.detected_headings}")
        report.append("")
        
        # 序号修正
        if self.report.corrected_numbering > 0:
            report.append(f"序号修正: {self.report.corrected_numbering} 处")
            report.append("-"*60)
            for i, heading in enumerate(self.numbering_corrector.corrections, 1):
                report.append(f"{i}. 级别 {heading.true_level}")
                report.append(f"   原文: {heading.original_text}")
                report.append(f"   修正: {heading.text}")
                report.append("")
        
        # 层次补全
        if self.report.inserted_placeholders > 0:
            report.append(f"层次补全: {self.report.inserted_placeholders} 处")
            report.append("-"*60)
            for i, heading in enumerate(self.hierarchy_corrector.inserted_headings, 1):
                report.append(f"{i}. 级别 {heading.true_level} (占位标题)")
                report.append(f"   内容: {heading.text}")
                report.append(f"   说明: 请手动修改为实际内容")
                report.append("")
        
        # 样式错误
        if self.report.style_mismatches > 0:
            report.append(f"样式错误: {self.report.style_mismatches} 处")
            report.append("   (已自动修正)")
            report.append("")
        
        report.append("="*60)
        
        report_text = "\n".join(report)
        
        # 保存到文件
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
            print(f"✓ 报告已保存: {output_path}")
        
        return report_text


# 便捷函数
def process_single_document(doc_path: str, output_path: str = None, 
                           numbering_style: str = 'auto') -> ProcessingReport:
    """
    处理单个文档的标题
    
    参数:
        doc_path: 文档路径
        output_path: 输出路径（None则覆盖原文件）
        numbering_style: 编号风格
        
    返回:
        处理报告
    """
    # TODO: 实现文档读取和写入
    # 这里需要集成 ONLYOFFICE Document Builder
    pass


if __name__ == "__main__":
    # 测试代码
    print("智能标题处理模块已加载")
    print("可用类:")
    print("  - HeadingProcessor: 主处理器")
    print("  - HeadingDetector: 标题检测器")
    print("  - RelativeFormatLevelInferenceEngine: 基于格式相对关系的级别推断引擎")
    print("  - HeadingNumberingCorrector: 序号修正器")
    print("  - HeadingHierarchyCorrector: 层次修正器")
    print()
    
    # 测试 RelativeFormatLevelInferenceEngine
    print("=" * 60)
    print("测试: 基于格式相对关系的级别推断")
    print("=" * 60)
    
    # 模拟测试数据（明确设置 numbering_depth）
    test_headings = [
        HeadingInfo(index=0, text="第一章 概述", numbering="第一章", numbering_depth=1),
        HeadingInfo(index=1, text="1.1 背景介绍", numbering="1.1", numbering_depth=2),
        HeadingInfo(index=2, text="第一章 目标", numbering="第一章", numbering_depth=1),
        HeadingInfo(index=3, text="1.1 技术路线", numbering="1.1", numbering_depth=2),
        HeadingInfo(index=4, text="1.1.1 具体方法", numbering="1.1.1", numbering_depth=3),
        HeadingInfo(index=5, text="1.1 总结", numbering="1.1", numbering_depth=2),
    ]
    
    engine = RelativeFormatLevelInferenceEngine()
    engine.infer_levels(test_headings)
    
    print("\n格式栈:", engine.format_stack)
    print("\n结果:")
    print("-" * 60)
    for h in test_headings:
        print(f"  标题: {h.text:<20} | 深度: {h.numbering_depth} | 级别: {h.true_level}")
    print()
    
    print("预期结果:")
    print("  第一章 (深度1) → 1级")
    print("  1.1 (深度2) → 2级")
    print("  第一章 (深度1) → 1级 (已知)")
    print("  1.1 (深度2) → 2级 (已知)")
    print("  1.1.1 (深度3) → 3级")
    print("  1.1 (深度2) → 2级 (已知)")
