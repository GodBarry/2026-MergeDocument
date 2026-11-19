"""
图片序号自动标注模块（优化版）
功能：
1. 检测文档中的图片（支持有无编号的图注）
2. 识别图注（传统规则 + AI 辅助）
3. 自动修正图片编号
4. 支持多种编号风格和空格变体
"""

import re
import os
import sys
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass

# 添加 Document Builder 路径
DOCBUILDER_PATH = r"D:\onlyoffice_documentbuilder"
if DOCBUILDER_PATH not in sys.path:
    sys.path.insert(0, DOCBUILDER_PATH)


@dataclass
class FigureInfo:
    """图片信息"""
    index: int                      # 在文档中的索引
    caption_text: str               # 图注文本
    numbering: Optional[str] = None # 编号（如 "1", "1.1"）
    caption_index: int = -1         # 图注段落索引
    has_numbering: bool = False     # 是否有编号
    is_corrected: bool = False      # 是否已修正
    original_caption: str = ""      # 原始图注
    chapter: int = 1                # 所属章节
    detection_method: str = "rule"  # 检测方法：rule 或 ai
    confidence: float = 1.0         # 置信度


class FigureNumberingPatterns:
    """
    图注编号模式识别器（优化版）
    支持各种空格变体和全角数字
    """
    
    # 图注模式（中文）- 支持各种空格变体
    CHINESE_PATTERNS = [
        # 标准格式（支持任意空格）
        (r'^图\s*(\d+)', 'simple'),              # 图1, 图 1, 图  1
        (r'^图\s*(\d+\.\d+)', 'chapter'),        # 图1.1, 图 1.1
        (r'^图\s*(\d+-\d+)', 'section'),         # 图1-1, 图 1-1
        (r'^图片\s*(\d+)', 'simple'),            # 图片1, 图片 1
        (r'^插图\s*(\d+)', 'simple'),            # 插图1, 插图 1
        
        # 带冒号的变体
        (r'^图\s*[：:]\s*(\d+)', 'simple'),      # 图：1, 图: 1
        (r'^图\s*[：:]\s*(\d+\.\d+)', 'chapter'), # 图：1.1
        (r'^图\s*[：:]\s*(\d+-\d+)', 'section'),  # 图：1-1
        
        # 全角数字
        (r'^图\s*([０-９]+)', 'simple'),         # 图１
        (r'^图\s*([０-９]+\.[０-９]+)', 'chapter'), # 图１.１
    ]
    
    # 图注模式（英文）
    ENGLISH_PATTERNS = [
        (r'^Figure\s+(\d+)', 'simple'),          # Figure 1
        (r'^Figure\s+(\d+\.\d+)', 'chapter'),    # Figure 1.1
        (r'^Fig\.\s*(\d+)', 'simple'),           # Fig. 1, Fig.1
        (r'^Fig\s+(\d+)', 'simple'),             # Fig 1
        (r'^Figure\s*[：:]\s*(\d+)', 'simple'),  # Figure: 1
        (r'^Fig\.\s*[：:]\s*(\d+)', 'simple'),   # Fig.: 1
    ]
    
    # 无编号图注关键词（用于识别没有序号的图注）
    CAPTION_KEYWORDS = [
        # 中文
        '图', '图片', '插图', '示意图', '流程图', '架构图', '结构图',
        '界面图', '截图', '效果图', '设计图', '原理图',
        # 英文
        'Figure', 'Fig.', 'Fig', 'Image', 'Diagram', 'Chart',
        'Screenshot', 'Illustration', 'Schematic'
    ]
    
    @classmethod
    def extract_numbering(cls, text: str) -> Tuple[Optional[str], Optional[str], str]:
        """
        提取图注编号
        
        返回: (numbering, style, content)
        """
        text = text.strip()
        
        # 尝试中文模式
        for pattern, style in cls.CHINESE_PATTERNS:
            match = re.match(pattern, text, re.IGNORECASE)
            if match:
                numbering = match.group(1)
                # 转换全角数字为半角
                numbering = cls._normalize_number(numbering)
                content = text[match.end():].strip()
                # 移除可能的冒号、空格、破折号等
                content = re.sub(r'^[：:\s\-—]+', '', content)
                return (numbering, style, content)
        
        # 尝试英文模式
        for pattern, style in cls.ENGLISH_PATTERNS:
            match = re.match(pattern, text, re.IGNORECASE)
            if match:
                numbering = match.group(1)
                content = text[match.end():].strip()
                content = re.sub(r'^[：:\s\-—]+', '', content)
                return (numbering, style, content)
        
        return (None, None, text)
    
    @classmethod
    def _normalize_number(cls, num_str: str) -> str:
        """转换全角数字为半角"""
        fullwidth = '０１２３４５６７８９'
        halfwidth = '0123456789'
        trans_table = str.maketrans(fullwidth, halfwidth)
        return num_str.translate(trans_table)
    
    @classmethod
    def is_caption_keyword(cls, text: str) -> bool:
        """检查文本是否包含图注关键词"""
        text_lower = text.lower()
        # 检查是否以关键词开头（更严格）
        for keyword in cls.CAPTION_KEYWORDS:
            if text_lower.startswith(keyword.lower()):
                return True
        return False
    
    @classmethod
    def detect_style(cls, figures: List[FigureInfo]) -> str:
        """检测主要编号风格"""
        style_counts = {'simple': 0, 'chapter': 0, 'section': 0}
        
        for fig in figures:
            if fig.has_numbering:
                numbering = fig.numbering
                if '.' in numbering:
                    style_counts['chapter'] += 1
                elif '-' in numbering:
                    style_counts['section'] += 1
                else:
                    style_counts['simple'] += 1
        
        if not any(style_counts.values()):
            return 'simple'
        
        return max(style_counts, key=style_counts.get)


class FigureDetector:
    """
    图片检测器（双模式：传统规则 + AI 辅助）
    从文档中检测图片和图注
    """
    
    def __init__(self, ai_analyzer=None, use_ai=True):
        self.ai_analyzer = ai_analyzer
        self.use_ai = use_ai and (ai_analyzer is not None)
        self.pattern_detector = FigureNumberingPatterns()
    
    def detect_figures(self, document_structure: Dict) -> List[FigureInfo]:
        """
        检测文档中的图片
        
        参数:
            document_structure: 文档结构（包含段落、图片等）
        
        返回:
            图片信息列表
        """
        figures = []
        paragraphs = document_structure.get('paragraphs', [])
        
        print(f"\n检测模式: {'传统规则 + AI 辅助' if self.use_ai else '仅传统规则'}")
        
        # 查找所有包含图注的段落
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            
            if not text:
                continue
            
            # 方法1: 传统规则检测
            is_caption_rule, confidence_rule = self._is_figure_caption_rule(text, para)
            
            # 方法2: AI 辅助检测（仅对边界情况）
            is_caption_ai = False
            confidence_ai = 0.0
            
            if self.use_ai and not is_caption_rule:
                # 只对可能是图注但规则不确定的情况使用 AI
                if self.pattern_detector.is_caption_keyword(text):
                    result = self._ai_detect_caption(text, para)
                    if result:
                        is_caption_ai = result.get('is_caption', False)
                        confidence_ai = result.get('confidence', 0.0)
            
            # 综合判断
            is_caption = is_caption_rule or (is_caption_ai and confidence_ai > 0.7)
            
            if is_caption:
                numbering, style, content = self.pattern_detector.extract_numbering(text)
                
                figure = FigureInfo(
                    index=len(figures),
                    caption_text=text,
                    caption_index=i,
                    numbering=numbering,
                    has_numbering=(numbering is not None),
                    original_caption=text,
                    detection_method='ai' if is_caption_ai else 'rule',
                    confidence=max(confidence_rule, confidence_ai)
                )
                
                # 推断所属章节
                figure.chapter = self._infer_chapter(i, paragraphs)
                
                figures.append(figure)
        
        # 统计检测方法
        rule_count = sum(1 for f in figures if f.detection_method == 'rule')
        ai_count = sum(1 for f in figures if f.detection_method == 'ai')
        
        print(f"  传统规则检测: {rule_count} 个")
        if self.use_ai:
            print(f"  AI 辅助检测: {ai_count} 个")
        
        return figures
    
    def _is_figure_caption_rule(self, text: str, para: Dict) -> Tuple[bool, float]:
        """
        传统规则判断是否为图注
        
        返回: (is_caption, confidence)
        """
        score = 0
        max_score = 100
        
        # 规则1: 包含图注关键词（40分）
        has_keyword = self.pattern_detector.is_caption_keyword(text)
        if has_keyword:
            score += 40
        
        # 规则2: 有编号模式（50分）
        numbering, _, _ = self.pattern_detector.extract_numbering(text)
        if numbering:
            score += 50
        
        # 规则3: 文本较短（10分）
        if len(text) < 200:
            score += 10
        
        # 规则4: 居中对齐（加分项，5分）
        if para.get('is_centered', False):
            score += 5
        
        # 规则5: 字体较小（图注通常比正文小，5分）
        font_size = para.get('font_size', 12)
        if font_size <= 10:
            score += 5
        
        # 置信度
        confidence = min(score / max_score, 1.0)
        
        # 阈值判断（60分以上认为是图注）
        is_caption = score >= 60
        
        return (is_caption, confidence)
    
    def _infer_chapter(self, para_index: int, paragraphs: List[Dict]) -> int:
        """推断图片所属章节"""
        # 向前查找最近的一级标题
        for i in range(para_index - 1, -1, -1):
            para = paragraphs[i]
            if para.get('isHeading') and para.get('true_level') == 1:
                # 尝试从标题中提取章节号
                text = para.get('text', '')
                # 匹配各种章节号格式
                patterns = [
                    r'^(?:第)?(\d+)',           # 第1章, 1
                    r'^Chapter\s+(\d+)',        # Chapter 1
                    r'^(\d+)\.',                # 1.
                ]
                for pattern in patterns:
                    match = re.match(pattern, text, re.IGNORECASE)
                    if match:
                        return int(match.group(1))
        
        return 1  # 默认第一章
    
    def _ai_detect_caption(self, text: str, para: Dict) -> Optional[Dict]:
        """
        使用 AI 检测图注
        
        注意：AI 是大语言模型，只能分析文本，不能识别图像
        """
        if not self.ai_analyzer:
            return None
        
        prompt = f"""判断以下文本是否为图片的图注（caption）。

注意：
1. 你是文本分析模型，只能分析文本内容，不能看到图像
2. 图注通常包含"图"、"Figure"等关键词
3. 图注可能有编号（如"图1"），也可能没有（如"系统架构图"）
4. 图注通常简短、描述性强

文本: "{text}"

特征:
- 居中: {para.get('is_centered', False)}
- 长度: {len(text)}字符
- 字体: {para.get('font_size', 12)}pt

请判断这是否为图注，返回 JSON:
{{
  "is_caption": true/false,
  "confidence": 0.0-1.0,
  "reasoning": "理由（说明为什么是或不是图注）"
}}
"""
        
        try:
            response = self.ai_analyzer.provider.complete(prompt)
            # 提取 JSON
            import json
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            if json_start >= 0 and json_end > json_start:
                result = json.loads(response[json_start:json_end])
                return result
        except Exception as e:
            print(f"  ⚠ AI 检测失败: {e}")
        
        return None


class FigureNumberingCorrector:
    """
    图片编号修正器
    """
    
    def __init__(self, numbering_style='auto', language='chinese'):
        self.style = numbering_style
        self.language = language
        self.corrections = []
        self.added_numbering = []  # 记录新增编号的图注
    
    def correct_numbering(self, figures: List[FigureInfo]) -> List[FigureInfo]:
        """
        修正图片编号
        
        参数:
            figures: 图片列表
        
        返回:
            修正后的图片列表
        """
        # 检测编号风格
        if self.style == 'auto':
            self.style = FigureNumberingPatterns.detect_style(figures)
        
        print(f"\n编号风格: {self.style}")
        print(f"语言: {self.language}")
        
        # 重新编号
        counter = 0
        chapter_counters = {}  # {chapter: counter}
        
        for figure in figures:
            chapter = figure.chapter
            
            if self.style == 'simple':
                # 连续编号
                counter += 1
                new_numbering = str(counter)
            
            elif self.style == 'chapter':
                # 按章节编号
                if chapter not in chapter_counters:
                    chapter_counters[chapter] = 0
                chapter_counters[chapter] += 1
                new_numbering = f"{chapter}.{chapter_counters[chapter]}"
            
            elif self.style == 'section':
                # 按节编号
                if chapter not in chapter_counters:
                    chapter_counters[chapter] = 0
                chapter_counters[chapter] += 1
                new_numbering = f"{chapter}-{chapter_counters[chapter]}"
            
            else:
                # 默认连续编号
                counter += 1
                new_numbering = str(counter)
            
            # 生成新图注
            original_caption = figure.caption_text
            _, _, content = FigureNumberingPatterns.extract_numbering(original_caption)
            
            # 如果没有内容，尝试提取
            if not content:
                # 移除可能的关键词
                content = re.sub(
                    r'^(?:图|图片|插图|Figure|Fig\.?)\s*[：:\s]*',
                    '',
                    original_caption,
                    flags=re.IGNORECASE
                ).strip()
                
                # 如果还是空的，使用默认描述
                if not content:
                    content = "图片说明"
            
            # 生成新图注
            if self.language == 'chinese':
                new_caption = f"图{new_numbering} {content}"
            else:
                new_caption = f"Figure {new_numbering} {content}"
            
            # 记录修改
            if new_caption != original_caption:
                figure.caption_text = new_caption
                figure.numbering = new_numbering
                figure.is_corrected = True
                
                if not figure.has_numbering:
                    # 原本没有编号，现在添加了
                    self.added_numbering.append(figure)
                else:
                    # 修正了编号
                    self.corrections.append(figure)
        
        return figures
    
    def generate_report(self) -> str:
        """生成修正报告"""
        total_changes = len(self.corrections) + len(self.added_numbering)
        
        if total_changes == 0:
            return "✓ 所有图片编号正确，无需修正"
        
        report = []
        report.append("="*60)
        report.append("图片编号修正报告")
        report.append("="*60)
        report.append(f"共处理 {total_changes} 处图片编号\n")
        
        # 修正的编号
        if self.corrections:
            report.append(f"修正编号: {len(self.corrections)} 处")
            report.append("-"*60)
            for i, fig in enumerate(self.corrections, 1):
                report.append(f"{i}. 图片 {fig.index + 1} (检测方法: {fig.detection_method})")
                report.append(f"   原图注: {fig.original_caption}")
                report.append(f"   新图注: {fig.caption_text}")
                report.append(f"   编号: {fig.numbering}")
                report.append("")
        
        # 新增的编号
        if self.added_numbering:
            report.append(f"\n新增编号: {len(self.added_numbering)} 处")
            report.append("-"*60)
            for i, fig in enumerate(self.added_numbering, 1):
                report.append(f"{i}. 图片 {fig.index + 1} (检测方法: {fig.detection_method})")
                report.append(f"   原图注: {fig.original_caption} (无编号)")
                report.append(f"   新图注: {fig.caption_text}")
                report.append(f"   编号: {fig.numbering}")
                report.append("")
        
        report.append("="*60)
        return "\n".join(report)


class FigureProcessor:
    """
    图片处理器（主入口）
    """
    
    def __init__(self, numbering_style='auto', language='chinese', 
                 ai_analyzer=None, use_ai=True):
        self.detector = FigureDetector(ai_analyzer, use_ai)
        self.corrector = FigureNumberingCorrector(numbering_style, language)
        self.ai_analyzer = ai_analyzer
        self.use_ai = use_ai
    
    def process_figures(self, document_structure: Dict) -> Tuple[List[FigureInfo], str]:
        """
        处理文档中的图片
        
        返回:
            (图片列表, 处理报告)
        """
        print("\n" + "="*60)
        print("图片编号处理")
        print("="*60)
        
        # 步骤1: 检测图片
        print("\n[1/2] 检测图片...")
        figures = self.detector.detect_figures(document_structure)
        
        with_numbering = sum(1 for f in figures if f.has_numbering)
        without_numbering = len(figures) - with_numbering
        
        print(f"  ✓ 检测到 {len(figures)} 个图片")
        print(f"    - 有编号: {with_numbering} 个")
        print(f"    - 无编号: {without_numbering} 个")
        
        # 步骤2: 修正编号
        print("\n[2/2] 修正图片编号...")
        figures = self.corrector.correct_numbering(figures)
        
        total_changes = len(self.corrector.corrections) + len(self.corrector.added_numbering)
        if total_changes > 0:
            print(f"  ✓ 修正了 {len(self.corrector.corrections)} 处编号")
            print(f"  ✓ 新增了 {len(self.corrector.added_numbering)} 处编号")
        else:
            print("  ✓ 所有编号正确")
        
        print("\n" + "="*60)
        print("✓ 图片处理完成")
        print("="*60)
        
        # 生成报告
        report = self.corrector.generate_report()
        
        return figures, report


# 便捷函数
def process_document_figures(document_structure: Dict,
                            numbering_style: str = 'auto',
                            language: str = 'chinese',
                            ai_analyzer=None,
                            use_ai: bool = True) -> Tuple[List[FigureInfo], str]:
    """
    处理文档中的图片编号
    
    参数:
        document_structure: 文档结构
        numbering_style: 编号风格（auto, simple, chapter, section）
        language: 语言（chinese, english）
        ai_analyzer: AI 分析器（可选）
        use_ai: 是否使用 AI 辅助（默认 True，需要 ai_analyzer）
    
    返回:
        (图片列表, 处理报告)
    """
    processor = FigureProcessor(numbering_style, language, ai_analyzer, use_ai)
    return processor.process_figures(document_structure)


if __name__ == "__main__":
    print("图片序号自动标注模块已加载（优化版）")
    print("\n新特性:")
    print("  ✓ 支持空格变体（图1, 图 1, 图  1）")
    print("  ✓ 支持全角数字（图１）")
    print("  ✓ 支持无编号图注（自动添加编号）")
    print("  ✓ 双模式检测（传统规则 + AI 辅助）")
    print("  ✓ AI 只分析文本，不识别图像")
    
    print("\n可用类:")
    print("  - FigureProcessor: 主处理器")
    print("  - FigureDetector: 图片检测器（双模式）")
    print("  - FigureNumberingCorrector: 编号修正器")
    
    # 使用示例
    example = """
# 使用示例:

from .figure_numbering import FigureProcessor
from ai_heading_analyzer import create_ai_analyzer

# 方式1: 仅使用传统规则
processor = FigureProcessor(
    numbering_style='chapter',
    language='chinese',
    use_ai=False  # 不使用 AI
)

# 方式2: 传统规则 + AI 辅助
ai_analyzer = create_ai_analyzer(
    api_key="your-deepseek-key",
    provider="deepseek"
)

processor = FigureProcessor(
    numbering_style='chapter',
    language='chinese',
    ai_analyzer=ai_analyzer,
    use_ai=True  # 使用 AI 辅助
)

# 处理文档
figures, report = processor.process_figures(document_structure)

print(report)

# 输出示例:
# ============================================================
# 图片编号修正报告
# ============================================================
# 共处理 5 处图片编号
#
# 修正编号: 3 处
# ------------------------------------------------------------
# 1. 图片 1 (检测方法: rule)
#    原图注: 图 3 系统架构图
#    新图注: 图1.1 系统架构图
#    编号: 1.1
#
# 新增编号: 2 处
# ------------------------------------------------------------
# 1. 图片 4 (检测方法: ai)
#    原图注: 系统架构示意图 (无编号)
#    新图注: 图2.1 系统架构示意图
#    编号: 2.1
"""
    
    print(example)
