"""
numbering_resolver.py
─────────────────────
解析 Word DOCX 文件中的 numbering.xml（宏文件），将段落的
(numId, ilvl) 对转换为实际渲染的序号字符串（如 "1.1."、"第二章"）。

Word 自动编号架构：
  document.xml  ← 段落只记录 <w:numPr><w:numId w:val="N"/><w:ilvl w:val="L"/></w:numPr>
  numbering.xml ← 定义：
      <w:abstractNum>   编号格式模板
          <w:lvl w:ilvl="0">  每层格式
              <w:numFmt>      格式类型（decimal / chineseCounting / ...）
              <w:lvlText>     显示模板（如 "%1." / "第%1章"）
      <w:num>  实例，绑定 abstractNum

用法：
    resolver = NumberingResolver("path/to/doc.docx")

    # 在段落解析循环中
    num_id = ...  # 从 GetNumPr().GetNumId().GetVal().ToInt()
    ilvl   = ...  # 从 GetNumPr().GetIlvl().GetVal().ToInt()

    # 获取该段落渲染后的序号文字（并推进计数器）
    numbering_text = resolver.resolve(num_id, ilvl)
    # e.g. "1.1."  "第二章"  ""（如果解析失败）

    # 告知 resolver 文档解析完毕（重置计数器供下一次使用）
    resolver.reset()
"""

import zipfile
import re
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

# OOXML 命名空间
W  = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

# 支持的数字格式 -> 渲染函数
def _to_decimal(n: int) -> str:
    return str(n)

def _to_chinese(n: int) -> str:
    """将整数转为中文数字（支持 1-99）"""
    units  = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    tens   = ["", "十", "二十", "三十", "四十", "五十", "六十", "七十", "八十", "九十"]
    if n <= 0:  return ""
    if n < 10:  return units[n]
    if n < 20:  return ("十" if n == 10 else "十" + units[n % 10])
    if n < 100: return tens[n // 10] + units[n % 10]
    return str(n)   # fallback

def _to_upper_letter(n: int) -> str:
    return chr(ord('A') + (n - 1) % 26)

def _to_lower_letter(n: int) -> str:
    return chr(ord('a') + (n - 1) % 26)

def _to_roman_upper(n: int) -> str:
    vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1]
    syms = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"]
    result = ""
    for v, s in zip(vals, syms):
        while n >= v:
            result += s
            n -= v
    return result

def _to_roman_lower(n: int) -> str:
    return _to_roman_upper(n).lower()

# numFmt 值 -> 渲染函数
_NUMFMT_RENDERERS = {
    "decimal":                          _to_decimal,
    "chineseCounting":                  _to_chinese,
    "chineseCountingThousand":          _to_chinese,
    "japaneseCounting":                 _to_chinese,  # Word 经常用此格式表示中文数字
    "japaneseCountingThousand":         _to_chinese,
    "chineseLegalSimplified":           _to_chinese,
    "ideographTraditional":             _to_chinese,
    "ideographZodiac":                  _to_chinese,
    "upperLetter":                      _to_upper_letter,
    "lowerLetter":                      _to_lower_letter,
    "upperRoman":                       _to_roman_upper,
    "lowerRoman":                       _to_roman_lower,
    "decimalZero":                      lambda n: str(n).zfill(2),
    "ordinal":                          lambda n: f"{n}.",
    "none":                             lambda n: "",
    "bullet":                           lambda n: "",
}


class _LevelDef:
    """一个编号层级的定义"""
    __slots__ = ("ilvl", "num_fmt", "lvl_text", "start_val", "is_legal_style")

    def __init__(self):
        self.ilvl: int         = 0
        self.num_fmt: str      = "decimal"
        self.lvl_text: str     = "%1."
        self.start_val: int    = 1
        self.is_legal_style: bool = False   # legalNumStyle


class _AbstractNum:
    """一个 abstractNum 定义"""
    def __init__(self):
        self.levels: Dict[int, _LevelDef] = {}

    def get_level(self, ilvl: int) -> Optional[_LevelDef]:
        return self.levels.get(ilvl)


class NumberingResolver:
    """
    解析 Word DOCX 的 numbering.xml（及关联 styles.xml），
    提供 resolve(numId, ilvl) -> str 接口。

    计数器状态在整个文档解析期间维护，使用 reset() 重置。
    """

    def __init__(self, docx_path: str):
        self.docx_path = docx_path
        self._abstract_nums: Dict[int, _AbstractNum] = {}
        # num_id -> (abstractNumId, lvl_overrides: dict[ilvl, _LevelDef])
        self._nums: Dict[int, Tuple[int, Dict[int, _LevelDef]]] = {}
        # 运行时计数器：num_id -> {ilvl: count}
        self._counters: Dict[int, Dict[int, int]] = {}
        # 当前每个 num_id 的活跃层级（用于在层级下降时重置子计数器）
        self._active_ilvl: Dict[int, int] = {}

        self._loaded = False
        self._load(docx_path)

    # ─────────────────────────────────────────────
    # 公共接口
    # ─────────────────────────────────────────────

    def resolve(self, num_id: int, ilvl: int) -> str:
        """
        推进 (num_id, ilvl) 的计数器，返回渲染后的序号字符串。
        如果解析失败，返回空字符串 ""。
        """
        if not self._loaded:
            return ""
        if num_id not in self._nums:
            return ""

        abstract_id, overrides = self._nums[num_id]
        abstract = self._abstract_nums.get(abstract_id)
        if abstract is None:
            return ""

        # 获取该层定义（先查 override，再查 abstractNum）
        level_def = overrides.get(ilvl) or abstract.get_level(ilvl)
        if level_def is None:
            return ""

        # 初始化计数器
        if num_id not in self._counters:
            self._counters[num_id] = {}
        counters = self._counters[num_id]

        # 如果层级下降（e.g. 从 ilvl=2 回到 ilvl=1），重置更深层的计数器
        prev_ilvl = self._active_ilvl.get(num_id, -1)
        if ilvl < prev_ilvl:
            for deeper in range(ilvl + 1, 10):
                counters.pop(deeper, None)
        self._active_ilvl[num_id] = ilvl

        # 推进当前层计数
        if ilvl not in counters:
            counters[ilvl] = level_def.start_val
        else:
            counters[ilvl] += 1

        # 渲染 lvl_text 模板（如 "%1.%2." / "第%1章"）
        return self._render_lvl_text(level_def.lvl_text, ilvl, counters, num_id, abstract, overrides)

    def reset(self):
        """重置所有运行时计数器（在开始解析新文档时调用）"""
        self._counters.clear()
        self._active_ilvl.clear()

    def has_numbering(self) -> bool:
        """文档是否包含 numbering.xml"""
        return self._loaded

    def get_num_ids(self) -> List[int]:
        """返回已解析的所有 numId 列表"""
        return list(self._nums.keys())

    # ─────────────────────────────────────────────
    # 内部解析
    # ─────────────────────────────────────────────

    def _load(self, docx_path: str):
        """从 DOCX（ZIP）中读取并解析 numbering.xml"""
        try:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                names = zf.namelist()

                if 'word/numbering.xml' not in names:
                    # 有些文档没有自动编号
                    self._loaded = False
                    return

                numbering_xml = zf.read('word/numbering.xml')
                self._parse_numbering(numbering_xml)
                self._loaded = True

        except Exception as e:
            print(f"  [NumberingResolver] 加载 numbering.xml 失败: {e}")
            self._loaded = False

    def _parse_numbering(self, xml_bytes: bytes):
        """解析 numbering.xml 的 XML"""
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError as e:
            print(f"  [NumberingResolver] XML 解析失败: {e}")
            return

        # 1. 解析所有 abstractNum
        for abstract_elem in root.findall(f"{W}abstractNum"):
            abstract_id_elem = abstract_elem.get(f"{W}abstractNumId")
            if abstract_id_elem is None:
                continue
            abstract_id = int(abstract_id_elem)
            abstract = _AbstractNum()

            for lvl_elem in abstract_elem.findall(f"{W}lvl"):
                ilvl_str = lvl_elem.get(f"{W}ilvl")
                if ilvl_str is None:
                    continue
                lvl = _LevelDef()
                lvl.ilvl = int(ilvl_str)

                start_elem = lvl_elem.find(f"{W}start")
                if start_elem is not None:
                    val = start_elem.get(f"{W}val")
                    if val: lvl.start_val = int(val)

                fmt_elem = lvl_elem.find(f"{W}numFmt")
                if fmt_elem is not None:
                    val = fmt_elem.get(f"{W}val")
                    if val: lvl.num_fmt = val

                text_elem = lvl_elem.find(f"{W}lvlText")
                if text_elem is not None:
                    val = text_elem.get(f"{W}val")
                    if val is not None:
                        lvl.lvl_text = val

                abstract.levels[lvl.ilvl] = lvl

            self._abstract_nums[abstract_id] = abstract

        # 2. 解析所有 num（绑定 abstractNum，包含 override）
        for num_elem in root.findall(f"{W}num"):
            num_id_str = num_elem.get(f"{W}numId")
            if num_id_str is None:
                continue
            num_id = int(num_id_str)

            abstract_ref_elem = num_elem.find(f"{W}abstractNumId")
            if abstract_ref_elem is None:
                continue
            abstract_id_str = abstract_ref_elem.get(f"{W}val")
            if abstract_id_str is None:
                continue
            abstract_id = int(abstract_id_str)

            # 解析层级覆盖（lvlOverride）
            overrides: Dict[int, _LevelDef] = {}
            for override_elem in num_elem.findall(f"{W}lvlOverride"):
                ilvl_str = override_elem.get(f"{W}ilvl")
                if ilvl_str is None:
                    continue
                ilvl = int(ilvl_str)

                # startOverride（重置起始值）
                start_override = override_elem.find(f"{W}startOverride")
                
                # 完整 lvl 定义覆盖
                lvl_override_elem = override_elem.find(f"{W}lvl")

                if lvl_override_elem is not None:
                    lvl = _LevelDef()
                    lvl.ilvl = ilvl
                    start_el = lvl_override_elem.find(f"{W}start")
                    if start_el is not None:
                        v = start_el.get(f"{W}val")
                        if v: lvl.start_val = int(v)
                    fmt_el = lvl_override_elem.find(f"{W}numFmt")
                    if fmt_el is not None:
                        v = fmt_el.get(f"{W}val")
                        if v: lvl.num_fmt = v
                    text_el = lvl_override_elem.find(f"{W}lvlText")
                    if text_el is not None:
                        v = text_el.get(f"{W}val")
                        if v is not None: lvl.lvl_text = v
                    overrides[ilvl] = lvl
                elif start_override is not None:
                    # 只覆盖起始值，沿用 abstractNum 的其他定义
                    v = start_override.get(f"{W}val")
                    if v:
                        # 复制一份 abstractNum 的层级定义并修改 start_val
                        abstract = self._abstract_nums.get(abstract_id)
                        if abstract:
                            orig = abstract.get_level(ilvl)
                            if orig:
                                import copy
                                lvl_copy = copy.copy(orig)
                                lvl_copy.start_val = int(v)
                                overrides[ilvl] = lvl_copy

            self._nums[num_id] = (abstract_id, overrides)

    def _render_lvl_text(
        self,
        lvl_text: str,
        current_ilvl: int,
        counters: Dict[int, int],
        num_id: int,
        abstract: _AbstractNum,
        overrides: Dict[int, _LevelDef],
    ) -> str:
        """
        将 lvlText 模板（如 "%1.%2."）替换为实际数字。
        %N 代表第 N-1 层（0-indexed ilvl = N-1）的计数值。
        """
        if not lvl_text:
            return ""

        result = lvl_text

        # 找所有 %N 占位符
        for match in re.finditer(r'%(\d+)', lvl_text):
            placeholder = match.group(0)       # 如 "%1"
            ref_ilvl = int(match.group(1)) - 1 # %1 -> ilvl=0

            # 获取该层的格式定义
            level_def = overrides.get(ref_ilvl) or abstract.get_level(ref_ilvl)
            num_fmt = level_def.num_fmt if level_def else "decimal"

            # 获取计数值
            count = counters.get(ref_ilvl, 1)

            # 应用格式
            renderer = _NUMFMT_RENDERERS.get(num_fmt, _to_decimal)
            rendered = renderer(count)

            result = result.replace(placeholder, rendered, 1)

        return result


# ─────────────────────────────────────────────────────────────────────────────
# 便捷函数：从 DOCX XML（paragraph element）直接读取 numId / ilvl
# 供 multi_copy_merger 调用（无需 DocBuilder API）
# ─────────────────────────────────────────────────────────────────────────────

def extract_num_pr_from_xml(para_xml: str) -> Tuple[Optional[int], Optional[int]]:
    """
    从段落 XML 片段中提取 (numId, ilvl)。
    返回 (None, None) 如果段落没有自动编号。
    """
    try:
        root = ET.fromstring(para_xml)
        pPr = root.find(f"{W}pPr")
        if pPr is None:
            return None, None
        num_pr = pPr.find(f"{W}numPr")
        if num_pr is None:
            return None, None
        ilvl_el = num_pr.find(f"{W}ilvl")
        num_id_el = num_pr.find(f"{W}numId")
        ilvl = int(ilvl_el.get(f"{W}val")) if ilvl_el is not None else None
        num_id = int(num_id_el.get(f"{W}val")) if num_id_el is not None else None
        return num_id, ilvl
    except Exception:
        return None, None


# ─────────────────────────────────────────────────────────────────────────────
# 命令行调试工具
# ─────────────────────────────────────────────────────────────────────────────

def _print_numbering_summary(docx_path: str):
    """打印文档中所有编号定义的摘要（调试用）"""
    resolver = NumberingResolver(docx_path)
    if not resolver.has_numbering():
        print("该文档没有 numbering.xml（无自动编号）")
        return

    print(f"=== {docx_path} 编号定义摘要 ===")
    print(f"abstractNum 数量: {len(resolver._abstract_nums)}")
    print(f"num（实例）数量:  {len(resolver._nums)}")
    print()

    for num_id, (abstract_id, overrides) in sorted(resolver._nums.items()):
        abstract = resolver._abstract_nums.get(abstract_id)
        print(f"  numId={num_id}  →  abstractNumId={abstract_id}")
        if abstract:
            for ilvl in sorted(abstract.levels.keys()):
                lvl = overrides.get(ilvl) or abstract.levels[ilvl]
                print(f"    ilvl={ilvl}: fmt={lvl.num_fmt:<20} start={lvl.start_val}  template={repr(lvl.lvl_text)}")
        print()

    print("=== 模拟渲染前10个序号 ===")
    for num_id in sorted(resolver._nums.keys())[:3]:
        print(f"\n  numId={num_id}:")
        resolver.reset()
        for i in range(5):
            for ilvl in range(min(3, 9)):
                text = resolver.resolve(num_id, ilvl)
                if text:
                    print(f"    ilvl={ilvl} -> '{text}'")
                    break


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python numbering_resolver.py <docx_path>")
        sys.exit(1)
    _print_numbering_summary(sys.argv[1])
