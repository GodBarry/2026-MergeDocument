"""
Anchor scoring engine for robust merge alignment.
"""

import re
import difflib
from typing import List, Optional


def _anchor_norm_text(s: str) -> str:
    """
    标准化锚点文本，用于标题匹配。
    参考 H:\\Projects\\mege_docs_zbh\\multi_copy_merger.py 的 normalize_para 函数。

    策略：去掉标题编号前缀，但保留足够的位置信息来区分同名标题。
    例如："1. 主要研究内容" 和 "1. 主要研究内容"（不同章节的同名小标题）
    通过前/后标题上下文来区分，不在本函数处理。
    """
    if not s:
        return ""
    # 第一步：去掉标题编号前缀（参考旧项目 normalize_para）
    # 支持：数字点号(1.1/1、)、中文数字(一、二)、第X章/节、括号编号((一)/(1))
    _NUMBERING_PATTERNS = [
        r'\d+(?:\.\d+)*[\.、，,\s\：\:]+',       # 数字型: 1. / 1、 / 1 ,
        r'[一二三四五六七八九十百]+[\.、，,\s\：\:]+',  # 中文数字型: 一、 / 一 、
        r'第.+?[章节篇编]\s*',                    # 第X章/节/篇/编
        r'\([一二三四五六七八九十\d]+\)[\.、，,\s\：\:]*',  # (一) / (1)
        r'（[一二三四五六七八九十\d]+）[\.、，,\s\：\:]*',  # （一） /
    ]
    for pat in _NUMBERING_PATTERNS:
        s = re.sub(pat, '', s)
    # 全角转半角
    _FULLWIDTH_MAP = str.maketrans({
        '，': ',', '。': '.', '；': ';', '：': ':', '！': '!', '？': '?',
        '（': '(', '）': ')', '【': '[', '】': ']', '《': '<', '》': '>',
        '、': ',', '．': '.', '\u201c': '"', '\u201d': '"', '\u2018': "'",
        '\u3008': '<', '\u3009': '>', '\uff5b': '{', '\uff5d': '}', '\u3014': '[', '\u3015': ']',
        '～': '~', '—': '-', '\u2013': '-', '\u2015': '-',
        '\u3000': ' ',
        '\u2019': "'",
    })
    s = s.translate(_FULLWIDTH_MAP)
    s = re.sub(r"[^a-zA-Z0-9\u4e00-\u9fa5]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def build_anchor_keys(
    anchor_lines=None,
    anchor_text: str = "",
    anchor_ctx: str = "",
    prev_para_text: str = "",
    next_para_text: str = "",
) -> List[str]:
    """Build ordered unique normalized keys."""
    keys: List[str] = []
    for v in (prev_para_text, next_para_text):
        n = _anchor_norm_text(v)
        if n:
            keys.append(n)
    if isinstance(anchor_lines, list):
        for ln in anchor_lines:
            n = _anchor_norm_text(ln)
            if n:
                keys.append(n)
    for v in (anchor_text, anchor_ctx):
        n = _anchor_norm_text(v)
        if n:
            keys.append(n)

    seen = set()
    out: List[str] = []
    for k in keys:
        if k not in seen:
            seen.add(k)
            out.append(k)
    return out


def score_anchor_candidate(
    cand_text_norm: str,
    cand_ctx_norm: str,
    keys: List[str],
    hint_idx: Optional[int] = None,
    cand_idx: Optional[int] = None,
) -> float:
    """Unified candidate score."""
    if not cand_text_norm:
        return -1e9

    score = 0.0
    for k in keys:
        if not k:
            continue
        klen = len(k)
        w = 2.0 if klen >= 18 else (1.4 if klen >= 8 else 0.9)
        if k in cand_text_norm:
            score += 1.6 * w
        elif cand_ctx_norm and k in cand_ctx_norm:
            score += 1.0 * w
        elif klen >= 12:
            sim = difflib.SequenceMatcher(
                None,
                cand_text_norm[: min(len(cand_text_norm), 120)],
                k[:120],
            ).ratio()
            score += 0.30 * sim

    if hint_idx is not None and cand_idx is not None:
        score -= (abs(cand_idx - hint_idx) / 80.0) * 0.25
    return score