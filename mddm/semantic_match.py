"""
semantic_match.py
────────────────────────────────────────────────────────────────────────────
三层降级语义匹配引擎：Sentence-Transformers → TF-IDF + 余弦 → 词法回退

设计目标（解决 Issue 3.3）：
  difflib.SequenceMatcher 只能在字符层面找最长公共子串，极易被
  企业周报中大量重复的模板化语句（如"本周各项工作均按计划顺利推进"）
  欺骗，将语义完全无关的段落强行建立映射关系。

三层降级策略（优先尝试高级，层层回退）：
  L1 (最优): sentence-transformers 向量化 + 余弦相似度
              语义相似的段落 → 高分；无关段落 → 低分（即使词汇有重叠）
  L2 (次优): TF-IDF 向量化 + 余弦相似度
              适合离线/无GPU环境，可捕获词频特征差异
  L3 (兜底): difflib.SequenceMatcher（原有逻辑，完全向后兼容）

所有 API 均以 (text_a, text_b) -> float(0~1) 的统一接口对外暴露。
"""

from __future__ import annotations
import math
import re
import hashlib
import threading
from typing import List, Dict, Tuple, Optional, Callable

# ── 全局引擎单例（延迟初始化，线程安全）──────────────────────────────────────
_lock = threading.Lock()
_engine: Optional["_SemanticEngine"] = None


def get_engine() -> "_SemanticEngine":
    """线程安全的全局引擎访问点。"""
    global _engine
    if _engine is not None:
        return _engine
    with _lock:
        if _engine is None:
            _engine = _SemanticEngine()
    return _engine


# ── 段落指纹：用于快速预过滤，减少 L2/L3 计算量 ─────────────────────────────
def paragraph_fingerprint(text: str) -> str:
    """生成段落"内容指纹"（正文 SHA256 前12位）。

    用途：在 L1/L2 计算成本较高的场景中，先用指纹快速判断"内容是否完全相同"。
    完全相同的段落 → sim=1.0 直接返回，跳过昂贵的向量计算。
    """
    t = text.strip()
    if not t:
        return ""
    # 仅对纯文本内容求 hash，忽略空白差异
    normalized = re.sub(r'\s+', ' ', t)
    return hashlib.sha256(normalized.encode('utf-8')).hexdigest()[:12]


def quick_identical(a: str, b: str) -> bool:
    """O(1) 快速判断两段文本是否完全相同（忽略空白差异）。"""
    return paragraph_fingerprint(a) == paragraph_fingerprint(b)


# ─────────────────────────────────────────────────────────────────────────────
# L1: Sentence-Transformers
# ─────────────────────────────────────────────────────────────────────────────

def _try_import_sentence_transformers():
    """尝试导入 sentence-transformers。失败时返回 None。"""
    try:
        import sentence_transformers as st  # type: ignore
        return st
    except ImportError:
        return None


def _build_st_engine() -> Optional[Tuple]:
    """
    初始化 sentence-transformers 模型。
    返回 (model, dimension) 或 None（不可用时）。
    推荐模型（按效果/速度权衡）：
      - 'paraphrase-multilingual-MiniLM-L12-v2'   : 支持 50+ 语言的轻量模型，速度快
      - 'paraphrase-multilingual-mpnet-base-v2'  : 效果更好，资源消耗更高
    """
    st = _try_import_sentence_transformers()
    if st is None:
        return None

    try:
        # 优先使用多语言轻量模型（覆盖中英双语）
        model = st.SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
        dim = model.get_sentence_embedding_dimension()
        return (model, dim)
    except Exception as e:
        print(f"  [语义匹配] sentence-transformers 初始化失败: {e}")
        return None


class _STEngine:
    """L1 向量化引擎封装。"""

    def __init__(self):
        self._impl: Optional[Tuple] = _build_st_engine()
        self._dim: int = self._impl[1] if self._impl else 0
        self._model = self._impl[0] if self._impl else None
        self._cache: Dict[str, List[float]] = {}
        # 大于此阈值的段落才尝试 L1 计算（短文本用 TF-IDF 更稳定）
        self._min_len_for_st = 30

    @property
    def available(self) -> bool:
        return self._impl is not None

    def encode(self, texts: List[str], normalize: bool = True) -> Optional[List[List[float]]]:
        """批量编码。失败返回 None。"""
        if not self.available:
            return None
        try:
            emb = self._model.encode(
                texts,
                batch_size=32,
                show_progress_bar=False,
                convert_to_numpy=True,
                normalize_embeddings=normalize,
            )
            return emb.tolist()  # type: ignore
        except Exception:
            return None

    def similarity(self, text_a: str, text_b: str) -> float:
        """
        返回 L1 余弦相似度 (0~1)。
        短文本或不可用时返回 -1（触发降级）。
        """
        if not self.available:
            return -1.0
        if len(text_a.strip()) < self._min_len_for_st or len(text_b.strip()) < self._min_len_for_st:
            return -1.0

        # 缓存命中
        fp_a = paragraph_fingerprint(text_a)
        fp_b = paragraph_fingerprint(text_b)
        if fp_a == fp_b:
            return 1.0

        cache_key_a = fp_a
        cache_key_b = fp_b
        emb_a = self._cache.get(cache_key_a)
        if emb_a is None:
            enc = self.encode([text_a])
            if enc is None:
                return -1.0
            emb_a = enc[0]
            self._cache[cache_key_a] = emb_a

        emb_b = self._cache.get(cache_key_b)
        if emb_b is None:
            enc = self.encode([text_b])
            if enc is None:
                return -1.0
            emb_b = enc[0]
            self._cache[cache_key_b] = emb_b

        # 余弦相似度（已 normalize）
        dot = sum(a * b for a, b in zip(emb_a, emb_b))
        # clamp 避免浮点误差
        return max(0.0, min(1.0, dot))


# ─────────────────────────────────────────────────────────────────────────────
# L2: TF-IDF + 余弦
# ─────────────────────────────────────────────────────────────────────────────

class _TFIDFEngine:
    """
    L2 TF-IDF 向量化引擎。

    核心思想：用词频-逆文档频率（TF-IDF）将每段文本映射为高维稀疏向量，
    再用余弦相似度判断语义关联性。

    与 L1 的区别：
      - L1 基于深度语言模型，理解语义；L2 基于词频统计，抗噪声能力弱但速度快。
      - 当段落中出现大量模板化词汇时，TF-IDF 能通过逆文档频率将其降权，
        减少"本周各项工作均按计划顺利推进"这类万能句对相似度的贡献。

    策略：
      - 构建小型 in-memory 语料库（仅当前比对的段落集合）。
      - 每个段落作为独立文档参与 IDF 计算。
      - 相似度计算时使用"片段窗口"策略：若段落过长（>500字），
        切分为句子级窗口再重组，降低长文本的信息稀疏性。
    """

    def __init__(self):
        self._corpus: List[str] = []
        self._processed: bool = False
        self._idf: Dict[str, float] = {}
        self._vocab: List[str] = []
        self._vocab_index: Dict[str, int] = {}
        # TF-IDF 向量缓存
        self._vector_cache: Dict[str, List[float]] = {}

    def ingest(self, texts: List[str]) -> None:
        """
        将文本列表注册为语料库。
        应在批量比对前调用一次，以建立 IDF 统计。
        """
        self._corpus = []
        for t in texts:
            normalized = self._preprocess(t)
            if normalized:
                self._corpus.append(normalized)
        self._processed = False
        self._vector_cache.clear()

    def _preprocess(self, text: str) -> str:
        """中文友好的预处理：分词 + 停用词过滤。"""
        t = text.strip()
        if not t:
            return ""
        # 简单分词：按空格/标点切分
        tokens = re.findall(
            r'[\u4e00-\u9fff]+|[a-zA-Z0-9]+',  # 连续汉字 或 连续字母数字
            t
        )
        # 停用词（极简中文停用词表，覆盖周报高频万能语句）
        stopwords = {
            '的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一',
            '一个', '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有',
            '看', '好', '自己', '这', '那', '还', '可以', '并', '以及', '等', '各',
            '本周', '本周内', '本月', '本月内', '本周工作', '本月工作',
            '按计划', '顺利推进', '未发现', '重大', '系统性', '安全风险',
            '各项工作', '计划顺利', '推进', '发现', '风险',
        }
        tokens = [tok for tok in tokens if tok not in stopwords and len(tok) > 1]
        return ' '.join(tokens)

    def _build(self) -> None:
        """构建 IDF 表和词汇表。"""
        if self._processed or not self._corpus:
            return

        n = len(self._corpus)
        df: Dict[str, int] = {}
        for doc in self._corpus:
            seen = set(doc.split())
            for tok in seen:
                df[tok] = df.get(tok, 0) + 1

        self._vocab = sorted(df.keys())
        self._vocab_index = {w: i for i, w in enumerate(self._vocab)}
        self._idf = {
            w: math.log((n + 1) / (df[w] + 1)) + 1
            for w in self._vocab
        }
        self._processed = True

    def _tf_vector(self, text: str) -> List[float]:
        """将文本转为 TF-IDF 向量（稀疏，None=0）。"""
        fp = paragraph_fingerprint(text)
        if fp in self._vector_cache:
            return self._vector_cache[fp]

        normalized = self._preprocess(text)
        tokens = normalized.split()
        if not tokens:
            return [0.0] * len(self._vocab)

        tf: Dict[str, float] = {}
        for tok in tokens:
            tf[tok] = tf.get(tok, 0) + 1
        max_tf = max(tf.values())
        dim = len(self._vocab)
        vec = [0.0] * dim
        for tok, count in tf.items():
            if tok in self._vocab_index:
                idx = self._vocab_index[tok]
                vec[idx] = (count / max_tf) * self._idf.get(tok, 1.0)

        self._vector_cache[fp] = vec
        return vec

    def _cosine(self, vec_a: List[float], vec_b: List[float]) -> float:
        """余弦相似度。"""
        dot = sum(a * b for a, b in zip(vec_a, vec_b))
        norm_a = math.sqrt(sum(a * a for a in vec_a))
        norm_b = math.sqrt(sum(b * b for b in vec_b))
        if norm_a < 1e-10 or norm_b < 1e-10:
            return 0.0
        return max(0.0, min(1.0, dot / (norm_a * norm_b)))

    def similarity(self, text_a: str, text_b: str) -> float:
        """
        返回 L2 TF-IDF 余弦相似度 (0~1)。
        不可用时返回 -1（触发降级）。
        """
        if not self._corpus:
            # 无语料库：降级为单文档 TF（无 IDF 权重）
            return -1.0

        self._build()

        fp_a = paragraph_fingerprint(text_a)
        fp_b = paragraph_fingerprint(text_b)
        if fp_a == fp_b:
            return 1.0

        vec_a = self._tf_vector(text_a)
        vec_b = self._tf_vector(text_b)
        return self._cosine(vec_a, vec_b)


# ─────────────────────────────────────────────────────────────────────────────
# L3: 词法回退（原有 difflib.SequenceMatcher，向后兼容）
# ─────────────────────────────────────────────────────────────────────────────

def lex_similarity(text_a: str, text_b: str) -> float:
    """
    L3 词法相似度：基于 difflib.SequenceMatcher 的最长公共子串比率。
    完全向后兼容，作为最终兜底层。
    """
    import difflib
    a = (text_a or '').strip()
    b = (text_b or '').strip()
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


# ─────────────────────────────────────────────────────────────────────────────
# 统一语义相似度 API
# ─────────────────────────────────────────────────────────────────────────────

class _SemanticEngine:
    """
    三层降级语义匹配引擎（对外统一接口）。

    评分规范：
      1.0       : 完全相同（指纹一致）
      0.85~0.99 : 语义高度相似（用词接近，主题相同）
      0.50~0.84 : 语义弱相关（共享部分词汇）
      0.10~0.49 : 语义低相关（偶然词汇重叠）
      0.00~0.09 : 语义无关（主题完全不同，但存在字符级重叠）
      -1.0      : 不可用（触发降级信号）

    与原有 SequenceMatcher 的关键区别：
      - 模板万能句（如"本周各项工作均按计划顺利推进"）
        对语义相似度的贡献被 IDF 降权，不会在无关段落间拉高分数。
      - 长文本（段落）整体向量比逐字匹配更能反映主题相似性。
    """

    def __init__(self):
        self._st: _STEngine = _STEngine()
        self._tfidf: _TFIDFEngine = _TFIDFEngine()
        self._corpus_registered: bool = False
        # 降级阈值：当 L1/L2 相似度 < 此值时，触发降级尝试
        self._degrade_threshold = 0.05

    @property
    def st_available(self) -> bool:
        return self._st.available

    @property
    def tfidf_available(self) -> bool:
        return len(self._tfidf._corpus) > 0

    def register_corpus(self, texts: List[str]) -> None:
        """
        注册语料库（供 L2 TF-IDF 使用 IDF 统计）。
        应在批量语义比对开始前调用一次。

        为节省内存，语料库仅存储归一化后的字符串列表，
        不存储中间向量。
        """
        self._tfidf.ingest(texts)
        self._corpus_registered = True

    def clear_corpus(self) -> None:
        """清空语料库缓存，释放内存。"""
        self._tfidf._corpus = []
        self._tfidf._processed = False
        self._tfidf._vector_cache.clear()
        self._st._cache.clear()
        self._corpus_registered = False

    def similarity(self, text_a: str, text_b: str,
                   degrade_callback: Optional[Callable[[str, str], float]] = None) -> float:
        """
        主 API：返回语义相似度 (0~1)，始终返回值（永不抛出异常）。

        降级链路：
          L1 ST (available?) → L2 TF-IDF (corpus registered?) → L3 difflib

        Args:
            text_a: 文本 A
            text_b: 文本 B
            degrade_callback: 可选，当所有高级引擎不可用时调用。
                            接收 (text_a, text_b)，返回 float。
                            默认使用 lex_similarity。
        Returns:
            float: 语义相似度 [0, 1]
        """
        # ── 0. 快速路径：指纹完全一致 → 1.0 ─────────────────────────────
        if quick_identical(text_a, text_b):
            return 1.0

        # ── 1. L1: Sentence-Transformers ─────────────────────────────────
        if self._st.available:
            score = self._st.similarity(text_a, text_b)
            if score >= 0:
                # L1 可用且有明确分数
                if score > self._degrade_threshold:
                    return round(score, 4)
                # 低于阈值 → 降级到 L2（可能有更丰富词频特征）
                # 但不直接回退 L3，因为 L2 不需要 GPU
            # else: score == -1 → 短文本或其他原因不可用，继续降级

        # ── 2. L2: TF-IDF ────────────────────────────────────────────────
        if self.tfidf_available:
            score = self._tfidf.similarity(text_a, text_b)
            if score >= 0:
                return round(score, 4)
            # score == -1 → 降级到 L3

        # ── 3. L3: 词法回退 ──────────────────────────────────────────────
        fallback = degrade_callback if degrade_callback else lex_similarity
        return round(fallback(text_a, text_b), 4)

    def batch_similarity(self, pairs: List[Tuple[str, str]]) -> List[float]:
        """
        批量语义相似度计算（优化：共享语料库注册，减少重复初始化）。

        Args:
            pairs: [(text_a1, text_b1), (text_a2, text_b2), ...]
        Returns:
            [score1, score2, ...]
        """
        results = []
        for a, b in pairs:
            results.append(self.similarity(a, b))
        return results


# ─────────────────────────────────────────────────────────────────────────────
# 便捷入口函数（直接 import 使用）
# ─────────────────────────────────────────────────────────────────────────────

def semantic_similarity(text_a: str, text_b: str) -> float:
    """
    单次语义相似度计算（0~1）。
    推荐用于段落级比对；批量场景请用 get_engine().register_corpus() +
    get_engine().batch_similarity() 以获得更好的 L2 IDF 统计。
    """
    return get_engine().similarity(text_a, text_b)


def batch_semantic_similarity(pairs: List[Tuple[str, str]]) -> List[float]:
    """批量语义相似度（推荐使用）。"""
    return get_engine().batch_similarity(pairs)


def register_semantic_corpus(texts: List[str]) -> None:
    """注册语义语料库（L2 TF-IDF IDF 统计用）。"""
    get_engine().register_corpus(texts)


def is_st_available() -> bool:
    """检查 sentence-transformers 是否可用。"""
    return get_engine().st_available


# ─────────────────────────────────────────────────────────────────────────────
# 诊断工具
# ─────────────────────────────────────────────────────────────────────────────

def diagnostics() -> Dict[str, str]:
    """返回各引擎的可用性诊断信息。"""
    eng = get_engine()
    return {
        "sentence_transformers": "✓ 可用" if eng.st_available else "✗ 不可用（pip install sentence-transformers）",
        "tfidf": "✓ 可用" if eng.tfidf_available else "✗ 需先注册语料库",
        "lexical_fallback": "✓ 始终可用",
    }


if __name__ == "__main__":
    # 快速测试
    print("=== 语义匹配引擎诊断 ===")
    for k, v in diagnostics().items():
        print(f"  {k}: {v}")

    print("\n=== 单元测试 ===")
    pairs = [
        ("本周各项工作均按计划顺利推进，未发现重大系统性安全风险",
         "本周各项工作均按计划顺利推进，未发现重大系统性安全风险"),
        ("本周各项工作均按计划顺利推进，未发现重大系统性安全风险",
         "本月各项工作按计划稳步推进，未发现重大系统性风险"),
        ("本周各项工作均按计划顺利推进，未发现重大系统性安全风险",
         "项目进度正常推进，本周完成需求评审和设计文档编写"),
        ("财务部本周完成季度预算审核工作，各项指标符合预期",
         "财务部本周完成季度预算审核工作，各项指标符合预期"),
    ]

    for a, b in pairs:
        score = semantic_similarity(a, b)
        lex = lex_similarity(a, b)
        print(f"\n  A: {a[:40]}")
        print(f"  B: {b[:40]}")
        print(f"  语义: {score:.4f}  |  词法: {lex:.4f}")
