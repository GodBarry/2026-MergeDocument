import sys
import os
import subprocess
import base64
from lxml import etree

MML_NS  = 'http://www.w3.org/1998/Math/MathML'
OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def is_mathtype_ole(ole_data: bytes) -> bool:
    """
    Check if the data is a valid OLE Compound Document.
    """
    return ole_data.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')


def convert_mathtype_to_mathml(ole_data: bytes) -> str:
    """
    Convert MathType OLE data (binary bytes) to MathML string using Ruby bridge.
    """
    bridge_script = os.path.join(os.path.dirname(__file__), 'resources', 'mathml_bridge.rb')
    if not os.path.exists(bridge_script):
        raise FileNotFoundError(f"Ruby bridge script not found at {bridge_script}")

    env = os.environ.copy()
    user_gem_dir = os.path.expanduser("~/.local/share/gem/ruby/3.2.0")
    if "GEM_PATH" in env:
        env["GEM_PATH"] = f"{user_gem_dir}:{env['GEM_PATH']}"
    else:
        env["GEM_PATH"] = user_gem_dir

    encoded_data = base64.b64encode(ole_data).decode('utf-8')

    try:
        process = subprocess.Popen(
            ['ruby', bridge_script],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env,
            text=True
        )
        stdout, stderr = process.communicate(input=encoded_data)

        if process.returncode != 0:
            raise RuntimeError(f"Ruby script failed with code {process.returncode}: {stderr}")

        return stdout.strip()

    except Exception as e:
        raise RuntimeError(f"Failed to convert MathType to MathML: {e}")


# ── MathML 预处理：修复 <mfenced> 包裹 <mtable> 的问题 ──────────────────

def _preprocess_mathml(dom):
    """
    预处理 MathML DOM，修复两类问题：

    1. 确保 <mfenced> 具有显式的 open/close 属性，
       以便 MML2OMML.XSL 能正确识别并生成可伸缩的 <m:d> delimiter。

    2. 修复被转置的列向量：
       MathType 对 \\begin{array}{c}X\\\\Y\\\\Z\\end{array} 有时会错误地输出
       单行多列的 <mtable>（即将列向量转置了）。
       检测方法：如果 <mfenced open="["> 或 <mfenced open="("> 内的 <mtable>
       只有一个 <mtr> 但有多个 <mtd>，则将其转置为多行单列结构。
    """
    for mfenced in dom.iter(f'{{{MML_NS}}}mfenced'):
        # 修复 1：确保 open/close 属性存在
        if 'open' not in mfenced.attrib:
            mfenced.set('open', '(')
        if 'close' not in mfenced.attrib:
            mfenced.set('close', ')')

        # 修复 2：将错误转置的列向量还原为列结构
        open_ch = mfenced.get('open', '(')
        if open_ch in ('[', '('):
            child_tags = [c.tag for c in mfenced]
            if len(child_tags) == 1 and child_tags[0] == f'{{{MML_NS}}}mtable':
                mtable = mfenced[0]
                mtrs = list(mtable)  # all mtr elements
                # 只有 1 个 mtr 且有多个 mtd → 就是被转置的列向量
                if len(mtrs) == 1:
                    mtr = mtrs[0]
                    mtds = list(mtr)
                    if len(mtds) > 1:
                        # 将单行多列转置为多行单列
                        # 删除原有的 mtr
                        mtable.remove(mtr)
                        for mtd in mtds:
                            new_mtr = etree.SubElement(mtable, f'{{{MML_NS}}}mtr')
                            new_mtr.append(mtd)  # 移动 mtd 到新行

    return dom


# ── OMML 后处理：修复退化的矩阵括号 ──────────────────────────────────────

def _get_run_text(el):
    """提取 <m:r> 元素的纯文本内容"""
    texts = []
    for t in el.iter(f'{{{OMML_NS}}}t'):
        if t.text:
            texts.append(t.text.strip())
    return ''.join(texts)


def _make_delimiter(open_ch, close_ch, matrix_el):
    """
    构造正确的 OMML <m:d> delimiter 结构：
      <m:d>
        <m:dPr>
          <m:begChr m:val="["/>
          <m:endChr m:val="]"/>
        </m:dPr>
        <m:e>
          <m:m>...</m:m>
        </m:e>
      </m:d>
    """
    M = OMML_NS
    d = etree.Element(f'{{{M}}}d')

    dPr = etree.SubElement(d, f'{{{M}}}dPr')
    beg = etree.SubElement(dPr, f'{{{M}}}begChr')
    beg.set(f'{{{M}}}val', open_ch)
    end_el = etree.SubElement(dPr, f'{{{M}}}endChr')
    end_el.set(f'{{{M}}}val', close_ch)

    e = etree.SubElement(d, f'{{{M}}}e')
    e.append(matrix_el)
    return d


def _postprocess_omml(omml_str: str) -> str:
    """
    后处理 OMML，检测并修复 MML2OMML.XSL 生成的退化矩阵括号。

    症状：本应是 <m:d>[<m:m>..</m:m>]</m:d> 的结构，被退化成了
    普通 <m:r>[</m:r> 紧跟数组元素 紧跟 <m:r>]</m:r>。

    MML2OMML.XSL 对不同类型数组的输出：
      - 多列矩阵 (nxm):   → <m:m>   (matrix)         ← 已修复
      - 单列数组 (nx1):   → <m:eqArr> (equation array)  ← 本次修复
      - 其他块结构:       → <m:d>, <m:f>, <m:rad> 等

    修复：检测「<m:r>[</m:r> + 块级元素 + <m:r>]</m:r>」三元组，
    重建为正确的可伸缩 <m:d> 结构。
    """
    BRACKET_PAIRS = {
        '[': ']',
        '(': ')',
        '{': '}',
        '\u27e8': '\u27e9',   # 〈 〉
        '\u2016': '\u2016',   # ‖ ‖
        '|': '|',
        '\u2308': '\u2309',   # ⌈ ⌉ ceiling
        '\u230a': '\u230b',   # ⌊ ⌋ floor
    }

    # 【精确匹配】：只将 m:m 和 m:eqArr 作为被括号包裹的目标。
    # - m:m     → 多列矩阵（nxm）
    # - m:eqArr → 单列向量 / \begin{array}{c}
    #
    # 其他 OMML 元素（m:sSup、m:f、m:rad、m:acc 等）是行内结构，
    # 把它们列入 BLOCK_TAGS 会导致误匹配：如 [expr^T] 中的 m:sSup
    # 被当成矩阵内容，把括号套到了错误的元素上，真正的矩阵内容丢失。
    M = OMML_NS
    BLOCK_TAGS = {
        f'{{{M}}}m',      # matrix（多列矩阵）
        f'{{{M}}}eqArr',  # equation array（单列列向量）
    }

    def normalize_bracket_text(el):
        """提取 <m:r> 的文本，规范化透明/不可见字符"""
        texts = []
        for t in el.iter(f'{{{M}}}t'):
            if t.text:
                # 去掉零宽度空格、非断行空格等不可见字符
                cleaned = ''.join(
                    c for c in t.text
                    if c not in '\u200b\u200c\u200d\ufeff\u00ad\u00a0'
                ).strip()
                texts.append(cleaned)
        return ''.join(texts)

    try:
        root = etree.fromstring(omml_str.encode('utf-8'))

        changed = True
        while changed:
            changed = False
            for parent in root.iter():
                children = list(parent)
                i = 0
                while i < len(children):
                    child = children[i]
                    if child.tag != f'{{{M}}}r':
                        i += 1
                        continue

                    ch_text = normalize_bracket_text(child)
                    if ch_text not in BRACKET_PAIRS:
                        i += 1
                        continue

                    open_ch = ch_text
                    expected_close = BRACKET_PAIRS[open_ch]

                    # 向后找: 第一个 BLOCK_TAG 元素（m:m 或 m:eqArr）
                    # 它标识了"矩阵括号"开始的信号
                    # 策略：找到 block_el 后，继续向后搜索匹配的 ] 括号，
                    # 把 block_el 和 ] 之间的所有内容打包进 <m:e> 里，
                    # 这样即使 ] 不紧跟 block_el 也能处理。
                    found = False

                    # ── 步骤1: 找 block_el（紧跟开括号的首个 m:m/m:eqArr）──
                    block_j = -1
                    for j in range(i + 1, min(i + 4, len(children))):
                        if children[j].tag in BLOCK_TAGS:
                            block_j = j
                            break
                        elif children[j].tag == f'{{{M}}}r':
                            if normalize_bracket_text(children[j]) != '':
                                break  # 遇到非空 run，停止向前找 block_el

                    if block_j == -1:
                        i += 1
                        continue

                    # ── 步骤2: 从 block_j+1 向后找 close bracket ────────────
                    # 允许中间有任意数量的行内元素（sSubSup、func 等）
                    close_j = -1
                    for k in range(block_j + 1, min(block_j + 20, len(children))):
                        if children[k].tag == f'{{{M}}}r':
                            kt = normalize_bracket_text(children[k])
                            if kt == expected_close:
                                close_j = k
                                break
                            elif kt in BRACKET_PAIRS:
                                # 遇到另一个开括号，停止（避免跨公式匹配）
                                break
                            elif kt != '':
                                # 遇到其他文字内容，继续向前（/、数字等）
                                pass

                    if close_j == -1:
                        i += 1
                        continue

                    # ── 步骤3: 收集 block_el ~ close_bracket 之间所有元素 ───
                    # 这些元素将作为 <m:d> 的 <m:e> 内容
                    content_els = children[block_j: close_j]  # 含 block_el，不含 close
                    close_el = children[close_j]

                    # ── 步骤4: 构建 <m:d> 并替换 ────────────────────────────
                    insert_pos = list(parent).index(child)
                    # 从父元素移除：开括号、所有内容元素、关括号
                    for el_to_remove in [child] + list(content_els) + [close_el]:
                        try:
                            parent.remove(el_to_remove)
                        except ValueError:
                            pass

                    # 新建 <m:d>，把所有内容元素放入单个 <m:e>
                    d_el = etree.Element(f'{{{M}}}d')
                    dPr = etree.SubElement(d_el, f'{{{M}}}dPr')
                    beg = etree.SubElement(dPr, f'{{{M}}}begChr')
                    beg.set(f'{{{M}}}val', open_ch)
                    end_el_node = etree.SubElement(dPr, f'{{{M}}}endChr')
                    end_el_node.set(f'{{{M}}}val', expected_close)
                    e_node = etree.SubElement(d_el, f'{{{M}}}e')
                    for content_el in content_els:
                        e_node.append(content_el)

                    parent.insert(insert_pos, d_el)
                    children = list(parent)
                    changed = True
                    found = True

                    i += 1  # 继续扫描当前 parent 的后续元素



        return etree.tostring(root, pretty_print=True, encoding='unicode')

    except Exception:
        # 后处理失败时安全回退
        return omml_str



# ── 主转换函数 ─────────────────────────────────────────────────────────────

def convert_mathml_to_omml(mathml_content: str, xsl_path: str = None) -> str:
    """
    Convert MathML string to OMML (Office Math ML) using XSLT.
    包含矩阵括号三段修复：
      1. MathML 预处理（确保 <mfenced> 属性完整）
      2. MML2OMML.XSL 转换
      3. OMML 后处理（修复退化的矩阵括号为 <m:d>）
    """
    if etree is None:
        raise ImportError("lxml library is required for MathML to OMML conversion. Please pip install lxml")

    if not xsl_path:
        xsl_path = os.path.join(os.path.dirname(__file__), 'resources', 'MML2OMML.XSL')

    if not os.path.exists(xsl_path):
        raise FileNotFoundError(f"XSL stylesheet not found at {xsl_path}")

    try:
        if isinstance(mathml_content, str):
            mathml_bytes = mathml_content.encode('utf-8')
        else:
            mathml_bytes = mathml_content

        # 步骤 1: 解析并预处理 MathML
        dom = etree.fromstring(mathml_bytes)
        dom = _preprocess_mathml(dom)

        # 步骤 2: XSL 转换 MathML → OMML
        xslt_doc = etree.parse(xsl_path)
        transform = etree.XSLT(xslt_doc)
        new_dom = transform(dom)
        omml_string = etree.tostring(new_dom, pretty_print=True, encoding='unicode')

        # 步骤 3: 后处理修复退化矩阵括号
        omml_string = _postprocess_omml(omml_string)

        return omml_string

    except Exception as e:
        raise RuntimeError(f"XSLT Conversion failed: {e}")


def convert_mathtype_to_omml(ole_data: bytes, xsl_path: str = None) -> str:
    """
    Full pipeline: MathType Binary -> MathML -> OMML
    """
    mathml = convert_mathtype_to_mathml(ole_data)
    if not mathml:
        return None
    omml = convert_mathml_to_omml(mathml, xsl_path)
    return omml



if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="将 MathType OLE 二进制文件转换为 OMML（Office 数学公式）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  python mathtype_converter.py --input formula.bin
  python mathtype_converter.py --input formula.bin --show-mathml
  python mathtype_converter.py --input formula.bin --xsl /path/to/MML2OMML.XSL
"""
    )

    parser.add_argument(
        "--input", "-i",
        required=True,
        metavar="FILE",
        help="MathType OLE 二进制文件路径（.bin 或直接从 docx 提取的 oleObject*.bin）"
    )
    parser.add_argument(
        "--xsl",
        default=None,
        metavar="FILE",
        help="MML2OMML.XSL 路径（默认：与本脚本同目录下的 MML2OMML.XSL）"
    )
    parser.add_argument(
        "--show-mathml",
        action="store_true",
        default=False,
        help="同时输出中间的 MathML（用于调试）"
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        metavar="FILE",
        help="将 OMML 结果写入指定文件（默认：打印到终端）"
    )

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"错误：文件不存在: {args.input}")
        sys.exit(1)

    with open(args.input, 'rb') as f:
        data = f.read()

    try:
        if args.show_mathml:
            print("=== Step 1: MathML ===")
            mathml = convert_mathtype_to_mathml(data)
            print(mathml)
            print("\n=== Step 2: OMML (with bracket fix) ===")
            omml = convert_mathml_to_omml(mathml, args.xsl)
        else:
            omml = convert_mathtype_to_omml(data, args.xsl)

        if args.output:
            with open(args.output, 'w', encoding='utf-8') as out:
                out.write(omml)
            print(f"OMML 已写入: {args.output}")
        else:
            print(omml)

    except Exception as e:
        print(f"转换失败: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)
