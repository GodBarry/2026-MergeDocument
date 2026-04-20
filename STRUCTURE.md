# 工程文件结构 (Project Structure)

本工程是一个多部门周报/文档自动合并系统，旨在解决多人协作环境下文档内容冲突、格式错乱、编号失效等痛点。以下是清理后的核心文件结构分解：

## 核心业务逻辑 (Core Business Logic)

*   **`run_merge.py`**:
    *   项目的总入口脚本。
    *   提供命令行接口（CLI），支持配置原稿路径、副本目录、输出文件名以及作者信息。
    *   处理最终文档的批注注入（修改者信息、校对建议）。
*   **`multi_copy_merger.py`**:
    *   核心合并引擎。
    *   实现“累积式三路合并”算法（Base vs. Multiple Copies）。
    *   包含冲突检测机制、样式保持逻辑以及基于 XML 的 OOXML 直接操作引擎。
*   **`anchor_engine.py`**:
    *   锚点定位引擎。
    *   通过文本指纹和上下文特征，在不同版本的文档中精准定位段落，确保增删改的操作位置准确。
*   **`heading_processor.py`**:
    *   智能标题处理器。
    *   负责标题识别、层级推断、编号修正以及缺失层级补全。
*   **`numbering_resolver.py`**:
    *   Word 自动编号解析器。
    *   解析 `word/numbering.xml`，将抽象编号转换为实际显示的文本。
*   **`logging_config.py`**:
    *   统一的日志配置模块。

## 增强功能模块 (Enhanced Features)

*   **`hybrid_anchor.py`**:
    *   混合锚点策略，结合了书签、标题和内容指纹的多层定位方式。
*   **`semantic_match.py`**:
    *   语义匹配辅助模块，提升在文本大幅改动时的对齐准确率。
*   **`figure_numbering.py`**:
    *   图表编号自动修复与重新整理。
*   **公式转换工具组**:
    *   `mathtype_converter.py`: MathType 公式转换逻辑。
    *   `convert_docx_formulas.py`: Docx 内嵌公式处理。
    *   `run_convert_formulas.py`: 批量执行公式转换。
    *   `mathml_bridge.rb`: Ruby 编写的 MathML 转换桥接逻辑。
    *   `MML2OMML.XSL` / `omml2mml.xsl`: 微软提供的 MathML 与 Office Math 互转的 XSLT 样式表。

## 配置与环境 (Config & Environment)

*   **`requirements.txt`**: 项目运行必须的 Python 依赖库清单。
*   **`setup_env.py` / `setup_env.ps1`**: 跨平台的开发环境初始化脚本（自动检查依赖、安装库、配置路径）。
*   **`使用说明.md`**: 原始的项目操作手册。

## 其他系统文件

*   **`.vscode/`**: 编辑器配置文件（断点调试、显示设置）。
*   **`__pycache__/`**: Python 编译缓存空间。
