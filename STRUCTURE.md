# 项目结构

```text
mddm/
  __init__.py                 Python 包
  merge_gui.py                桌面界面
  run_merge.py                命令行入口与任务输出
  multi_copy_merger.py        DOCX 比对、冲突检测与合稿写回
  word_comments.py            人工批注、选区及回复关系
  output_batch.py             合稿及报告的批次写入和恢复
  anchor_engine.py            文本锚点辅助
  numbering_resolver.py       Word 自动编号解析
  heading_processor.py        标题识别与编号处理
  process_headings.py         标题处理命令行工具
  mathtype_converter.py       公式转换
  convert_docx_formulas.py     DOCX 公式转换实现
  run_convert_formulas.py      公式转换命令行工具
  hybrid_anchor.py            混合锚点模块
  semantic_match.py           文本相似度模块
  figure_numbering.py         图片编号模块
  logging_config.py           日志配置
  resources/                  公式转换样式表与 Ruby 桥接脚本
tests/                        自动化回归测试
docs/                         部署说明和软件设计说明书
packaging/windows/            Windows 分发启动脚本和说明
releases/                     已验收的工程部署包
pyproject.toml                包配置、依赖和安装入口
requirements.txt              核心依赖
setup_env.py / setup_env.ps1  环境安装辅助脚本
```

根目录的同名 Python 脚本保留为兼容入口，实际实现位于 `mddm/`。Python 包可安装为 wheel。标题和公式转换是独立工具，实际依赖与启动方式见 [部署与使用](docs/部署与使用.md)。

`hybrid_anchor.py` 和 `semantic_match.py` 为辅助模块，不能据此推断默认合并流程启用了它们。当前源码缺失 `ai_heading_analyzer.py`，文本校对和图片编号相关选项不可用。
