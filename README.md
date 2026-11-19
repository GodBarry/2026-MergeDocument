# 多部门文档类自动合并软件系统

MDDM 用于合并多个部门基于同一份 Word 原稿编辑的 DOCX 副本。系统保留部门副本中的人工批注，生成合稿、冲突报告和变更日志，提供桌面界面和命令行入口。

## 快速开始

从 [工程部署包](https://github.com/GodBarry/2026-MergeDocument/tree/codex/word-merge-release-fixes/releases) 下载 ZIP 并完整解压。进入“Windows离线运行包”文件夹，先安装 Python 3.12 x64，启用 pip 和 Tcl/Tk，再双击“启动软件.cmd”。首次启动从包内 wheelhouse 安装程序和依赖。

从源码部署时，可在项目目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install .
.\.venv\Scripts\pythonw.exe -m mddm.merge_gui
```

在界面中选择原稿、副本目录和合稿保存位置，点击“开始合并”。合稿与两个文本报告保存在所选位置。

## 命令行

```powershell
.\.venv\Scripts\python.exe -m mddm.run_merge "原稿.docx" "副本目录" -o "输出\合稿.docx"
```

原稿是各部门编辑前的共同基准。副本目录放置各部门的 DOCX 文件。输出可以指定 DOCX 文件或目录。作者默认从文档属性识别，也可用 `--authors` 按副本文件名排序顺序指定。完整参数见：

```powershell
.\.venv\Scripts\python.exe -m mddm.run_merge --help
```

为避免将旧合稿再次作为副本扫描，合稿宜保存在副本目录外。汇总人员需结合冲突报告、作者批注和原稿复核结果。

## 已知限制

- 文本校对与图片编号选项依赖仓库缺失的 `ai_heading_analyzer` 模块，当前不可用。
- `--diff-mode direct` 当前未切换到底层直接比对实现。
- 公式转换工具需要额外环境。MathType 二进制转换依赖 Ruby 和 `mathtype_to_mathml_plus`，OCR 依赖需另行安装。

## 文档和验证

部署、离线安装与附加工具见 [部署与使用](docs/部署与使用.md)，代码布局见 [项目结构](STRUCTURE.md)。[软件设计说明书](docs/多部门文档类自动合并软件系统设计说明书V1.0.docx) 随源码和工程部署包一同提供。

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```
