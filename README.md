# Multi-Department Doc Merger (MDDM) | 多部门文档自动合并系统

![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## 📖 项目简介

**MDDM** 是一款专门为中大型机构、多级部门协作设计的周报/报告自动合并工具。在多人并发修改同一份文档时，传统的合并方式往往会导致编号错乱、格式丢失或内容覆盖。本项目通过深度解析 OOXML（Office Open XML）底层结构，实现了比人工更精准、比普通合并更智能的无损合并。

### 核心痛点解决：
- **无感合并**: 自动识别各部门提交副本相对于原稿的增、删、改，并累积应用到合并结果中。
- **编号修复**: 自动修复因合并导致的 Word 列表编号（1.1, 1.1.1 等）重复或断档问题。
- **冲突追踪**: 以 Word 批注（Comments）和文本高亮（Revisions）形式标记不同作者的修改轨迹。
- **数学公式支持**: 完美处理 MathType 和 Office Math 公式转换。

---

## ✨ 核心特性

- **三路合并 (3-Way Merge)**: 基于 Base（原稿）识别多个 Copy（副本）的差异并进行冲突合并。
- **智能锚点定位**: 结合语义、格式和指纹，即使标题被修改也能正确定位修改位置。
- **层级推断**: `HeadingProcessor` 模块可自动识别文档结构，并补全缺失的标题层级。
- **公式桥接**: 支持 MathML 与 Office Math 的双向转换。
- **环境自愈**: 提供一键式环境检查与安装脚本，降低上手门槛。

---

## 🚀 快速开始

### 1. 环境准备
确保已安装 Python 3.8+。推荐使用脚本自动初始化：
```powershell
# Windows PowerShell
.\setup_env.ps1
```
或直接安装依赖：
```bash
pip install -r requirements.txt
```

### 2. 执行合并

**桌面界面**

```bash
python merge_gui.py
```

在窗口中选择 Word 原稿、副本目录和合稿保存位置，然后点击「开始合并」。界面会显示运行日志，完成后可直接打开合稿、冲突报告和变更日志。桌面界面使用随 Windows 版 Python 提供的 Tkinter。若 Linux 环境缺少 Tkinter，可安装发行版提供的 Python Tk 软件包（如 Debian/Ubuntu 的 `python3-tk`）。

**命令行**

将所有部门的副本放入一个文件夹中，然后运行：
```bash
python run_merge.py original_file.docx ./copies/ -o merged_output.docx
```

### 3. 参数说明
- `original_file.docx`: 基础版原稿。
- `./copies/`: 存放所有部门待合并副本的目录。
- `-o`: 输出文件名。
- `--authors`: (可选) 手动指定对应顺序的文件作者名。

---

## 📂 模块说明

| 模块名称 | 功能描述 |
| :--- | :--- |
| `run_merge.py` | 命令行入口与输出管理 |
| `multi_copy_merger.py` | 核心 XML 合并引擎 |
| `heading_processor.py` | 标题逻辑与文档结构分析 |
| `anchor_engine.py` | 内容对齐与定位锚点 |
| `numbering_resolver.py` | 动态编号转义处理器 |
| `mathtype_converter.py` | 数学公式转换适配器 |

---

## 🛠️ 项目结构
详情请参阅 [STRUCTURE.md](./STRUCTURE.md)。

## 📝 许可证
本项目采用 [MIT License](LICENSE) 许可。

---
**提示**: 上传至 GitHub 前，请确保已删除敏感数据和个人文档。
