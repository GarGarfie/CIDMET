# CIDMET — 跨数据库识别与去重匹配导出工具

[![en](https://img.shields.io/badge/%F0%9F%87%AC%F0%9F%87%A7_English-Click-blue?style=for-the-badge)](README.md) [![zh](https://img.shields.io/badge/%F0%9F%87%A8%F0%9F%87%B3_%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87-%E7%82%B9%E5%87%BB-red?style=for-the-badge)](README_zh.md) [![ru](https://img.shields.io/badge/%F0%9F%87%B7%F0%9F%87%BA_%D0%A0%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9-%D0%9D%D0%B0%D0%B6%D0%BC%D0%B8%D1%82%D0%B5-green?style=for-the-badge)](README_ru.md)

---

## 简介

**CIDMET**（Cross-database Identification and De-duplication Matching Export Tool）是一款面向文献计量学研究的桌面应用程序。它以本地 BibTeX 参考文献库为输入，自动将其条目与三大学术数据库 — **Web of Science (WoS)**、**Scopus** 和 **Engineering Village (EI/Compendex)** — 的导出数据进行匹配，然后以各数据库的原生格式提取匹配子集，并提供合并去重导出为任一数据库风格功能。

这解决了文献计量分析中的一个常见痛点：当你需要特定数据库格式的导出文件用于 VOSviewer、CiteSpace 或 Bibliometrix 等工具时，这些软件或工具会要求单一数据库风格结果的文件，而整个检索结果集往往需要对多个数据库结果进行合并和去重。

## 功能特性

- **多数据库支持** — WoS（TXT / XLS）、Scopus（CSV / TXT，支持中英文）、EI（CSV / TXT）
- **三级匹配策略**
  - DOI 精确匹配（100% 置信度）
  - 标题精确匹配，含 Unicode 归一化处理（99% 置信度）
  - 模糊标题匹配，含作者和年份验证（可调阈值，默认 90%）
- **格式保留子集导出** — 输出文件可直接用于文献计量软件
- **合并导出与作者格式转换** — 自动将作者姓名格式转换为 WoS / Scopus / EI 的规范格式
- **自动去重** — 检测并允许用户审查被多个数据库记录匹配的条目

## 匹配策略

| 层级 | 方法 | 置信度 | 说明 |
|------|------|--------|------|
| 1 | DOI 精确匹配 | 100% | 归一化 DOI 比较（不区分大小写，去除前缀） |
| 2 | 标题精确匹配 | 99% | NFKD 归一化、不区分大小写、去除特殊字符 |
| 3 | 模糊标题匹配 | 可配置（默认 90%） | RapidFuzz 相似度 + 第一作者姓氏验证（≥80%）+ 年份验证 |

## 支持的格式

| 数据库 | 导入格式 | 子集输出 |
|--------|----------|----------|----------|
| Web of Science | TXT（标记格式）、XLS | TXT、XLS |
| Scopus | CSV、TXT（中/英文） | CSV、TXT |
| Engineering Village | CSV、TXT | CSV、TXT |

## 安装

**要求：** Python 3.9+

```bash
# 克隆仓库
git clone https://github.com/GarGarfie/CIDMET.git
cd CIDMET

# 安装依赖
pip install -r requirements.txt
```

### 依赖项

| 包 | 用途 |
|----|------|
| PySide6 ≥ 6.5 | GUI 框架（Qt for Python） |
| bibtexparser ≥ 1.4, < 2.0 | BibTeX 文件解析 |
| rapidfuzz ≥ 3.0 | 模糊字符串匹配 |
| chardet ≥ 5.0 | 字符编码检测 |
| xlrd ≥ 2.0 | 读取 Excel .xls 文件 |
| xlwt ≥ 1.3 | 写入 Excel .xls 文件 |
| openpyxl ≥ 3.1 | 写入 Excel .xlsx 文件 |

## 使用方法

```bash
python main.py
```

### 工作流程

1. **选择 BibTeX 文件** — 选择你的目标参考文献库（`.bib` 文件）
2. **添加数据库文件** — 拖放或浏览选择 WoS / Scopus / EI 的导出文件
3. **设置输出目录** — 选择子集和合并文件的保存位置
4. **调整模糊匹配阈值**（可选） — 滑块范围 50%–100%，默认 90%
5. **点击"运行匹配"** — 工具处理文件并执行三级匹配
6. **查看结果** — 在"结果"标签页中查看统计信息、匹配详情和未匹配条目
7. **导出合并文件**（可选） — 选择目标格式模板并导出合并记录

## 项目结构

```
CIDMET/
├── main.py              # 应用程序入口
├── gui_app.py           # PySide6 图形界面（主窗口、拖放、进度、标签页）
├── parsers.py           # 数据库格式解析器（WoS/Scopus/EI × TXT/CSV/XLS）
├── matcher.py           # 三级匹配引擎
├── writers.py           # 格式保留子集写入器 & 合并导出
├── utils.py             # 编码检测、DOI/标题归一化、辅助函数
├── draw_flowchart.py    # 数据流程图生成器
├── requirements.txt     # Python 依赖
└── fileTemplate/        # 各数据库导出文件示例
```

## 作者格式转换（合并导出）

合并来自不同数据库的匹配记录时，CIDMET 会自动将作者姓名转换为目标格式：

| 目标格式 | 缩写形式 | 全称形式 |
|----------|----------|----------|
| WoS | `Gu, S; Wu, YQ` | `Gu, Sheng; Wu, Yanqi` |
| Scopus | `Gu, S.; Wu, Y.Q.` | `Gu, Sheng; Wu, Yanqi` |
| EI | `Gu, Sheng (1); Wu, Yanqi (1)` | — |

## 许可证

本项目基于 [MIT 许可证](LICENSE) 发布。

## 引用

如果您在研究中使用了 CIDMET，请考虑引用本项目：

> Xiao, S (2026). *CIDMET: Cross-database Identification and De-duplication Matching Export Tool*. GitHub. https://github.com/GarGarfie/CIDMET

<details>
<summary>BibTeX</summary>

```bibtex
@software{cidmet2026,
  author    = {Xiao, Shuoting},
  title     = {CIDMET: Cross-database Identification and De-duplication Matching Export Tool},
  year      = {2026},
  url       = {https://github.com/GarGarfie/CIDMET}
}
```

</details>
