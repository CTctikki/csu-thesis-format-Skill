---
name: csu-thesis-format
description: "辅助中南大学本科毕业设计(论文) Word 文档排版、审查和修复。适用于中国学生处理中南大学本科毕业论文 DOCX/DOC 文件、套用学校论文模板、检查页面设置、封面、中英文摘要、目录、正文标题、图表、公式、页码、引用和参考文献格式。"
---

# 中南大学论文排版

## 快速开始

这是一个面向中南大学本科毕业设计(论文)排版的 skill。开始处理文档前，先阅读 `references/csu-thesis-format-rules.md`，那里是从学校 Word 模板中提炼出的完整排版规则。

检查一篇论文 DOCX：

```bash
python scripts/check_csu_thesis_docx.py path/to/thesis.docx
```

编辑前提取文档样式和版式信息：

```bash
python scripts/extract_docx_format.py path/to/thesis.docx
```

需要新建论文、复制样式或比对版式时，优先使用 `assets/csu-thesis-template.docx` 作为标准模板资产。仓库同时保留了学校原始 Word 模板 `assets/附件6：中南大学毕业设计(论文)模版.doc`，用于溯源或与学校下发文件逐项比对。

## 工作流程

1. 确认输入文档类型。
   - 如果用户提供 `.docx`，直接分析。
   - 如果用户提供旧版 `.doc`，先用 LibreOffice 转成 `.docx`：

```bash
soffice --headless --convert-to docx --outdir output_dir input.doc
```

2. 排版前先读取规则。
   - 用 `references/csu-thesis-format-rules.md` 判断字体、字号、行距、页边距、页码、标题层级和图表公式格式。
   - 规则有歧义时，以 `assets/csu-thesis-template.docx` 的实际版式为准。

3. 先分析，再修改。
   - 运行 `extract_docx_format.py` 查看节、样式、表格和段落样例。
   - 运行 `check_csu_thesis_docx.py` 获取可能不符合模板的项目清单。

4. 保守编辑。
   - 优先保留用户已有论文内容和章节结构。
   - 保留图片、表格、公式、题注、交叉引用和 Word 域。
   - 常规段落样式和页面设置优先用 `python-docx`；`python-docx` 无法表达的设置再直接改 OOXML。
   - 修格式时不要删除用户内容。确实需要移动内容时，保持可追溯。

5. 版式重要时做视觉验证。
   - 用 LibreOffice 把修改后的 DOCX 导出为 PDF。
   - 重点检查封面、中英文摘要、目录、每章首页、图表密集页、公式页和参考文献页。
   - 每轮明显修改后重新运行检查脚本。

## 修改优先级

按这个顺序处理：

1. 页面设置和分节：A4、页边距、页眉页脚距离、封面无页眉页脚、正文前罗马页码、正文阿拉伯页码。
2. 前置部分：封面、中文摘要、英文摘要、目录。
3. 正文层级：章首页、各级标题、正文段落、文献引用。
4. 对象格式：公式、表格、图片、题注。
5. 后置部分：结束语、致谢、参考文献。

## 脚本说明

脚本依赖 Python、`python-docx` 和 `lxml`。它们输出 Markdown 格式报告。检查脚本是启发式工具，因为 Word 文档常混用样式、直接格式、文本框和域代码；脚本结果用于定位疑点，最终仍要结合模板和 PDF 渲染效果人工确认。

## 资源清单

- `references/csu-thesis-format-rules.md`：从中南大学模板提炼出的排版规则和 Word 样式说明。
- `assets/csu-thesis-template.docx`：由原始 `.doc` 转换得到的可复用中南大学论文 DOCX 模板。
- `assets/附件6：中南大学毕业设计(论文)模版.doc`：用户提供的学校原始 Word 模板。
- `scripts/extract_docx_format.py`：提取 DOCX 的结构、样式、节、表格和段落样例。
- `scripts/check_csu_thesis_docx.py`：按中南大学论文规则快速检查 DOCX。
