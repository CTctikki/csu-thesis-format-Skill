---
name: csu-thesis-format
description: Use when editing, auditing, or finalizing Central South University undergraduate thesis DOCX/DOC files, especially for cover pages, abstract/TOC pagination, section breaks, headers/footers, heading styles and colors, formulas, figures/tables, and final submission formatting in Word.
---

# 中南大学论文排版

## Overview
把中南大学本科论文排版当成 `Word 结构问题`，不只是文本样式问题。目录、页码、页眉、公式编号、标题颜色链都可能跨 `section / field / styles.xml / theme1.xml / header rels` 同时出问题。

## When to Use
- 用户提到“中南大学”“本科毕业设计(论文)”模板、学长论文、学院格式。
- 需要处理 `.docx` 或旧 `.doc` 的论文定稿排版。
- 需要核对封面、摘要、目录、页眉页脚、页码、公式、图表、参考文献。
- 出现真实 Word 疑难：目录重复、目录页码漂移、标题在预览里变蓝、页眉和学长不一样、公式只是文本、文件只读/锁定、Word 和预览不一致。

## Required Workflow
1. 先读 `references/csu-thesis-format-rules.md`，锁定学校硬规则。
2. 如果是接近提交的真实论文，继续读 `references/csu-thesis-revision-playbook.md`。
3. 如果出现 Word/预览/目录/页眉/颜色异常，再读 `references/csu-thesis-real-world-failure-modes.md`。
4. 编辑前先备份并复制成新文件版本，不在原件上硬覆盖。
5. 先跑：
```bash
python scripts/extract_docx_format.py path/to/thesis.docx
python scripts/check_csu_thesis_docx.py path/to/thesis.docx
```
6. 按固定顺序修：
   - `section / 分页 / 页码`
   - `封面 / 中文摘要 / 英文摘要 / 目录`
   - `正文标题 / 正文段落 / 标题颜色链`
   - `图题 / 表题 / 表格 / 公式 / 参考文献`
7. 自动目录、页码域、页眉页脚、字段更新放到最后一轮，在 Word 中完成。
8. 最终必须渲染并人工复核关键页：封面、摘要、目录、正文首页、公式页、图表密集页、参考文献页。

## Hard Rules
- `section` 先于页码。封面、摘要、英文摘要、目录、正文至少分开思考。
- 自动目录只适合作为定稿动作。脚本阶段允许静态目录占位，最终交导师时优先交 `静态定稿目录`。
- 标题颜色要改 `样式层 + 链接字符样式 + 主题色链`，不能只刷某个 run。
- 公式编号必须全文只用一种方案；推荐按章编号，如 `(2-1)`。
- 渲染器不是 Word。预览通过不等于 Word 通过；最终判断以 Word 真实分页和字段显示为准。
- 文件被 Word 占用时，优先关掉后台 `WINWORD` 或另存新文件，不在锁文件上硬写。

## Quick Commands
旧版 `.doc` 转 `.docx`：
```bash
soffice --headless --convert-to docx --outdir output_dir input.doc
```

提取结构和样式：
```bash
python scripts/extract_docx_format.py path/to/thesis.docx
```

启发式排版检查：
```bash
python scripts/check_csu_thesis_docx.py path/to/thesis.docx
```

## Resources
- `references/csu-thesis-format-rules.md`
  学校模板硬规则：页面、字体字号、标题、目录、公式、图表、参考文献。
- `references/csu-thesis-revision-playbook.md`
  真实论文从“草稿”到“可交导师终稿”的操作顺序。
- `references/csu-thesis-real-world-failure-modes.md`
  这次实操验证过的高频坑点、症状、根因和修法。
- `assets/csu-thesis-template.docx`
  可供脚本和 Word 现代工具直接复用的 DOCX 模板。
- `assets/附件6：中南大学毕业设计(论文)模版.doc`
  学校原始模板，用于溯源比对。
- `scripts/extract_docx_format.py`
  看 section、style、header/footer、table、XML 样例。
- `scripts/check_csu_thesis_docx.py`
  快速发现格式和结构风险，尤其是目录、页码、公式、标题颜色链等。
