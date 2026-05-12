---
name: csu-thesis-format
description: Use when editing, auditing, or finalizing Central South University undergraduate thesis DOCX/DOC files, especially for cover pages, abstract/TOC pagination, section breaks, headers/footers, heading styles and colors, formulas, figures/tables, citation superscripts, inline feature subscripts/superscripts, and final submission formatting in Word.
---

# 中南大学论文排版

## Overview
把中南大学本科论文排版当成 `Word 结构问题`，不只是文本样式问题。目录、页码、页眉、公式编号、标题颜色链都可能跨 `section / field / styles.xml / theme1.xml / header rels` 同时出问题。

## When to Use
- 用户提到“中南大学”“本科毕业设计(论文)”模板、学长论文、学院格式。
- 需要处理 `.docx` 或旧 `.doc` 的论文定稿排版。
- 需要核对封面、摘要、目录、页眉页脚、页码、公式、图表、参考文献。
- 出现真实 Word 疑难：目录重复、目录页码漂移、目录是手打静态目录而不是可更新 TOC、标题在预览里变蓝、页眉和学长不一样、正文引用不是上标、公式只是文本、公式对象被改坏、正文变量/特征只有普通文本或下划线不是下标、指数不是上标、文件只读/锁定、Word 和预览不一致。

## Required Workflow
1. 先读 `references/csu-thesis-format-rules.md`，锁定学校硬规则。
2. 如果是接近提交的真实论文，继续读 `references/csu-thesis-revision-playbook.md`。
3. 如果出现 Word/预览/目录/页眉/颜色异常，再读 `references/csu-thesis-real-world-failure-modes.md`。
4. 如果目录是手打静态文本、需要改成真正可更新的自动目录、或者目录更新后正文页码/分节被带乱，再读 `references/csu-thesis-auto-toc-workflow.md`。
5. 如果展示公式只是文本、`OMaths.Count = 0`、或者必须恢复为真正的 Word 公式对象，再读 `references/csu-thesis-word-equation-objects.md`。
6. 编辑前先备份并复制成新文件版本，不在原件上硬覆盖。
7. 先跑：
```bash
python scripts/extract_docx_format.py path/to/thesis.docx
python scripts/check_csu_thesis_docx.py path/to/thesis.docx
python scripts/audit_csu_word_structures.py path/to/thesis.docx
```
8. 按固定顺序修：
   - `section / 分页 / 页码`
   - `封面 / 中文摘要 / 英文摘要 / 目录`
   - `正文标题 / 正文段落 / 标题颜色链`
   - `图题 / 表题 / 表格 / 公式 / 引用上标 / 特征上下标 / 参考文献`
9. 自动目录、页码域、页眉页脚、字段更新放到最后一轮，在 Word 中完成。
10. 最终必须渲染并人工复核关键页：封面、摘要、目录、正文首页、公式页、图表密集页、参考文献页。

## Hard Rules
- `section` 先于页码。封面、摘要、英文摘要、目录、正文至少分开思考。
- 自动目录优先服务于“还要继续改正文”的版本；如果是最终稳定交付版，先在 Word 中更新，再决定是否保留自动 TOC 或静态定稿目录。
- 标题颜色要改 `样式层 + 链接字符样式 + 主题色链`，不能只刷某个 run。
- 公式编号必须全文只用一种方案；推荐按章编号，如 `(2-1)`。
- 当公式已经被改坏、必须恢复为真正的 Word 公式对象时，优先走 `Word COM + OMaths.Add() + BuildUp()` 路线，不要再用 `python-docx` 直接覆盖公式正文。
- 正文引用统一用方括号数字并设置为真正的上标 run，如 `[1]`、`[1-3]`、`[14,15]`；不要把参考文献列表开头的 `[1]`、`[2]` 改成上标。
- 特征符号统一用真正的下标/上标 run，不要只用普通字符凑外观：`d_WC`、`f_WC`、`x_Ni`、`x_Co`、`a_WC`、`A_sp`、`T_norm` 等下标化，`f_WC^(2/3)`、`R^2`、`cm^-3`、`μm^-1`、`mm^2/s` 等指数上标化。
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

统一正文引用上标和特征上下标：
```bash
python scripts/normalize_citations_and_features.py input.docx output.docx
```

## Resources
- `references/csu-thesis-format-rules.md`
  学校模板硬规则：页面、字体字号、标题、目录、公式、图表、参考文献。
- `references/csu-thesis-revision-playbook.md`
  真实论文从“草稿”到“可交导师终稿”的操作顺序。
- `references/csu-thesis-real-world-failure-modes.md`
  这次实操验证过的高频坑点、症状、根因和修法。
- `references/csu-thesis-auto-toc-workflow.md`
  把手打目录改成真正 Word 自动目录时的分节、页码、标题层级、TOC 样式和验收流程。
- `references/csu-thesis-word-equation-objects.md`
  展示公式从伪公式恢复为真正 Word 公式对象时的操作方法、验证原则和脚本用法。
- `references/inline-subscript-config-example.json`
  第 2 章变量下标统一的示例配置，可直接按段落范围和变量名复用。
- `assets/csu-thesis-template.docx`
  可供脚本和 Word 现代工具直接复用的 DOCX 模板。
- `assets/附件6：中南大学毕业设计(论文)模版.doc`
  学校原始模板，用于溯源比对。
- `scripts/extract_docx_format.py`
  看 section、style、header/footer、table、XML 样例。
- `scripts/check_csu_thesis_docx.py`
  快速发现格式和结构风险，尤其是目录、页码、公式、标题颜色链等。
- `scripts/audit_csu_word_structures.py`
  专查 Word 结构层：分节页码、目录域/静态目录、页眉页脚图片关系、公式对象与展示公式编号。
- `scripts/repair_formula_objects_with_word.ps1`
  在 Windows + Microsoft Word 环境下，把展示公式表格重建成真正的 Word 公式对象。
- `scripts/unify_inline_subscripts.py`
  把正文中的 `f_WC`、`x_Co`、`T_norm` 一类变量统一成真正的下标 run，而不是裸下划线文本。
- `scripts/normalize_citations_and_features.py`
  把正文引用 `[1]`、`[1-3]` 统一成上标，同时把 `dWC/fWC/xNi/aWC/Asp/Tnorm/fWC2/3/R2/cm-3` 等特征和指数拆成真正的下标/上标 run；会跳过参考文献列表条目编号。
