# 中南大学本科毕业论文排版 Skill

这个仓库把中南大学本科毕业设计(论文) Word 模板中的排版要求和真实修订流程整理成一个可复用的 Codex skill，并附带原始模板、转换后的 DOCX 模板、排版规则文档和自动检查脚本。

目标很简单：以后中国学生或任意 AI agent 只要读完这个仓库，就能更稳定地辅助中南大学本科毕业论文排版、格式检查和模板比对。

## 仓库内容

```text
.
├── SKILL.md
├── agents/
│   └── openai.yaml
├── assets/
│   ├── csu-thesis-template.docx
│   └── 附件6：中南大学毕业设计(论文)模版.doc
├── references/
│   ├── csu-thesis-format-rules.md
│   ├── csu-thesis-revision-playbook.md
│   ├── csu-thesis-real-world-failure-modes.md
│   ├── csu-thesis-auto-toc-workflow.md
│   ├── csu-thesis-word-equation-objects.md
│   └── formula-object-config-example.json
│   └── inline-subscript-config-example.json
└── scripts/
    ├── repair_formula_objects_with_word.ps1
    ├── unify_inline_subscripts.py
    ├── normalize_citations_and_features.py
    ├── audit_csu_word_structures.py
    ├── check_csu_thesis_docx.py
    ├── extract_docx_format.py
    ├── test_unify_inline_subscripts.py
    ├── test_normalize_citations_and_features.py
    └── test_audit_csu_word_structures.py
```

主要文件说明：

- `SKILL.md`：给 Codex/agent 读取的 skill 入口，说明什么时候使用、如何检查和修改论文格式。
- `references/csu-thesis-format-rules.md`：从模板中提炼出的中文排版规则，是主要规范依据。
- `references/csu-thesis-revision-playbook.md`：从真实论文修订中提炼出的稳定工作顺序。
- `references/csu-thesis-real-world-failure-modes.md`：目录重复、标题变蓝、页眉异常、公式编号等高频坑点。
- `references/csu-thesis-auto-toc-workflow.md`：把手打目录改成真正 Word 自动目录时的分节、页码重启、TOC 样式和验收流程。
- `references/csu-thesis-word-equation-objects.md`：把伪公式恢复成真正 Word 公式对象的验证方法、版式规则和脚本用法。
- `references/inline-subscript-config-example.json`：统一正文变量下标风格的示例配置。
- `assets/附件6：中南大学毕业设计(论文)模版.doc`：学校原始 Word 模板。
- `assets/csu-thesis-template.docx`：由原始 `.doc` 转换得到的 DOCX 模板，方便脚本和现代 Word 工具处理。
- `scripts/check_csu_thesis_docx.py`：按中南大学模板规则对 DOCX 做启发式格式检查。
- `scripts/extract_docx_format.py`：提取 DOCX 的分节、样式、表格和段落格式信息，便于排版前诊断。
- `scripts/audit_csu_word_structures.py`：专项审计 Word 结构层问题，如页眉页脚、目录域、页码分节、公式对象和展示公式编号。
- `scripts/repair_formula_objects_with_word.ps1`：在 Windows + Microsoft Word 环境下，用 Word COM 把展示公式重建为真正的公式对象。
- `scripts/unify_inline_subscripts.py`：把第 2 章正文里的变量名从裸下划线文本统一成真正的下标 run。
- `scripts/normalize_citations_and_features.py`：把正文文献引用统一成上标，同时把 `dWC/fWC/xNi/aWC/Asp/Tnorm/fWC2/3/R2/cm-3` 等特征和指数统一成真正的下标/上标 run。

## 适用场景

- 检查中南大学本科毕业论文是否符合学校模板。
- 从学校 Word 模板中快速查找页边距、字号、标题、图表、公式、参考文献等规则。
- 让 Codex、ChatGPT 或其他 agent 辅助论文排版。
- 对已有 DOCX 论文进行格式体检，找出可能需要人工修正的地方。
- 作为后续学院/专业论文模板 skill 的基础。

## 给学生的快速用法

如果你只是想看排版规则，直接阅读：

```text
references/csu-thesis-format-rules.md
```

如果你想让 agent 帮你排版，把本仓库发给它，并让它先读：

```text
SKILL.md
references/csu-thesis-format-rules.md
```

然后把你的论文 DOCX 交给 agent，让它先运行检查脚本，再按规则修改。

## 给 Codex/agent 的用法

当用户请求“中南大学论文排版”“毕业论文格式检查”“按中南大学模板修 Word”时，应优先读取：

```text
SKILL.md
references/csu-thesis-format-rules.md
```

处理流程建议：

1. 先备份用户论文。
2. 使用 `extract_docx_format.py` 提取现有格式。
3. 使用 `check_csu_thesis_docx.py` 生成启发式检查报告。
4. 使用 `audit_csu_word_structures.py` 专查 Word 结构层问题。
5. 继续阅读 `references/csu-thesis-revision-playbook.md`，按真实定稿顺序处理 section、页码、目录、页眉页脚和公式。
6. 遇到目录重复、标题变蓝、页眉不对、公式像正文等问题时，查 `references/csu-thesis-real-world-failure-modes.md`。
7. 如果目录是手打静态文本、需要改成真正可更新 TOC，或目录更新后第 1 章页码被带乱，继续读 `references/csu-thesis-auto-toc-workflow.md`。
8. 遇到 `OMaths.Count = 0`、公式只是文本、或必须恢复为真正 Word 公式对象时，继续读 `references/csu-thesis-word-equation-objects.md`，必要时直接用 `scripts/repair_formula_objects_with_word.ps1`。
9. 如果正文引用还是普通 `[1]`、`[1-3]`，或表格外特征仍是 `dWC/fWC/xNi/aWC/Asp/Tnorm/fWC2/3` 这种紧凑普通文本，优先用 `scripts/normalize_citations_and_features.py` 处理。
10. 如果正文变量出现 `f_WC`、`x_Co`、`T_norm` 这种下划线混排不统一的问题，也可用 `scripts/unify_inline_subscripts.py` 配合 `references/inline-subscript-config-example.json` 处理。
11. 导出 PDF 或页面截图进行视觉检查，重点查看封面、摘要、目录、每章首页、公式页、图表页和参考文献页。

## 安装为 Codex Skill

如果要让 Codex 自动发现这个 skill，可以把仓库克隆到 Codex skills 目录：

```bash
git clone https://github.com/CTctikki/csu-thesis-format-Skill.git ~/.codex/skills/csu-thesis-format
```

Windows PowerShell 示例：

```powershell
git clone https://github.com/CTctikki/csu-thesis-format-Skill.git "$env:USERPROFILE\.codex\skills\csu-thesis-format"
```

之后可以用类似下面的提示词调用：

```text
使用 $csu-thesis-format 帮我检查并规范中南大学本科毕业论文 DOCX。
```

## 脚本依赖

需要 Python 3，以及以下 Python 包：

```bash
pip install python-docx lxml
```

如果要修复真正的 Word 公式对象，还需要：

- Windows
- Microsoft Word
- PowerShell

如果需要把旧版 `.doc` 转成 `.docx`，建议安装 LibreOffice，并使用：

```bash
soffice --headless --convert-to docx --outdir output_dir input.doc
```

## 检查论文格式

运行：

```bash
python scripts/check_csu_thesis_docx.py path/to/你的论文.docx
```

输出内容包括：

- 页面尺寸和页边距问题
- 摘要、目录、参考文献等标题格式问题
- 正文标题层级问题
- 图题、表题和公式编号问题
- 关键词数量和参考文献编号连续性提示

注意：该脚本是启发式检查工具，不能替代人工审阅。Word 文档可能包含文本框、域代码、直接格式、图片公式等复杂内容，最终仍应导出 PDF 后目视确认。

## 统一引用上标和特征上下标

运行：

```bash
python scripts/normalize_citations_and_features.py input.docx output.docx
```

这个脚本用于处理两类高频定稿问题：

- 正文引用 `[1]`、`[1-3]`、`[14,15]` 改为真正的上标 run，并跳过 `参考文献` 标题之后的文献列表编号。
- 特征符号统一成真正的下标/上标 run，例如 `d_WC`、`f_WC`、`x_Ni`、`a_WC`、`A_sp`、`T_norm`、`f_WC^(2/3)`、`R^2`、`cm^-3`、`μm^-1`、`mm^2/s`。

## 审计 Word 结构层

运行：

```bash
python scripts/audit_csu_word_structures.py path/to/你的论文.docx
```

这个脚本重点输出：

- 分节、页码格式和页眉页脚关系
- 自动目录域与静态目录并存风险
- 页眉图片/图片关系文件是否存在
- OMath 公式对象数量
- 展示型公式候选和编号连续性

## 提取文档样式

运行：

```bash
python scripts/extract_docx_format.py path/to/你的论文.docx
```

适合在正式修改前使用，帮助了解当前论文的：

- 分节数量和页面设置
- 段落样式使用情况
- 表格数量和基本结构
- XML 中的段落和字符属性

## 模板规则摘要

完整规则见 `references/csu-thesis-format-rules.md`。常用要求包括：

- A4 纵向，页边距上 2.5 cm、下 2.5 cm、左 3.0 cm、右 2.0 cm。
- 正文 1.5 倍行距，中文小四号宋体，英文小四号 Times New Roman。
- 正文段落首行缩进两个汉字。
- 每章单独起页。
- 章标题三号黑体居中，上下各空一行。
- 二级标题小四号黑体，三级标题小四号楷体GB2312。
- 表题在表格上方居中，图题在图片下方居中。
- 参考文献按出现顺序编号，使用方括号序号。

## 维护说明

更新模板或规则时，建议按以下顺序操作：

1. 替换或新增 `assets/` 中的模板文件。
2. 重新提取模板样式和说明文字。
3. 更新 `references/csu-thesis-format-rules.md`。
4. 如检查逻辑变化，更新 `scripts/check_csu_thesis_docx.py`。
5. 用模板和至少一篇真实论文试跑两个脚本，并手动核对目录、页眉页脚和公式页。

## 重要说明

- 本仓库用于学习、论文排版辅助和 agent 工作流复用。
- 原始 Word 模板来自用户提供的中南大学毕业设计(论文)模板文件，请按学校要求和实际学院通知使用。
- 不同学院、年份或专业可能存在补充要求；正式提交前请以当年学校/学院最新通知为准。
- 自动脚本只能发现常见格式问题，不能保证论文内容、学术规范、查重、引用准确性或最终版式完全合格。
