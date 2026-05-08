# 中南大学论文 Word 公式对象修复

这份说明只沉淀一条已经在真实论文上验证过的方法：当展示公式已经被改坏、只是普通文本，或者必须按模板要求恢复为真正的 Word 公式对象时，不再用 `python-docx` 直接改公式正文，而是切换到 **Windows + Microsoft Word COM** 来重建公式对象。

## 什么时候必须用这条方法

- `Word OMaths.Count = 0`，但论文里明明存在多条展示公式。
- 公式看起来像普通正文，变量带下划线、编号靠空格推右。
- 前一轮脚本或批量替换把原来的公式排版冲坏了。
- 已经接近交稿，需要公式对象、编号、续行位置都稳定符合模板。

## 什么时候不要先用这条方法

- 还在早期草稿阶段，只是想先统一“公式单独成行 + 行末编号”。
- 公式很少，且人工在 Word 里逐条重建更快。
- 当前机器没有 Microsoft Word，只有 LibreOffice 或纯脚本环境。

## 正确方法

### 1. 先把公式问题当成 Word 结构问题

中南大学论文里的展示公式，稳定做法不是“公式正文一个段落 + 末尾空格 + 编号”，而是：

- 左侧空白占位单元格
- 中间公式本体单元格
- 右侧编号单元格

短公式：

- 单独一行
- 公式居中
- 编号在同一行最右侧

长公式：

- 优先在 `=` 处转行；否则在 `+`、`-`、`×`、`÷` 等符号处转行
- 最好让续行与上一行的关键数学位置对齐
- 编号只放在最后一行的右端

### 2. 不要再用这几种方法硬改公式

- 不要用 `python-docx` 的 `paragraph.text = ...` 或 `cell.text = ...` 批量替公式正文。
- 不要把带下划线的变量文本直接当作“已经是公式”。
- 不要只看第三方渲染器；真正判断要回到 Word 本体或 Word 导出的 PDF。

这些做法很容易把原本还能用的公式区域彻底压扁成普通文本。

### 3. 用 Word COM 重建 OMath

验证过的流程是：

1. 复制出新版本文件。
2. 用 Word COM 打开副本。
3. 找到承载展示公式的三列表格。
4. 清空公式单元格和编号单元格。
5. 把 **Word 线性公式语法** 写入中间单元格。
6. 对该单元格调用：

```powershell
$range.OMaths.Add($range) | Out-Null
$range.OMaths.Item(1).BuildUp()
```

7. 右侧单元格单独写 `(2-1)`、`(2-2)` 这种编号。
8. 长公式拆成多行时，让同一个表格包含多行，最后一行右侧才放编号。
9. 保存后，用 Word 自己导出 PDF 再检查公式页。

## 可直接复用的脚本

脚本：

```text
scripts/repair_formula_objects_with_word.ps1
```

示例配置：

```text
references/formula-object-config-example.json
```

调用方式：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/repair_formula_objects_with_word.ps1 `
  -InputDocx path\to\thesis.docx `
  -OutputDocx path\to\thesis_formula_fixed.docx `
  -ConfigJson references\formula-object-config-example.json
```

## 配置文件说明

每个公式表格都用一个 `tableIndex` 描述。索引是 **Word COM 的 1 基表格索引**，不是 `python-docx` 的 0 基索引。

每个 `rows` 元素表示表格中的一行：

- `formula`：Word 线性公式语法
- `number`：该行右侧编号；若是长公式的中间行可留空
- `align`：`center` 或 `left`
- `leftIndentCm`：续行缩进，长公式常用

脚本还支持几个 ASCII 宏，避免在 JSON 里硬塞 Unicode：

- `[RHO]` -> `ρ`
- `[TIMES]` -> `×`
- `[SUM]` -> `∑`
- `[ALPHA]` / `[BETA]` / `[GAMMA]` / `[ETA]` / `[THETA]` / `[LAMBDA]` / `[MU]` / `[PI]`

## 第 2 章变量写法统一

当公式对象修好后，正文里经常还残留一批“假下标变量”，例如：

- `f_WC`
- `x_Co`
- `x_Ni`
- `d_WC`
- `a_WC`
- `A_sp`
- `T_norm`

如果论文已经进入定稿阶段，建议：

- 公式区用真正的 Word 公式对象
- 正文说明中的变量用统一下标风格
- 不要一部分是 `f_WC`，一部分是 `fWC`，一部分又是图片公式截图

## 验证要求

至少人工检查：

- 第一处展示公式页
- 第一条长公式续行页
- `(2-5)`、`(2-6)`、`(2-7)` 这种连续短公式页
- `MAPE` 或其他含求和符号的公式页

确认：

- 公式对象确实显示成数学公式，而不是普通文本
- 编号在右端，且连续
- 长公式编号落在最后一行
- 续行位置没有跑偏
- “式中”说明没有混入公式对象链

## 与审计脚本配合

如果 `scripts/audit_csu_word_structures.py` 输出：

- `OMath 对象数 = 0`
- 但展示公式候选数很多

就应该优先怀疑：

1. 文档里的展示公式只是伪公式；
2. 或者上一轮脚本把公式对象改坏了；
3. 这时应切到 `repair_formula_objects_with_word.ps1` 这条路线，而不是继续用 `python-docx` 改公式正文。
