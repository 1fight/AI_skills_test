# docx-to-markdown-with-format-extractor（stdout-json / no-open / no-json-files）

## 一、Skill 职责总述（Purpose）

本 Skill 的职责是：

通过调用 **officecli** 工具，从 Word（.docx）文档中读取并提取【全部内容及其格式语义】，并将其转换为一个结构清晰、信息不丢失的 **Markdown 中间表示**。

本 Skill 的职责范围包含：

- 调用 officecli 命令读取 docx（只读）
- 解析 officecli 的 **JSON 标准输出（stdout）**
- 按规则转换为 Markdown
- 显式表达格式语义（如加粗、对齐、表格结构、批注等）

本 Skill 不负责：

- 内容生成
- 内容改写 / 总结
- 写回 docx

---

## 二、工具依赖（Tooling）

- 必须使用：**officecli（跨平台：Windows / Linux / macOS）**
- officecli 是唯一允许的 docx 解析工具
- 所有结构与格式事实，必须来自 officecli 输出
- LLM 不允许直接“理解或解析 docx 文件”

---

## 三、运行前自检（必须遵守）

### 3.1 自动定位 officecli（必须）

Skill 必须先自动定位 `officecli` 可执行文件，不要求用户手动配置。

定位顺序：

1. 优先使用 PATH 中可执行命令：`officecli`
2. 若 PATH 不可用，按系统候选路径依次检测：
   - Windows：
     - `%LOCALAPPDATA%\\OfficeCli\\officecli.exe`
     - `%USERPROFILE%\\AppData\\Local\\OfficeCli\\officecli.exe`
   - Linux / macOS：
     - `/usr/local/bin/officecli`
     - `/usr/bin/officecli`
     - `~/.local/bin/officecli`

规则：

- 必须通过 `--version` 验证候选可执行文件。
- 一旦找到可执行路径，后续所有命令必须复用同一路径。
- 若全部不可用，立即报错并停止。

### 3.2 自动定位 DOCX 路径（必须）

- 输入参数为：`{{DOCX_PATH}}`
- 支持绝对路径和相对路径
- 若传入路径不存在，可做文件名级别检索（同名或高相似度），但必须把最终命中路径显式回报

### 3.3 会话内执行（禁止落地脚本）

- 转换逻辑必须在当前会话中执行
- 允许使用单行命令或内存变量
- **禁止**为转换流程创建 `*.ps1` / `*.sh` / `*.py` 临时脚本文件
- **禁止**将中间 JSON 落盘

### 3.4 不使用常驻模式

- **禁止**使用：`officecli open ...`
- **禁止**使用：`officecli close ...`

---

## 四、officecli 调用流程（必须严格执行）

**规则：**核心流程必须先执行以下三步；若核心输出不足以表达细节（如封面结构、章节层级、表格单元格文本、页眉 run 级样式），可进入“增强读取步骤”。所有读取命令都必须使用 `--json`，并直接解析 stdout；**不允许**将 JSON 输出重定向到文件。

### STEP 1：提取章节结构（标题 / outline）（必须）

必须调用：

```powershell
officecli view '{{DOCX_PATH}}' outline --json
```

stdout JSON 用于获取：

- 所有标题文本
- 标题层级（heading level）
- 标题顺序

该输出必须作为 Markdown 标题映射的唯一依据。

### STEP 2：提取正文、表格与 run 级格式（annotated）（必须）

必须调用：

```powershell
officecli view '{{DOCX_PATH}}' annotated --json
```

stdout JSON 用于获取：

- 正文段落及其顺序（不得改动）
- 每个段落内的 runs（及其 bold/italic/underline 等属性）
- 表格结构（行、列、单元格）及单元格内 runs
- 批注（如存在，按 officecli 输出呈现）
- 段落级信息（如 alignment/indent/spacing 等，**仅在 officecli 输出提供时**才可表达）

### STEP 3：提取页眉（Header）（必须）

必须调用：

```powershell
officecli get '{{DOCX_PATH}}' /header --depth 8 --json
```

stdout JSON 用于获取：

- 页眉段落与 runs
- 页眉文字格式信息
- 对齐方式（若 officecli 提供）

### STEP 4：增强读取（可选，但在细节不足时必须执行）

当出现以下任一情况，必须进入增强读取：

- `outline.data.headings` 为空或标题层级不足
- `annotated` 仅给出表格尺寸，未给出单元格文本
- 页眉只返回段落文本，未返回 run 级样式
- 用户明确要求“封面/章节/表格/页眉”细粒度还原

建议调用（按需）：

```powershell
officecli get '{{DOCX_PATH}}' /body --depth 6 --json
officecli get '{{DOCX_PATH}}' /body/p[n] --depth 6 --json
officecli get '{{DOCX_PATH}}' /body/tbl[n] --depth 10 --json
officecli get '{{DOCX_PATH}}' /header --depth 12 --json
```

规则：

- 增强读取仍然属于只读流程
- 仅可补充 officecli 明确给出的字段，不得猜测

---

## 五、LLM 的输入数据来源（必须遵守）

LLM 的输入只能来自以下来源的 **officecli stdout JSON（不落盘）**：

1. 核心三步：`outline` / `annotated` / `/header`
2. 增强读取：`/body`、`/body/p[n]`、`/body/tbl[n]`、更深层 `/header`

LLM 不允许：

- 直接读取 docx 文件
- 猜测或补全任何格式字段
- 推断缺失字段
- 使用“看起来合理”的方式修正结构

---

## 六、JSON 契约（Schema，必须遵守）

以下契约用于约束“每一块”可解析字段。所有字段仅以 officecli 实际 stdout 为准。

### 6.1 通用响应包装

所有命令都应符合：

```json
{
  "success": true,
  "data": { ... }
}
```

规则：

- `success` 必须存在且为布尔值
- `success=false` 时必须停止转换并报错
- `data` 缺失或类型异常时必须停止转换并报错

### 6.2 `view outline --json` 契约

最小字段集合（按你提供的 Word 实测）：

```json
{
  "success": true,
  "data": {
    "fileName": "string",
    "paragraphs": 0,
    "tables": 0,
    "images": 0,
    "equations": 0,
    "headers": ["string"],
    "footers": ["string"],
    "headings": []
  }
}
```

字段约束：

- `data.headings` 允许为空数组
- `data.headings` 非空时，标题文本与层级只能取该数组内容
- 若 `headings=[]`，必须触发“标题回退规则”或输出无章节说明

### 6.3 `view annotated --json` 契约

最小字段集合（按你提供的 Word 实测）：

```json
{
  "success": true,
  "data": {
    "View": "annotated",
    "Content": "string"
  }
}
```

其中 `Content` 是多行字符串，行级语法必须按以下模式解析：

1. 表格尺寸行：
  - `[/body/tbl[n]] [Table: RxC]`（`x` 可为 `×` / `x` / `X`）
2. 空段落行：
   - `[/body/p[n]] [] <- ...` 或 `[/body/p[n]] [] ← ...`
3. run 行：
   - `[/body/p[n]] 「文本」 <- ...` 或 `[/body/p[n]] 「文本」 ← ...`
4. TOC 行（目录线索）：
  - 行尾元信息包含 `toc N`（例如 `toc 1`、`toc 2`）

字段约束：

- 不得假设 `Content` 是 JSON 数组
- run 的样式仅能从行尾元信息中提取（如 `bold` / `italic` / `underline`）
- `Table: RxC` 只代表结构，不代表已拿到单元格文本
- 目录层级可仅从 `toc N` 或显式编号文本中提取，禁止凭语义猜层级

### 6.4 `get /header --depth N --json` 契约

最小字段集合（按你提供的 Word 实测）：

```json
{
  "success": true,
  "data": {
    "path": "/header",
    "type": "header",
    "text": "string",
    "childCount": 0,
    "format": {},
    "children": [
      {
        "path": "/header/p[1]",
        "type": "paragraph",
        "text": "string",
        "preview": "string",
        "style": "string",
        "childCount": 0,
        "format": {
          "effective.alignment": "center"
        },
        "children": []
      }
    ]
  }
}
```

字段约束：

- `data.children` 若存在，页眉段落必须按其顺序输出
- 对齐优先读取：`format.effective.alignment`，其次 `format.alignment`
- run 级样式仅在 `children` 中明确提供时才可输出

### 6.5 增强读取 `get /body` / `get /body/p[n]` / `get /body/tbl[n]` 契约

增强读取节点至少满足递归结构：

```json
{
  "success": true,
  "data": {
    "path": "string",
    "type": "string",
    "text": "string",
    "childCount": 0,
    "format": {},
    "children": []
  }
}
```

字段约束：

- `path` 必须可回溯到源节点（如 `/body/p[38]`、`/body/tbl[2]`）
- 表格单元格文本仅在节点树明确出现时可写入 Markdown
- 未出现的字段一律视为不存在，禁止补全

---

## 七、LLM 的任务与约束（Prompt 语义）

LLM 负责：

- 将 officecli 的结构化输出转换为 Markdown 表示
- 不遗漏信息（在 Markdown 可表达范围内）
- 不改变语义
- 保持顺序稳定

LLM 不负责：

- 判断格式是否“合理”
- 调整写作风格
- 合并内容
- 重排内容

---

## 八、Markdown 输出约定（必须遵守）

### 【0】封面（来自 body 前导内容）

封面内容应独立成块，不得混入正文章节。

```markdown
:::cover
...封面标题/副标题/作者/日期/单位等（按原顺序）...
:::
```

封面判定规则（只能基于 officecli 字段）：

- 优先：使用显式封面节点或 section 信息（若 officecli 提供）
- 次优：从文档起始位置到“首个章节标题节点”之前的段落/表格
- 若无法确定边界，不得猜测；应输出 `cover-note` 说明依据不足

### 【1】页眉（来自 /header）

页眉内容必须单独成块输出，不得混入正文。格式如下：

```markdown
:::header
...页眉内容（含 run 级格式）...
（align: center|left|right|justify，如 officecli 提供）
:::
```

- 若 header 有多个段落，必须按顺序逐段输出
- 若有 run 级样式，必须映射（bold/italic/underline）
- 若 officecli 未提供 alignment，则不得输出 align 标注行

### 【2】标题（来自 outline）

按照 heading level 映射为 Markdown 标题：

- level 1 → `#`
- level 2 → `##`
- level 3 → `###`
- level 4 → `####`（如果 officecli 输出存在更深层级，按此规律继续）
- 标题文本必须与 officecli 输出 **完全一致**
- 标题顺序必须与 officecli 输出一致

> 注意：标题映射的“层级与文本”只能来自 outline JSON，禁止从正文猜标题。

当 `outline` 为空时的严格回退规则：

- 允许按以下优先级回退（必须逐级尝试）：
  1. `body/p[n]` 的显式 heading 样式字段（如 `Heading1/标题1`）
  2. `annotated.Content` 行尾元信息中的 `toc N`
  3. 显式编号文本模式：`1 ...` 作为 H1，`1.1 ...` 作为 H2，`1.1.1 ...` 作为 H3
- 回退生成的标题必须标注来源：
  - `(heading-source: body-style-fallback)`
  - `(heading-source: annotated-toc-fallback)`
  - `(heading-source: numbered-text-fallback)`
- 若以上都不存在，才允许不生成章节标题，并输出 `heading-note`

### 【3】正文段落（来自 annotated）

- 段落顺序必须保持
- 不允许合并或拆分段落
- 段落之间使用空行分隔（仅作为 Markdown 语法需要，不代表合并）

#### 3.1 run 级格式映射（仅在 officecli 明确提供时）

- bold → `**文本**`
- italic → `*文本*`
- underline → `__文本__`（若 officecli 输出 `underline=true`）

若同一 run 同时包含多种格式：按嵌套表达，且不得凭空添加未提供的样式。  
例如同时 bold+italic：`***文本***`；若 underline 也为 true：`__***文本***__`。

#### 3.2 段落级格式标注（仅在 officecli 提供时）

若 officecli 为段落提供 alignment 等信息，可在段落文本后单独一行显式标注，例如：

```markdown
这一段的文字内容……
（alignment: center）
```

- 只能输出 officecli JSON 中明确存在的字段和值
- officecli 没给就不写

### 【4】表格（来自 annotated）

- 表格必须转换为 Markdown 表格
- 行列结构必须保持
- 单元格内仍需保留 run 级格式（bold/italic/underline）
- 不允许改变单元格内容顺序

若 `annotated` 只有 `Table: 行×列` 而无单元格文本：

- 必须使用 `get /body/tbl[n] --depth ... --json` 补齐单元格内容
- 若补齐后仍无单元格文本，允许输出空表格占位，但必须追加结构注记：
  - `table-note: source=/body/tbl[n], rows=?, cols=?, cells-missing=true`

表格最小输出约束（新增，必须）：

- 每个 `tbl[n]` 至少输出 3 行 Markdown 表格：
  - 表头行：`| ... |`
  - 分隔行：`| --- | ...`
  - 至少一行数据或空占位行
- 若未满足上述 3 行，视为表格输出失败

示例：

```markdown
| **模块名称** | **功能说明** |
|-------------|--------------|
| 调度模块    | 负责任务协调 |
```

若表格存在合并单元格等 Markdown 难以表达的结构：必须按 officecli 输出尽可能保留信息，并在表格后追加“结构注记”（仅描述 officecli 明确提供的结构事实，不得猜测）。

### 【5】批注 / 注释（如 officecli 输出存在）

- 必须保留批注内容与其关联位置（按 officecli 提供的锚点/范围）
- 表达格式建议（可根据 officecli 可用字段调整）：

```markdown
:::comment id=... author=... date=...
批注内容……
:::
```

若 officecli 未提供 author/date/id，则不得输出对应字段。

---

## 九、输出完整性检查（新增，必须）

生成 Markdown 前，必须进行一致性检查：

1. 是否输出了 `:::header`
2. 是否根据可用信息输出了 `:::cover` 或 `cover-note`
3. 是否按顺序输出正文段落（含空段）
4. 每个 `tbl[n]` 是否都有对应 Markdown 表格或 `table-note`
5. 标题来源是否满足“outline 优先、body-style 回退”规则
6. 在文档存在章节线索时，是否至少生成 1 个 `# ` 与 1 个 `## `
7. 若源文档存在表格（`outline.data.tables > 0`），Markdown 是否存在 `|` 表格语法行

若任一项不满足，必须报错或补齐，不可静默跳过。

---

## 十、用户要求与默认行为

- 若用户对某一内容区域（页眉 / 正文 / 表格 / 批注）提出明确要求：
  - 允许在 Markdown **表达层面**按要求调整（例如是否显示 alignment 行）
  - 但不得改写、总结、合并内容
- 若用户未提出要求：
  - 必须完全遵循 officecli 输出
  - 不得进行任何风格、结构、语义优化

---

## 十一、错误处理与输出要求（必须遵守）

### 11.1 officecli 失败

- 任一步命令非 0 退出码、或 stdout 不是有效 JSON：
  - Skill 必须停止转换并报告错误
  - 错误报告必须包含：
    - 哪一步失败（outline / annotated / header）
    - 原始错误信息（stderr / 解析失败信息）
  - 不允许继续“猜测式输出 Markdown”

### 11.2 缺失字段

- 如果 JSON 中缺少某些格式字段：
  - 必须视为“不存在该格式”
  - 不得补全

---

## 十二、明确禁止的行为（Do Not List）

- 不调用 officecli 而直接生成结果
- 忽略 officecli 的格式字段
- 为缺失信息“合理化补全”
- 总结、概括、改写内容
- 输出 Word 样式名（如 Heading1、Normal）
- 使用 `officecli open/close`
- 将 JSON 重定向写入 `outline.json / annotated.json / header.json`（不落盘）
- 为本次转换流程创建临时脚本文件（如 `run_docx_to_md.ps1`）

---

## 十三、最终目标

输出一个完整的、结构稳定的、带格式语义的 Markdown 中间表示，其所有内容与格式都可以一一追溯到 officecli 的原始 stdout JSON。