# Source to Docx Structured Writer

## Skill 目标（Purpose）

本 Skill 用于在不破坏既有 Word 模板（.docx）的前提下，基于源码分析结果，按照每一章节的写作要求，生成高质量、可交付的技术文档正文内容，并安全地注入到 Word 文档中。

本 Skill 必须保证：
- 不修改模板的封面
- 不修改页眉 / 页脚
- 不修改样式定义
- 不修改分节（Section）
- 不进行任何 raw XML 操作
- 只在正文区域写入内容

## 核心设计原则（Design Principles）

### 1. 模板不可破坏原则
- Word 模板被视为"只读结构"
- 任何结构性修改都是禁止行为

### 2. 职责分离原则
- LLM：负责源码分析与内容生成
- officecli：负责将内容安全写入 Word
- officecli 不做任何语义判断或分析

### 3. 基于文档结构定位
- 不依赖人工锚点
- 使用 Word 的 Heading（标题样式）作为隐式锚点

### 4. 一次性扫描原则（核心）
- **只在 STEP 1 一次性扫描 Word 文档**
- 提取所有章节标题和写作要求后，立即关闭文档
- 将所有信息存入内存结构
- 后续所有操作只使用内存结构，不再读取 Word 文档
- **禁止读取 Word 文档的批注、注释等内容**

### 5. 自动提取写作要求
- 从 Word 模板中自动读取每个章节的写作要求
- 写作要求可能没有特定格式，使用 LLM 智能识别
- 用户不应手动复制任何要求文本
- 写作要求作为 Prompt 的一部分传给 LLM

## 输入定义（Inputs）

### 输入一：template_docx
一个已经设计完成的 Word 模板，包含：
- 封面
- 页眉 / 页脚
- 目录
- 已定义好的样式
- 正文中包含多个使用 Heading 样式的章节标题
- 每个章节标题后包含写作要求（使用特定样式或格式标识）

### 输入二：source_code
项目源码，可以是：
- 完整代码
- 多文件
- 经摘要后的源码内容

### 输入三：requirements_style（可选）
用于标识写作要求的样式名称。如果未提供，Skill 将尝试自动识别写作要求段落。

## 执行流程（Execution Flow）

### STEP 1：一次性扫描模板结构并提取所有章节写作要求

**重要：此步骤只执行一次，之后不再读取 Word 文档。**

执行命令：
```bash
officecli view template.docx outline --json
```

获取 Word 文档中所有标题结构，包括：
- 标题文本
- 标题级别（Heading1 / Heading2）
- 顺序位置

同时，对每个章节标题，执行以下操作：

1. 识别该标题后的写作要求段落：
   - 写作要求可能是一段或多段普通文本，没有特定格式标记
   - 使用 LLM 智能识别：分析标题后的段落内容，判断哪些是写作要求
   - 写作要求通常包含：目标、作用、内容特点、编写要求、注意事项等关键词
   - 收集标题后到下一个同级标题之间的所有段落

2. 提取写作要求文本：
   - 保留完整的写作要求内容
   - 去除格式标记，只保留纯文本
   - 如果找到多个要求段落，合并为一个要求文本

3. 将所有章节信息存入内存结构：
```json
{
  "sections": [
    {
      "section_id": "SECTION_ID",
      "title": "章节标题",
      "heading_level": 1,
      "paragraph_index": 123,
      "requirements": "写作要求文本内容",
      "content_start_index": 125,
      "content_end_index": 140
    }
  ]
}
```

**关键约束：**
- 此步骤完成后，关闭 Word 文档
- 后续所有操作只使用内存中的结构
- 不再读取 Word 文档的任何内容（包括批注、注释等）

### STEP 2：章节与标题匹配验证

对提取的每个章节信息：

1. 验证标题存在性：
   - 确保在模板中找到了对应的标题
   - 确保标题级别匹配

2. 验证写作要求存在性：
   - 确保每个章节都有写作要求
   - 如果某个章节缺少写作要求，发出警告但继续处理

3. 若未找到标题或写作要求：
   - Skill 必须立即终止
   - 返回错误："未在模板中找到所需章节标题或写作要求"

### STEP 3：章节内容范围定义

对于已匹配的标题，在内存结构中定义内容范围：

- 章节正文范围定义为：
  - 从：当前标题和写作要求之后的第一个正文段落（content_start_index）
  - 到：下一个"同级标题"之前的段落（content_end_index）

- Skill 允许：
  - 删除该范围内的正文段落

- Skill 禁止：
  - 删除或修改任何标题
  - 删除或修改写作要求段落
  - 修改更高或更低级别的标题

### STEP 4：按章节调用 LLM 生成内容

每一个章节必须独立调用一次 LLM。

#### 系统提示词（System Prompt）
```
你是一名技术文档写作 AI，你的工作是分析源码并撰写专业的技术文档内容。
你不得生成标题、格式说明、样式信息或任何标记语言。
你的输出必须是可直接插入 Word 正文的自然语言段落。
```

#### 用户提示词模板（User Prompt）
```
章节标题：{section.title}

章节写作要求：
{section.requirements}

源码内容：
{source_code}

输出要求：
- 只输出章节正文内容
- 不输出章节标题
- 不提及 Word、样式或格式
- 不贴代码（除非写作要求明确要求）
- 严格按照写作要求生成内容
```

### STEP 5：中间内容缓冲（Intermediate Buffer）

- 将 LLM 输出按空行拆分为多个段落
- 形成如下结构：
```json
{
  "section_id": "SECTION_ID",
  "paragraphs": [
    "第一段正文内容",
    "第二段正文内容",
    "第三段正文内容"
  ]
}
```

### STEP 6：内容写入 Word（officecli 渲染）

**注意：此时才重新打开 Word 文档进行写入操作。**

对每个章节执行：

1. 清空该章节范围内原有正文段落（保留标题和写作要求）
   - 使用内存结构中的 content_start_index 和 content_end_index 定位范围

2. 逐段写入新内容：
```bash
officecli add template.docx \
  /paragraph[after_heading:"章节标题"] \
  --type paragraph \
  --prop style=BodyText \
  --prop text="段落正文内容"
```

约束规则：
- 只允许使用模板中已存在的样式
- 不允许创建或修改样式
- 不允许在章节范围之外写入内容
- 不修改标题和写作要求段落
- **不读取 Word 文档的任何批注、注释等内容**

### STEP 7：后处理（可选）

- 如模板包含目录（TOC），可统一更新：
```bash
officecli set output.docx '/toc[1]' --prop update=true
```

禁止任何结构性操作：
- 不修改页眉
- 不修改页脚
- 不修改分节
- **不读取 Word 文档的批注、注释等内容**

## 安全校验规则（Validation）

Skill 必须校验：

1. 所有章节标题均成功匹配
2. 所有章节都有写作要求
3. 未引入新的样式
4. 注入前后：
   - 章节数量一致
   - 页眉 / 页脚数量一致
   - 标题和写作要求段落保持不变
5. 所有内容仅位于合法章节范围内
6. **整个过程中，只在 STEP 1 和 STEP 6 两个时刻访问 Word 文档**

任一校验失败 → Skill 终止并报错。

## 最终输出（Final Output）

输出文件：output.docx

保证：
- 模板结构完整不变
- 内容符合章节写作要求
- 文档可直接交付使用

## 核心原则（Core Principle）

本 Skill 不负责"生成 Word 文档"。

本 Skill 的职责是：
1. 从 Word 模板中自动提取写作要求
2. 基于源码分析生成内容
3. 将内容安全地写入一个已经设计完成的 Word 模板中

## 依赖的 Skill

本 Skill 依赖以下 Skill：
- minimax-docx：用于 Word 文档的读取、分析和写入操作

## 使用示例

### 基本用法

```bash
# 使用默认设置
source-to-docx-structured-writer \
  --template template.docx \
  --source-code ./src \
  --output output.docx
```

### 指定写作要求样式

```bash
source-to-docx-structured-writer \
  --template template.docx \
  --source-code ./src \
  --requirements-style "写作要求" \
  --output output.docx
```

### 使用源码摘要

```bash
source-to-docx-structured-writer \
  --template template.docx \
  --source-code-summary "项目源码摘要内容..." \
  --output output.docx
```

## 注意事项

1. **模板准备**：确保 Word 模板中的每个章节都有明确的写作要求
2. **样式一致性**：写作要求应使用一致的样式或标记格式
3. **源码质量**：提供完整、结构清晰的源码有助于生成更好的文档
4. **备份模板**：在执行前建议备份原始模板文件
5. **逐步验证**：对于大型文档，建议先测试单个章节的生成效果

## 故障排查

### 问题：未找到章节标题
**解决方案**：检查模板中的标题是否使用了正确的 Heading 样式

### 问题：未找到写作要求
**解决方案**：
- 检查写作要求是否使用了指定的样式
- 或检查写作要求是否包含明确的标记
- 或调整 requirements_style 参数

### 问题：生成的内容不符合要求
**解决方案**：
- 检查写作要求是否清晰明确
- 检查源码是否完整
- 调整 LLM 的提示词参数

### 问题：Word 文档结构被破坏
**解决方案**：
- 立即停止使用
- 检查 officecli 命令是否正确
- 使用备份的模板文件重新开始
'@"