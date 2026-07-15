---
name: course-review-pipeline
description: >-
  期末复习笔记处理流水线：扫描空中课堂 session 目录、合并同日双场次、提取加密 PDF 和 PPTX
  复习资料为幻灯片截图、识别关键图解、分批补充 session_notes（可按环境与用户要求启用并行代理；默认也可顺序执行）（嵌入图片+法条+
  案例+答题框架）、幻灯片对账与逐张验证、生成索引和验证报告。支持混合源（同一课程含
  加密PDF+PPTX+非加密PDF）、多教师各自对应复习资料、任意原始编号映射。适用于法学院
  或其他学科的期末复习包制作。触发词：期末复习、复习包、课程笔记整合、session 合并、
  复习资料处理、空中课堂笔记、exam review pipeline。
version: 1.1.0
---

# 期末复习笔记处理流水线

将空中课堂/录播课程的原始 session 笔记与 PPT/PDF 复习资料整合，产出带幻灯片截图、法条原文、案例分析和答题框架的完整期末复习包。

## Codex 兼容说明

- 本技能为 Codex 兼容副本，目标是在不影响 Qoder 原版的前提下做最小兼容改造。
- Step 5 / Step 6 中涉及的并行代理仅在**用户明确要求**且当前环境提供相应能力时启用；否则由当前代理顺序执行同样的步骤。
- 如果当前课程只有 transcript、暂时没有可用的 PPT/PDF 复习资料，仍可先完成文本精修与结构化整理，图片相关步骤按实际情况跳过或延后。
- `session_notes/` 下按教学日合并后的笔记是当前主产物；`session_notes/_old/` 仅作为备份和回查参考，不作为后续改写目标。

## 前置条件

- Python 3.10+ 且已安装 `PyMuPDF`（`pip install pymupdf`）
- Windows 环境（PPTX 导出依赖 COM 自动化 `win32com`，macOS/Linux 需用 `libreoffice --headless` 替代）
- 课程目录遵循约定结构（见下方"目录结构约定"）

## 总体流程

```
Step 1  扫描课程结构 → 日期-Session-教师-主题映射
Step 2  合并同日 Session → 构建映射表 + 重编号为 01-N
Step 3  提取复习资料    → PDF 解密渲染 + PPTX 导出幻灯片（支持混合源）
Step 4  识别关键幻灯片  → 筛选框架图/对比表/流程图
Step 5  分批补充        → 嵌入图片 + 法条 + 案例 + 答题框架
Step 5.5 幻灯片对账     → 扫描引用、补缺失图片、修复错位
Step 6  图片验证        → 逐张 Read 比对 + 生成 img_verify 报告
Step 7  索引 + 清理    → index.md + 中间文件清理
```

## 目录结构约定

```
output/{课程名}/
├── sessions/                          # 原始 session 目录
│   ├── 01a-0302_1005-教师名/
│   │   └── transcript.md
│   ├── 01b-0302_1100-教师名/          # 同日第二场
│   │   └── transcript.md
│   └── ...
├── session_notes/                     # 处理后的笔记（本流水线产出）
│   ├── _old/                          # 原始笔记备份
│   ├── slides/                        # 幻灯片截图（最终引用目标）
│   ├── 01-course_summary.md
│   ├── 01-exam_focus.md
│   ├── ...
│   ├── index.md                       # 编号-日期-教师-主题对照表
│   └── img_verify_{课程}.md           # 图片验证报告
├── review_materials/                  # 用户提供的复习资料（PDF/PPTX，可混合）
│   ├── 01-竞争法-概述.pdf             # 加密 PDF
│   ├── topic02-银行.pptx              # PPTX
│   └── ...
├── pdf_slides/                        # PDF 渲染的幻灯片（中间产物）
│   └── {PDF文件名}/slide-{NN}.png
├── pptx_slides/                       # PPTX 导出的幻灯片（中间产物）
│   └── {PPTX文件名}/slide-{NN}.png
└── extracted_text/                    # PDF 提取的文本（中间产物）
    └── {PDF文件名}.md
```

目录命名模式：`{序号}{可选后缀}-{日期}_{时间}-教师名`，例如 `01a-0302_1005-樊健`。

**多教师课程**：同一课程可能有不同教师负责不同板块（如竞争法学：吴清卿负责 01-08 竞争法、袁波负责 09-14 反垄断法），每位教师有各自的复习资料。在 Step 1 的映射表中需建立「教师 → session 范围 → 复习资料」的对应关系。

## Step 1: 扫描课程结构

遍历 `sessions/` 子目录，解析每个目录名，提取：序号、后缀（a/b/c）、日期（MMDD）、时间（HHMM）、教师名。读取每个 session 的 `transcript.md` 前 30 行获取主题。

产出：`{课程名}_structure.md`，包含完整的日期-Session-教师-主题映射表，标记同日多 session 为合并候选。

**关键判断**：
- 同日期 + 不同时间后缀（a/b）→ 标记为"同日双场次，需合并"
- 同序号 + c 后缀 → 可能是重复机位，检查内容是否重复

## Step 2: 合并同日 Session 并重编号

1. 备份：`mkdir -p session_notes/_old && cp session_notes/*.md session_notes/_old/`
2. **构建映射表**：原始 session 编号未必从 01 开始（如信托法原编号为 20-41），需建立 `{旧编号} → {新编号}` 的完整映射。映射表按教学日期排序，为每个教学日分配连续编号 `01, 02, ..., N`
3. 同日双场次合并为一份笔记，用 `## 上半场` / `## 下半场` 分隔
4. 产出文件：`{NN}-course_summary.md` + `{NN}-exam_focus.md`
5. 删除旧编号文件（备份保留在 `_old/`）

合并时保留原笔记的核心内容结构，仅做拼接和格式统一。实质性补充在 Step 5 进行。

映射表格式示例见 [reference.md](reference.md) §3.3。

## Step 3: 提取复习资料

### 加密 PDF（PyMuPDF）

```python
import fitz
doc = fitz.open(pdf_path)
if doc.is_encrypted:
    doc.authenticate(password)  # 用户提供密码
# 文字提取
for page in doc:
    text += page.get_text()
# 页面渲染为图片
for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=200)
    pix.save(f"slides/slide-{i+1:02d}.png")
```

### PPTX（win32com，仅 Windows）

```python
import win32com.client, os
ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = 1
pres = ppt.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
for i, slide in enumerate(pres.Slides):
    slide.Export(os.path.abspath(f"slides/slide-{i+1:02d}.png"), "PNG", 1920, 1080)
pres.Close()
ppt.Quit()
```

**macOS/Linux 替代方案**：`libreoffice --headless --convert-to pdf file.pptx` 转为 PDF 后再用 PyMuPDF 渲染。

所有幻灯片图片先输出到工作目录的子文件夹（如 `pdf_slides/`、`pptx_slides/`），按章节分子目录存放。

### 判断 PDF 类型

提取文字后立即检查每页字符数。如果平均每页 < 100 字符，说明 PDF 是**幻灯片图片型**（非文字型），必须渲染为图片而非仅提取文字。

### 混合源处理

一门课程可能同时包含多种复习资料类型（如竞争法学：8 份加密 PDF + 5 份 PPTX + 1 份非加密 PDF）。处理策略：

1. 按类型分别处理：PDF 走 PyMuPDF，PPTX 走 win32com
2. 输出到**独立的中间目录**：`pdf_slides/` 和 `pptx_slides/`
3. 在 Step 4 识别关键幻灯片时统一从各中间目录筛选
4. 建立「教师 → 复习资料类型 → 中间目录」的对应关系

## Step 4: 识别关键幻灯片

逐张查看导出的幻灯片图片（用 Read 工具打开 PNG），筛选包含以下内容的幻灯片：

- **框架图**：制度结构、法律关系图、三方/四方主体关系
- **对比表**：新旧法条对比、不同制度对比、构成要件对比
- **流程图**：审批流程、交易流程、分析步骤
- **构成要件图**：法律行为的成立/生效要件

筛选后的图片复制到 `session_notes/slides/`，命名规范：`{topic前缀}-slide-{NN}.png`。

建立幻灯片映射表（在 Step 6 的 index.md 中体现），记录每张图对应的 session 编号和内容摘要。

## Step 5: 分批补充 session_notes

> **配合技能**：执行本步骤时，应同时加载 `session-notes-refinement` 技能，获取精修的深度标准（概念拓展式图片说明、法条查询模板、案例分析框架、答题模板生成的详细方法论和正反例）。pipeline 负责编排"做什么"，refinement 技能提供"怎么写好"的标准。

将课程按内容板块分为 2-4 批。若用户明确要求并行代理且环境支持，可按批次并行处理；否则由当前代理按 A → B → C → D 顺序完成，标准保持一致。

### 分批策略

| 批次 | 覆盖范围 | 示例 |
|------|---------|------|
| A | 总论/基础（01-04） | 概述、设立要件、基本原则 |
| B | 法律关系深化（05-08） | 权利义务、财产独立性 |
| C | 专题一（09-11） | 具体制度、监管规则 |
| D | 专题二（12-14） | 实务案例、前沿问题 |

### 分批执行要点

每一批执行时都应明确：负责的 session 列表 + 对应转写稿路径 + 幻灯片图片路径 + 复习资料文本。

对每份 session_notes 执行：

1. **嵌入幻灯片图片**：在相关知识点处插入 `![概念说明](slides/xxx.png)`
2. **图片说明采用"概念拓展模式"**：
   - 先抛出"为什么"的问题
   - 再借图片讲清制度逻辑和深层原理
   - 不是描述"PPT 展示了什么"，而是用图片辅助理解概念
3. **补充法条原文**：引用相关法律条文的关键条款
4. **添加案例分析框架**：典型案件的事实-争点-裁判-启示结构
5. **添加答题模板**：针对考试常见题型的分析步骤

### 并行代理时的防护规则（仅在显式启用时加入 prompt）

- **Copy slides first**：先将幻灯片图片拷贝到 `session_notes/slides/`，再写 `![]()` 引用。否则引用会指向不存在的文件。
- **Read before Edit**：每次 Edit 前必须 Read 目标文件，确保 `old_string` 与当前文件内容匹配。
- **Edit 失败后 re-Read 重试**：如果 Edit 报匹配失败（前次编辑改变了文件内容），重新 Read 后再构造正确的 `old_string`。
- **PowerShell for Chinese filenames**：Windows 下 `ls` 对中文文件名产生乱码，必须使用：
  ```powershell
  powershell -Command "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-ChildItem 'path'"
  ```
- **避免中文数字命名**：幻灯片文件命名用阿拉伯数字（如 `antitrust-12-slide-04.png`），不要用中文数字（如 `antitrust-一-slide-07.png`），避免编码问题。

### 图片说明示例

**差的写法**（描述模式）：
> PPT 展示了信用证的 12 步流程图。

**好的写法**（概念拓展模式）：
> 为什么信用证需要 12 步之多？核心在于买卖双方互不信任——买方怕付了钱拿不到货，卖方怕发了货收不到钱。银行作为信用中介介入后，形成了三条分离的线路：合同线、货物线、单据线。下图展示了这一机制如何运作：
> ![信用证 12 步流程](slides/topic02-slide-15.png)

## Step 5.5: 幻灯片对账（Post-enrichment Reconciliation）

本轮补充完成后，笔记中引用的图片可能并未全部拷贝到 `session_notes/slides/`。此步骤修复断裂引用。

1. 用 Grep 扫描所有 `session_notes/*.md` 中的 `![](...)` 引用
2. 逐个检查引用的图片文件是否存在于 `session_notes/slides/`
3. **缺失图片**：从源目录（`pdf_slides/` 或 `pptx_slides/`）查找并拷贝
4. **错位图片**：标记可疑引用，留待 Step 6 深度验证
5. **孤立图片**：`slides/` 中存在但未被任何笔记引用的文件，记录但不删除

产出：对账摘要（记录缺失/修复/可疑数量），作为 Step 6 的输入。

## Step 6: 图片验证

逐张检查所有图片引用。这是质量保证的核心步骤，不是简单的 checklist 项。若用户明确要求并行代理且环境支持，可为每门课分配独立验证代理；否则由当前代理顺序完成。

### 验证执行流程

1. 用 Grep 提取所有 `![...](...)` 引用
2. 对每个引用：
   - 用 **Read 工具打开 PNG 图片**，查看实际内容
   - 读取 markdown 中该图片前后的描述文字
   - 判断：图片内容是否与描述一致？
3. **不匹配时**：回到源幻灯片目录，逐张查看找到正确图片，替换引用
4. 产出 `session_notes/img_verify_{课程}.md` 验证报告

### 验证报告格式

```markdown
# {课程名} 图片验证报告

## 统计
- 总引用数：83
- 唯一图片数：74
- 验证通过：80
- 已修正：3
- 缺失：0

## 修正记录
| 文件 | 行号 | 原引用 | 修正后 | 原因 |
|------|------|--------|--------|------|
| 03-course_summary.md | L45 | topic02-slide-25.png | topic02-slide-29.png | slide-25 是银团贷款，slide-29 才是纸浆厂风险表 |

## 缺失图片
（无）
```

## Step 7: 索引生成与清理

### 索引文件

生成 `session_notes/index.md`，格式：

```markdown
# {课程名} 期末复习索引

| 编号 | 日期 | 教师 | 主题 | 复习资料 | 幻灯片数 |
|------|------|------|------|---------|---------|
| 01 | 03-02 | 樊健 | 信托法概述 | 无 | 0 |
| 02 | 03-09 | 吴清卿 | 竞争法概述 | 01-竞争法-概述.pdf | 8 |
```

### 清理中间文件

- 删除临时 PowerShell 脚本（如有）
- 清理编码异常产生的乱码文件名
- 确认 `_old/` 备份完整
- 中间产物（`pdf_slides/`、`pptx_slides/`、`extracted_text/`）保留供后续复查，不删除

## 已知陷阱与经验教训

- **PDF 以图片为主**：字符数偏少说明 PDF 是幻灯片图片型，必须渲染为图片而非仅提取文字
- **图片引用错位**：导出的幻灯片文件编号 ≠ PPT 内页码（封面页偏移），如 `slide-30.png` 对应 PPT 第 29 页。复制时务必用 Read 打开图片核对内容，不要靠文件名编号推断
- **加密 PDF**：部分复习资料有密码保护，需要用户提供密码后用 `doc.authenticate()` 解密
- **win32com 路径**：必须使用 `os.path.abspath()` 传入绝对路径，相对路径会报错；COM 对象必须显式 `pres.Close()` + `ppt.Quit()`，否则进程挂起
- **同日多机位**：`01c-0302_1005-第05机位` 这种通常是同一堂课的不同摄像角度，内容与 01a 重复，应跳过
- **图片验证必须逐张完成**：内容补充完成后必须逐张抽查图片引用准确性，曾发现 slide-25 被错误用于展示 slide-30 的内容
- **Windows 中文文件名乱码**：`ls`/`dir` 对中文文件名产生乱码，必须用 PowerShell + UTF8 编码
- **Python 路径非标准**：Python 可执行文件可能不在系统 PATH 中，需先 `where python` 或 `which python` 确认完整路径
- **Edit 的 old_string 失效**：同一文件多次 Edit 时，前一次编辑改变了文件内容，可能导致后续 `old_string` 匹配失败。每次 Edit 前必须 re-Read
- **引用不存在的图片**：补充内容时若先写 `![](slides/xxx.png)` 但未先将图片拷贝到 `slides/` 目录，会导致断裂引用。必须在 Step 5.5 做对账修复
- **中文数字命名导致编码问题**：`antitrust-一-slide-07.png` 等中文数字命名在某些工具中产生乱码，统一使用阿拉伯数字

## 验证清单

```
- [ ] sessions/ 目录已全部扫描，无遗漏
- [ ] 同日 session 已合并，编号连续无跳号
- [ ] 映射表已建立（旧编号 → 新编号），处理了任意起始编号
- [ ] 加密 PDF 已成功解密并渲染
- [ ] PPTX 幻灯片已全部导出
- [ ] 混合源（PDF + PPTX）已分别处理到独立中间目录
- [ ] 关键幻灯片已筛选并复制到 session_notes/slides/
- [ ] 每份 session_notes 已嵌入图片并补充内容
- [ ] 幻灯片对账完成（Step 5.5）：断裂引用已修复
- [ ] 图片验证步骤已完成（Step 6）：img_verify 报告已生成
- [ ] 所有图片引用已验证（文件存在 + 内容匹配）
- [ ] index.md 已生成且信息完整
- [ ] _old/ 备份完整
- [ ] 中间文件已清理（临时脚本、乱码文件名）
```

## 辅助文件

详细的 Python 脚本模板、并行代理 prompt 模板和完整示例，见 [reference.md](reference.md)。
