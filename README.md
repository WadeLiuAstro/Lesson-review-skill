# Lesson Review Skills

面向课程复习资料整理的两个可复用 Agent Skills，遵循 [Agent Skills](https://agentskills.io/) 的目录约定。它们可以组合使用，也可以按任务独立调用。

本仓库主要服务于录播课程、课堂转写稿、PPT/PDF 复习资料的结构化整理，默认输出中文说明和 Markdown 笔记。

## 包含的 Skills

| Skill | 适用场景 | 主要产出 |
| --- | --- | --- |
| `course-review-pipeline` | 扫描课程目录、合并场次、提取 PPT/PDF、筛选关键幻灯片并执行整体对账 | 完整复习包、索引、验证报告 |
| `session-notes-refinement` | 将 Whisper 转写稿精修为系统笔记与考前速查材料 | `course_summary`、`exam_focus`、案例与答题框架 |

`course-review-pipeline` 负责端到端编排；`session-notes-refinement` 提供单份或多份笔记的内容深度标准。完整流程通常先运行 pipeline，再对重点 session 使用 refinement，最后回到 pipeline 完成图片和文件对账。

## 目录结构

```text
Lesson-review-skill/
├── course-review-pipeline/
│   ├── SKILL.md
│   └── reference.md
└── session-notes-refinement/
    ├── SKILL.md
    └── reference.md
```

每个目录中的 `SKILL.md` 定义触发条件和执行流程，`reference.md` 保存较长的模板、示例与辅助说明。

## 安装

### 通用 Agent Skills 目录

将仓库克隆到本地，再把需要的 Skill 目录复制到 Agent 支持的 Skill 搜索路径。不同客户端的具体路径可能不同，请以对应产品文档为准。

```bash
git clone https://github.com/WadeLiuAstro/Lesson-review-skill.git
mkdir -p ~/.agents/skills
cp -R Lesson-review-skill/course-review-pipeline ~/.agents/skills/
cp -R Lesson-review-skill/session-notes-refinement ~/.agents/skills/
```

### Codex（PowerShell）

```powershell
git clone https://github.com/WadeLiuAstro/Lesson-review-skill.git
New-Item -ItemType Directory -Force "$HOME/.agents/skills" | Out-Null
Copy-Item -Recurse Lesson-review-skill/course-review-pipeline "$HOME/.agents/skills/"
Copy-Item -Recurse Lesson-review-skill/session-notes-refinement "$HOME/.agents/skills/"
```

Codex 通常会自动发现 Skill；如果没有出现在技能选择器中，请重启 Codex。

## 使用示例

显式调用完整流水线：

```text
$course-review-pipeline
请扫描这个课程目录，合并同日场次，提取复习资料并生成完整复习包。
```

只精修已有转写稿：

```text
$session-notes-refinement
请把这些 Whisper 转写稿整理为 course_summary 和 exam_focus，并标记需要回看确认的内容。
```

调用前应提供输入目录、期望输出位置、课程范围，以及可用的复习资料。涉及加密文件时，应在本地会话中提供密码，不要把密码写入仓库或公开日志。

## 可选环境依赖

实际依赖取决于输入格式和当前 Agent 环境：

- Python 及 PyMuPDF：解析、解密和渲染 PDF。
- `python-pptx`：读取 PPTX 文本与结构。
- Windows、Microsoft PowerPoint 与 `pywin32`：通过 COM 高保真导出幻灯片。
- OCR、图片查看或法律数据库/MCP：用于图片识别、法条和案例核验。

这些依赖不会随 Skill 自动安装。缺少专用工具时，Skill 应降级处理并明确标注未核实内容。

## 安全与数据边界

- 批处理前备份原始课程目录，避免覆盖已有笔记。
- 不要提交课程文件、访问令牌、PDF 密码或其他敏感信息。
- 引用法条、案例和图片前进行来源核验；无法确认的内容标记为 `需回看确认`。
- 尊重课程资料、录音、幻灯片及第三方内容的版权和使用限制。

## 贡献

欢迎通过 Issue 或 Pull Request 提交兼容性修复、工作流改进和经过脱敏的通用示例。提交前请确认 Markdown 为 UTF-8，并移除课程材料、个人信息和凭据。

## 许可证

本仓库尚未添加开源许可证。在许可证明确前，内容默认保留所有权利；如需复制、修改或再分发，请先联系仓库所有者。
