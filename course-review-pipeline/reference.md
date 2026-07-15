# reference.md — course-review-pipeline 辅助文件

本文档包含 Python 脚本模板、子代理 prompt 模板和完整示例，供执行时按需引用。

## 1. Python 脚本模板

### 1.1 批量解密并渲染 PDF 幻灯片

```python
"""
extract_pdfs.py — 批量解密 PDF 并渲染为幻灯片图片
用法: python extract_pdfs.py <pdf_dir> <output_dir> [password]
"""
import sys, os, fitz

def render_pdf(pdf_path, output_dir, password=None):
    doc = fitz.open(pdf_path)
    if doc.is_encrypted:
        if not password:
            print(f"SKIP (encrypted, no password): {pdf_path}")
            return 0
        if not doc.authenticate(password):
            print(f"FAIL (wrong password): {pdf_path}")
            return 0
    name = os.path.splitext(os.path.basename(pdf_path))[0]
    subdir = os.path.join(output_dir, name)
    os.makedirs(subdir, exist_ok=True)
    count = 0
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=200)
        pix.save(os.path.join(subdir, f"slide-{i+1:02d}.png"))
        count += 1
    doc.close()
    print(f"OK ({count} pages): {name}")
    return count

if __name__ == "__main__":
    pdf_dir = sys.argv[1]
    output_dir = sys.argv[2]
    password = sys.argv[3] if len(sys.argv) > 3 else None
    os.makedirs(output_dir, exist_ok=True)
    total = 0
    for f in sorted(os.listdir(pdf_dir)):
        if f.lower().endswith(".pdf"):
            total += render_pdf(os.path.join(pdf_dir, f), output_dir, password)
    print(f"\nTotal: {total} slides rendered")
```

### 1.2 批量导出 PPTX 幻灯片

```python
"""
export_pptx_slides.py — 批量导出 PPTX 幻灯片为 PNG (Windows only)
用法: python export_pptx_slides.py <pptx_dir> <output_dir>
"""
import sys, os, time

def export_slides(pptx_path, output_dir):
    import win32com.client
    abs_path = os.path.abspath(pptx_path)
    name = os.path.splitext(os.path.basename(pptx_path))[0]
    subdir = os.path.join(output_dir, name)
    os.makedirs(os.path.abspath(subdir), exist_ok=True)
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = 1
    try:
        pres = ppt.Presentations.Open(abs_path, WithWindow=False)
        count = 0
        for i, slide in enumerate(pres.Slides):
            out_path = os.path.abspath(os.path.join(subdir, f"slide-{i+1:02d}.png"))
            slide.Export(out_path, "PNG", 1920, 1080)
            count += 1
        pres.Close()
        print(f"OK ({count} slides): {name}")
        return count
    finally:
        ppt.Quit()
        time.sleep(1)

if __name__ == "__main__":
    pptx_dir = sys.argv[1]
    output_dir = sys.argv[2]
    os.makedirs(output_dir, exist_ok=True)
    total = 0
    for f in sorted(os.listdir(pptx_dir)):
        if f.lower().endswith((".pptx", ".ppt")):
            total += export_slides(os.path.join(pptx_dir, f), output_dir)
    print(f"\nTotal: {total} slides exported")
```

### 1.3 合并同日 Session 笔记

```python
"""
merge_notes.py — 按合并映射表合并同日 session_notes
用法: 在工作目录生成 merge_map.json 后运行此脚本
"""
import json, os, shutil

def merge_notes(notes_dir, merge_map, old_backup=True):
    if old_backup:
        backup_dir = os.path.join(notes_dir, "_old")
        os.makedirs(backup_dir, exist_ok=True)
        for f in os.listdir(notes_dir):
            if f.endswith(".md") and f != "index.md":
                shutil.copy2(os.path.join(notes_dir, f), backup_dir)

    for entry in merge_map:
        new_num = entry["num"]          # "01"
        label = entry["label"]          # "上半场" / "下半场" 或直接合并
        sources = entry["sources"]      # [{"num": "20", "label": "上半场"}, ...]

        for ftype in ["course_summary", "exam_focus"]:
            parts = []
            if entry.get("date"):
                parts.append(f"<!-- 日期: {entry['date']} -->\n")
            for src in sources:
                src_file = os.path.join(notes_dir, f"{src['num']}-{ftype}.md")
                if os.path.exists(src_file):
                    with open(src_file, "r", encoding="utf-8") as f:
                        content = f.read()
                    if len(sources) > 1 and src.get("label"):
                        parts.append(f"\n## {src['label']}\n\n{content}")
                    else:
                        parts.append(content)

            out_file = os.path.join(notes_dir, f"{new_num}-{ftype}.md")
            with open(out_file, "w", encoding="utf-8") as f:
                f.write("\n\n---\n\n".join(parts))

    # 清理旧文件
    old_nums = set()
    for entry in merge_map:
        for src in entry["sources"]:
            old_nums.add(src["num"])
    new_nums = {entry["num"] for entry in merge_map}
    for num in old_nums - new_nums:
        for ftype in ["course_summary", "exam_focus"]:
            old_file = os.path.join(notes_dir, f"{num}-{ftype}.md")
            if os.path.exists(old_file):
                os.remove(old_file)
```

## 2. 子代理 Prompt 模板

### 2.1 课程结构分析代理

```
快速分析 {课程名} 课程结构，建立日期-Session-教师-主题映射表。

## 需要检查的内容

### 1. Sessions 目录
路径：`{sessions_path}`
读取每个子目录名和 transcript.md 前 30 行。

### 2. Session Notes 目录
路径：`{notes_path}`
读取现有笔记文件的标题和结构。

### 3. 复习资料目录（如有）
路径：`{materials_path}`
列出 PDF/PPTX 文件名。

## 产出
写一份 `{课程名}_structure.md`，包含：
- 完整映射表（日期、session编号、教师、主题、时长）
- 同日多 session 标记（需合并的项）
- 复习资料与 session 的对应关系
- 建议的板块划分（用于后续分批处理）
```

### 2.2 笔记补充代理（通用模板）

```
你的任务是对 {课程名} session_notes 第 {起始}-{结束} 讲进行精细化补充。

## 项目导向
学习辅助项目。笔记应覆盖法条原文、构成要件、案例分析和答题框架。

## 你负责的文件

| Session | 日期 | 教师 | 主题 |
|---------|------|------|------|
{session_table}

## 文件路径
- 当前 session_notes：`{notes_path}`
- 转写稿：`{sessions_path}`
- 幻灯片图片：`{slides_path}`（如有）
- 复习资料文本：`{extracted_text_path}`（如有）

## 补充要求

### 图片嵌入
在相关知识点处插入 `![概念说明]({slides_rel_path})`。
图片说明采用"概念拓展模式"：
- 先抛出"为什么"的问题
- 再借图片讲清制度逻辑
- 不要描述"PPT 展示了什么"

### 法条原文
引用相关法律条文的关键条款号和内容摘要。

### 案例分析框架
用「事实 → 争点 → 裁判 → 启示」四段结构。

### 答题模板
针对考试常见题型给出分析步骤框架。

## 工作流程
1. 读取当前 session_notes 了解已有内容
2. 读取对应转写稿获取课堂讨论细节
3. 读取复习资料（如有）获取幻灯片和知识框架
4. 用 Edit 工具补充每份笔记
5. 完成后报告每份文件的修改摘要
```

### 2.3 图片引用验证代理

```
你的任务是逐一验证 {课程名} session_notes 中的图片引用是否准确匹配描述内容。

## 工作流程

### 1. 提取所有图片引用
用 Grep 搜索 `!\[` 开头的图片引用：
`{notes_path}`

### 2. 逐张验证
对每个图片引用：
1. 用 Read 工具打开图片文件
2. 查看图片实际内容
3. 对比 markdown 中该图片前后的描述文字
4. 判断：图片内容是否与描述一致？

### 3. 如果不匹配
在原始幻灯片目录中找到正确图片：
- PDF 幻灯片：`{pdf_slides_path}`
- PPTX 幻灯片：`{pptx_slides_path}`
修正图片引用和文件。

### 4. 产出验证报告
写 `{课程名}_img_verify.md`，列出：
- 验证通过的图片（简要列表）
- 修正过的图片（原引用 → 修正后）
- 缺失的图片（标记需要补充的位置）
```

## 3. 完整示例

### 3.1 映射表格式示例（信托法）

```markdown
# 信托法课程结构

## 基本信息
- 教师：樊健
- 教学日：14 天
- 有效 Session：22 个（含同日双场次）
- 复习资料：无（仅课堂笔记）

## 日期-Session 映射表

| 日期 | Session | 时间 | 主题 | 合并 |
|------|---------|------|------|------|
| 03-02 | 01a+01b | 10:05+11:00 | 信托法概述 + 案例分析 | 是 |
| 03-09 | 02 | 10:05 | 信托的设立 | 否 |
| 03-16 | 03a+03b | 10:05+11:00 | 信托财产 + 受益人地位 | 是 |
| 03-23 | 04 | 10:05 | 受托人职责 | 否 |
| 03-30 | 05a+05b | 10:05+11:00 | 法律关系辨析 | 是 |
| 04-13 | 06a+06b | 10:05+11:00 | 信托财产独立性 | 是 |
| 04-20 | 07a+07b | 10:05+11:00 | 受益人保护 | 是 |
| 04-27 | 08a+08b | 10:05+11:00 | 民事信托分论 | 是 |
| 05-09 | 09 | 10:05 | 强制执行排除 | 否 |
| 05-11 | 10 | 10:05 | 金融信托导论 | 否 |
| 05-18 | 11 | 10:05 | 投资者适当性 | 否 |
| 05-25 | 12 | 10:05 | 销售合规 | 否 |
| 06-01 | 13 | 10:05 | 托管专题 | 否 |
| 06-08 | 14 | 10:05 | 信托收据 | 否 |

## 板块划分（用于分批子代理）
- A（01-04）：总论 — 概述、设立、财产、职责
- B（05-08）：法律关系深化 — 辨析、独立性、保护、民事信托
- C（09-14）：金融信托专题 — 执行排除、资管、适当性、合规、托管、信托收据
```

### 3.2 index.md 格式示例

```markdown
# 国际金融法 期末复习索引

> 14 讲 | 教师：李仁 | 74 张 PPT 幻灯片截图 | 复习资料：7 份 PPTX

| 编号 | 日期 | 教师 | 主题 | 幻灯片 | 复习资料来源 |
|------|------|------|------|--------|-------------|
| 01 | 02-26 | 李仁 | 导论 | 3 | topic01-导论.pptx |
| 02 | 03-05 | 李仁 | 银行与国际银行业务 I&II | 12 | topic02-银行.pptx |
| 03 | 03-12 | 李仁 | 证券与国际证券业务 I&II | 15 | topic03-证券.pptx |
| 04 | 03-19 | 李仁 | 国际金融危机与影子银行 | 8 | topic04-危机.pptx |
| 05 | 03-26 | 李仁 | 保险与国际保险业务 | 10 | topic05-保险.pptx |
| 06 | 04-02 | 李仁 | 金融衍生品 | 11 | topic06-衍生品.pptx |
| 07 | 04-09 | 李仁 | 国际货币制度 | 8 | topic07-货币.pptx |
| 08 | 04-16 | 李仁 | 项目融资 | 7 | topic02-银行.pptx |
| 09-14 | ... | ... | ... | ... | ... |
```

### 3.3 合并映射表 JSON 格式

```json
[
  {
    "num": "01",
    "date": "03-02",
    "teacher": "樊健",
    "topic": "信托法概述 + 案例分析",
    "sources": [
      {"num": "23", "label": "上半场"},
      {"num": "24", "label": "下半场"}
    ]
  },
  {
    "num": "02",
    "date": "03-09",
    "teacher": "樊健",
    "topic": "信托的设立",
    "sources": [
      {"num": "25", "label": null}
    ]
  }
]
```

## 4. 依赖安装

```bash
# 必需
pip install pymupdf

# Windows PPTX 导出（通常已随 Office 安装）
pip install pywin32

# 验证
python -c "import fitz; print('PyMuPDF OK')"
python -c "import win32com.client; print('win32com OK')"
```

如果 Python 环境不确定，运行以下命令检测：
```bash
python --version && python -c "import sys; print(sys.executable)"
```

## 5. macOS/Linux 适配

当运行环境非 Windows 时，PPTX 导出需替换为 LibreOffice：

```bash
# 将 PPTX 转为 PDF
libreoffice --headless --convert-to pdf --outdir /tmp/pdf_out input.pptx

# 再用 PyMuPDF 渲染 PDF 为图片（复用 extract_pdfs.py）
python extract_pdfs.py /tmp/pdf_out /tmp/slides_out
```

加密 PDF 的处理在所有平台一致（PyMuPDF 跨平台）。

## 6. Windows 中文文件名处理

在 Windows 上，`ls`/`dir` 对中文文件名产生乱码。必须使用 PowerShell 并指定 UTF8 编码：

```powershell
# 列出目录（正确处理中文文件名）
powershell -Command "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-ChildItem 'path' | Select-Object Name"

# 搜索特定模式的文件
powershell -Command "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-ChildItem 'path' -Filter '*.png' | ForEach-Object { $_.Name }"
```

Python 脚本中也需要设置编码：

```python
import subprocess, sys
result = subprocess.run(
    ["powershell", "-Command",
     "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
     "Get-ChildItem 'path' -Filter '*.png' | ForEach-Object { $_.Name }"],
    capture_output=True, text=True, encoding="utf-8"
)
files = result.stdout.strip().split("\n")
```

## 7. 幻灯片对账脚本（Step 5.5）

```python
"""
reconcile_slides.py — 扫描 session_notes 中的图片引用，检查文件存在性，从源目录补缺
用法: python reconcile_slides.py <notes_dir> <slides_dir> <source_dirs...>
"""
import sys, os, re, shutil, glob

def extract_references(notes_dir):
    """提取所有 markdown 文件中的图片引用"""
    refs = []
    pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')
    for f in sorted(glob.glob(os.path.join(notes_dir, "*.md"))):
        if f.endswith("index.md") or f.endswith("img_verify"):
            continue
        with open(f, "r", encoding="utf-8") as fh:
            for lineno, line in enumerate(fh, 1):
                for match in pattern.finditer(line):
                    alt_text = match.group(1)
                    img_path = match.group(2)
                    refs.append({
                        "file": os.path.basename(f),
                        "lineno": lineno,
                        "alt": alt_text,
                        "path": img_path
                    })
    return refs

def reconcile(notes_dir, slides_dir, source_dirs):
    refs = extract_references(notes_dir)
    missing = []
    fixed = []
    ok = 0

    for ref in refs:
        full_path = os.path.join(notes_dir, ref["path"])
        if os.path.exists(full_path):
            ok += 1
            continue

        # 尝试从源目录查找
        filename = os.path.basename(ref["path"])
        found = False
        for src_dir in source_dirs:
            for root, dirs, files in os.walk(src_dir):
                if filename in files:
                    src_file = os.path.join(root, filename)
                    os.makedirs(os.path.dirname(full_path), exist_ok=True)
                    shutil.copy2(src_file, full_path)
                    fixed.append(ref)
                    found = True
                    break
            if found:
                break

        if not found:
            missing.append(ref)

    print(f"Reconciliation Report:")
    print(f"  Total references: {len(refs)}")
    print(f"  OK: {ok}")
    print(f"  Fixed (copied from source): {len(fixed)}")
    print(f"  Missing (not found in any source): {len(missing)}")

    if fixed:
        print(f"\nFixed:")
        for r in fixed:
            print(f"  {r['file']}:{r['lineno']} -> {r['path']}")

    if missing:
        print(f"\nMissing:")
        for r in missing:
            print(f"  {r['file']}:{r['lineno']} -> {r['path']} ({r['alt']})")

if __name__ == "__main__":
    notes_dir = sys.argv[1]
    slides_dir = sys.argv[2]
    source_dirs = sys.argv[3:]
    reconcile(notes_dir, slides_dir, source_dirs)
```

## 8. 图片验证子代理 Prompt 模板

```
你的任务是逐一验证 {课程名} session_notes 中所有图片引用的准确性。

## 工作流程

### 1. 提取所有图片引用
用 Grep 在 {notes_path} 下搜索 `!\[` 开头的图片引用，记录文件名、行号、图片路径。

### 2. 逐张验证
对每个图片引用：
1. 用 Read 工具打开图片文件（PNG）
2. 查看图片实际内容（标题、主题、关键信息）
3. 读取 markdown 文件中该图片前后的描述文字
4. 判断：图片内容是否与描述一致？

**关键判断标准**：
- 图片的主题/标题是否与上下文讨论的概念匹配？
- 如果图片是表格，表格的行列标题是否与描述的对比维度一致？
- 如果图片是流程图，步骤数量和内容是否与描述吻合？

### 3. 如果不匹配
在源幻灯片目录中查找正确图片：
- PDF 幻灯片：{pdf_slides_path}
- PPTX 幻灯片：{pptx_slides_path}

查找策略：
- 先看同一 PDF/PPTX 的相邻幻灯片（±5 页）
- 如果找不到，扩大搜索到其他 PDF/PPTX
- 找到后，用 Edit 修正 markdown 中的引用，并拷贝正确图片到 slides/

### 4. 产出验证报告
写 `{notes_path}/img_verify_{课程}.md`，包含：

```markdown
# {课程名} 图片验证报告

## 统计
- 总引用数：N
- 唯一图片数：N
- 验证通过：N
- 已修正：N
- 缺失：N

## 修正记录
| 文件 | 行号 | 原引用 | 修正后 | 原因 |
|------|------|--------|--------|------|
| ... | ... | ... | ... | ... |

## 缺失图片
| 文件 | 行号 | 引用路径 | 描述文字 |
|------|------|---------|---------|
| ... | ... | ... | ... |
```

## 重要提醒
- 每张图都要用 Read 打开看，不要靠文件名推断内容
- 导出编号和 PPT 内页码可能有偏移（封面页），注意核对
- 验证过程中发现的其他问题（如格式错误）也可顺手修复
```

## 9. 竞争法学混合源处理示例

竞争法学课程同时包含三种复习资料：

| 类型 | 数量 | 教师 | Session 范围 | 处理方式 |
|------|------|------|-------------|---------|
| 加密 PDF | 8 份 | 吴清卿 | 01-08（竞争法） | PyMuPDF 解密（密码 `<用户提供的密码>`）+ 渲染 |
| PPTX | 5 份 | 袁波 | 09-14（反垄断法） | win32com 导出 |
| 非加密 PDF | 1 份 | 袁波 | 09-14（反垄断法） | PyMuPDF 直接渲染 |

处理流程：
1. `competition_pdf_slides/` ← 8 份加密 PDF 渲染
2. `antitrust_pptx_slides/` ← 5 份 PPTX 导出
3. `antitrust_pdf_slides/` ← 1 份非加密 PDF 渲染
4. Step 4 筛选关键幻灯片时，从三个中间目录分别筛选
5. 拷贝到 `session_notes/slides/` 时统一命名：
   - 竞争法部分：`topic{NN}-slide-{NN}.png`
   - 反垄断法 PPTX：`antitrust-{session}-slide-{NN}.png`
   - 反垄断法 PDF：`antitrust-pdf-{session}-slide-{NN}.png`

子代理分批时按教师/板块划分：
- Agent A/B：sessions 01-08（竞争法），提供 `competition_pdf_slides/` 路径
- Agent C/D：sessions 09-14（反垄断法），提供 `antitrust_pptx_slides/` + `antitrust_pdf_slides/` 路径
