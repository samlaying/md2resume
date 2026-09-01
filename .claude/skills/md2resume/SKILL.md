---
name: md2resume
description: Markdown → 一页 DOCX 中文简历。主路径是参考模板克隆法：复刻已排版 DOCX 的配色、色块、字体层级、照片和分割线，只替换文字；内置成果蓝条、一页校准与受限字体净化。触发：用户提到简历、md2resume、生成 DOCX、克隆/套用参考简历排版、换字体/换日期/调一页、只要 DOCX 或 Courier 字体警告。
---

# md2resume — Markdown 转精美简历

把结构化 Markdown 渲染成带设计感的一页 DOCX 中文简历，视觉对齐 professional 风格：深蓝标题、蓝色段落下划线、经历卡片分隔线、蓝色 bullet、**成果蓝条高亮**。

项目踩坑与防护规则见同目录的 `../md2resume-pitfalls/SKILL.md`；遇到一页、字体、模板、锁文件或只交付 DOCX 的任务时，先加载它。

本 skill 位于 `.claude/skills/md2resume/`；代码与脚本在**仓库根目录**（运行命令时假设 cwd = 仓库根）。

## 两条渲染路径

### 路径 A：参考模板克隆法（主，高保真）⭐

拿一份**已经排好版的 docx 当骨架**，克隆其段落样式（配色 / 字体 / 色块 / 证件照 / 页边距），只替换文字内容 → 像素级复刻原版质感。适合"照着某份参考简历的风格出新版"。

- 脚本：`scripts/build_docx_from_reference.py`
- 样本骨架（脱敏，可直接用）：`assets/reference-template.docx`（假名 / 假联系方式 / 剪影照片）
- 想用自己的参考模板：加 `--reference 你的模板.docx`（**模板里的证件照会被原样保留进每一份产出**，所以用脱敏样本时记得换成自己的）

```bash
python3 scripts/build_docx_from_reference.py 简历.md \
  --reference assets/reference-template.docx \
  --output 简历.docx --scale 0.99
```

自动处理：
- **受限字体净化** ⭐：参考模板里宏样式带出来的 bare `Courier`（导 PDF 会报"受限字体无法嵌入/显示空白"）自动替换成可嵌入的 `Arial`。换名精确匹配 `"Courier"`，不会误伤可嵌入的 `"Courier New"`。
- **`--scale`**：整体等比缩放字号/段距压一页（`0.95`≈正文 9pt；溢出调小、留白多调大，区间 0.8–1.0）。
- **成果蓝条**：每段经历的**最后一条要点**自动渲染成浅蓝底 `#EBF3FB` + 左蓝竖条 `#2E75B6` + 深蓝字。公司无关（已通用化）。

**输入约定（此法专用，与路径 B 不同）**：
- `# 姓名｜职位`（全角 `｜`）；联系方式写**一行** `电话：… ｜ 微信：… ｜ 邮箱：…`（全角 `｜`，非 bullet）；可附 `求职方向：` / `个人定位：`
- 段落名固定：`## 教育背景` / `## 实习经历` / `## 核心能力`
- 经历：`### 公司｜部门·职位｜日期`（全角 `｜`）→ 空行 `**【项目】标题**`（整条加粗，无圆点）→ `- **要点**`；经历段数 ≥1 即可
- 成果行：放该段最后一条（普通 `- ` 或 `- > ` 都行）

### 路径 B：纯代码渲染法（免模板，更灵活）

不依赖参考 docx，纯代码套内置设计系统生成。输入更自由（半角 `|` 分字段、bullet 写联系方式、段落数不限）。适合从零生成、或没有参考模板时。

- 解析：`core/parser.py::parse_markdown` → `ResumeData`
- 渲染：`core/docx_renderer.py::render_docx(data, photo_path, scale=1.0, dense=True)`
- 命令行：`python3 build_b_resume.py examples/sample_zh.md --scale 0.90`

格式细则、加粗纪律、成果蓝条规则、一页适配参数详见 **references/markdown-conventions.md**。

## 何时用

- 给一段经历 / 旧简历 / 笔记，要生成精美简历
- 照某份参考 docx 的风格出新版 → **路径 A**
- 从零生成 / 调字号一页 / 没有参考模板 → **路径 B**
- 简历模板带有受限字体 Courier → **路径 A** 已自动净化

## 配色（两法一致）

主蓝 `#2E75B6` / 深蓝 `#1F4E79` / 灰 `#595959` / 正文 `#404040` / 成果条底 `#EBF3FB`。字体：中文微软雅黑 + 西文 Arial。

## 做简历的标准流程

1. 内容整理成 Markdown（路径 A：全角 `｜` + 一行联系方式 + 每段成果放最后一条；路径 B：见 references）
2. 路径 A：`scripts/build_docx_from_reference.py … --reference <模板> --scale 0.99`
3. 出图自检一页与样式：
   - `qlmanage -t -s 1700 -o /tmp/ql 简历.docx`（Quick Look 缩略图），或
   - 如需页数 QA，可临时用 LibreOffice 渲染到 `/tmp`，检查后删除临时 PDF；交付目录只保留 DOCX
4. 核对：① 一页 ② 每段成果蓝条在位 ③ 加粗克制（每条最多 1–2 个关键词）
5. 按需微调 `--scale`

## 注意

- 路径 A 的参考模板里**证件照会被保留**：用 `assets/reference-template.docx`（剪影占位）产出后，正式投递前换回带自己照片的模板，或事后在 Word 里替换图片。
- 两法解析器不同（A 自带 `parse_resume`、B 用 `core/parser.py`），Markdown 方言略有差异——照上面的"输入约定"写即可。
