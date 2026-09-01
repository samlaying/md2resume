# md2resume

**Markdown → 一页 DOCX 简历**

仓库里只有这一条 Skill：把结构化 Markdown 套进参考 DOCX 模板，输出一份可投递的一页 Word 简历。不启动 Web 服务，不生成 PDF/HTML/PNG。

Skill 路径：`.claude/skills/md2resume/SKILL.md`（踩坑见 `.claude/skills/md2resume-pitfalls/`）。

## 启动方式

```bash
pip install python-docx

python3 scripts/build_docx_from_reference.py examples/sample_reference.md \
  --reference assets/reference-template.docx \
  --output output/resume.docx \
  --scale 0.99
```

- 输入：Markdown 简历
- 输出：`output/resume.docx`（唯一交付物）
- `--scale`：整体缩放字号和段距。内容偏少可略增大（如 `0.99`），溢出到两页再降到 `0.90`

自己的参考模板（含证件照）用 `--reference 你的模板.docx` 传入，不要把真实模板提交进仓库。仓库里的 `assets/reference-template.docx` 是脱敏剪影样本。

## Markdown 约定（路径 A）

```markdown
# 姓名｜职位

电话：138-0000-0000 ｜ 微信：demo_wechat ｜ 邮箱：name@example.com
求职方向：AI 产品经理
个人定位：一句话定位

## 教育背景
**示例大学｜专业 · 本科｜2022.09–2026.06**
核心课程：……
成绩排名：专业前 10%

## 实习经历

### 公司｜部门 · 职位｜2025.01–2025.06

**【项目】项目标题**
- **要点标签**：说明与数据
- > 该段最后一条成果，会渲染成浅蓝底 + 左蓝条

## 核心能力
- **能力标签**：说明
```

完整可运行样本：`examples/sample_reference.md`。

要点：

- 姓名行用全角 `｜` 分隔职位
- 联系方式写**一行**，不要写成 bullet
- 段落名固定为 `教育背景` / `实习经历` / `核心能力`
- 每段经历的**最后一条**作为成果蓝条
- 每条只加粗 1–2 个关键词

## Tech Stack

- Word 输出：[python-docx](https://python-docx.readthedocs.io/)
- 参考模板克隆：`scripts/build_docx_from_reference.py`
