---
name: resume-breakdown
description: 把一份简历（PDF/MD/DOCX）拆解成 Word 面试问答手册：真实标题目录、原文关键词加粗、拆解点、问题占位与默认基础问题。Use when the user wants to generate a resume breakdown or interview Q&A handbook from a resume file.
---

# 简历拆解 · 面试问答手册生成

把候选人的简历整理成一份可长期使用的 Word 面试手册：

- 技能特长：原文 + 关键词加粗 + 拆解点 + 空白问答区
- 工作经历：原文 + 岗位默认基础问题（小字示例回答）
- 项目经历：项目描述/职责/成果原文 + 拆解点 + 空白问答区
- 自动生成四级目录（章节 → 条目 → 预判问题 → Q1/Q2/Q3），支持导航窗格

## 输入

1. 简历文件：PDF / Markdown / DOCX
2. 目标岗位（决定默认问题库与拆解侧重）
3. 输出路径（默认放在简历同目录）

## 工作流

1. **读取简历**：PDF 优先用 PyMuPDF 渲染成图片后逐页阅读（中文 PDF 直接提取文字可能乱码）。
2. **提取内容**：技能特长、工作经历、项目经历。自我评价默认不拆解。
3. **创建内容模块**：复制 `scripts/content_template.py` 为 `content_<姓名>.py`，放在简历所在项目之外的独立目录。
4. **填写字段**：
   - `NAME` / `ROLE` / `OUT_PATH`
   - `SKILLS`：每条技能的原文（用 `**关键词**` 标记加粗）与拆解点
   - `WORK_EXPERIENCES`：多段工作经历列表，只保留原文
   - `WORK_EXPERIENCE_DEFAULT_QA`：岗位默认基础问题与示例回答
   - `PROJECTS`：项目概述/内容/成果/拆解点；`content_label`、`duties_label`、`results_label` 可按简历原文调整
5. **生成文档**：
   ```bash
   python scripts/build_docx.py "<content_<姓名>.py 的路径>"
   ```
6. **更新目录**（Windows，可选）：
   ```bash
   python scripts/update_toc.py "<输出的 docx 路径>"
   ```
   或在 Word/WPS 中右键目录 →「更新域」。
7. **验证**：打开文档检查目录层级、加粗、问答占位、页脚页码。

## 排版约定

- 标题 1：大章节（技能特长 / 工作经历 / 项目经历）
- 标题 2：技能条目、公司、项目
- 标题 3：`【预判问题与参考答案】`（进入目录）
- 标题 4：`Q1：…`（进入目录）
- `SectionLabel` 自定义样式：`【原文】`、`【拆解点】`，不进入目录
- 默认基础问题的示例回答：9pt 小字，不进入目录
- 简历没有技能栏时，`SKILLS = []`，生成器自动省略该章节并重排编号

## 隐私与上传规则（重要）

**绝对不要提交或上传以下内容：**

- 任何真实简历原文、姓名、手机号、邮箱、身份证号、住址
- `content_<姓名>.py` 等个人内容模块
- 生成的 docx（含个人信息）
- 简历 PDF / 截图 / 第三方思维导图（可能含版权水印）
- `_tmp_*` 临时目录、`__pycache__`

本仓库只包含通用脚本、模板和通用问题库。个人内容模块应放在仓库之外的独立目录。

## 依赖

- Python 3.9+
- `pip install python-docx`
- 可选：`pip install pywin32`（仅 Windows，用于自动更新目录）

## 限制

- 目录自动更新依赖 Word/WPS COM，仅 Windows 可用；其他平台请在办公软件中手动更新域。
- PDF 若为图片型或字体缺少编码映射，需要先渲染成图片再阅读。
