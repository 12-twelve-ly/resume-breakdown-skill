# resume-breakdown-skill

把一份简历拆解成 Word 版「面试问答手册」的 Claude Code Skill。

输出包含：

- 四个层级的目录：章节 → 条目 → 【预判问题与参考答案】 → Q1/Q2/Q3
- 原文关键词加粗、逐条拆解点
- 工作经历下的默认基础问题（附 9pt 小字示例回答）
- 问答空白区，方便粘贴自己整理的题库
- 页脚页码、导航窗格支持

## 安装

把本目录放到 Claude Code 的 skills 目录下：

```bash
# 用户级
cp -r resume-breakdown-skill ~/.claude/skills/resume-breakdown

# 或项目级
cp -r resume-breakdown-skill .claude/skills/resume-breakdown
```

## 使用

依赖：

```bash
pip install python-docx
pip install pywin32   # 可选，仅 Windows，用于自动更新目录
```

1. 复制模板：

   ```bash
   cp scripts/content_template.py content_候选人.py
   ```

2. 填写 `NAME`、`ROLE`、`OUT_PATH`、`SKILLS`、`WORK_EXPERIENCES`、`WORK_EXPERIENCE_DEFAULT_QA`、`PROJECTS`。

3. 生成文档：

   ```bash
   python scripts/build_docx.py content_候选人.py
   ```

4. 更新目录（Windows）：

   ```bash
   python scripts/update_toc.py 输出的文档.docx
   ```

   其他平台在 Word/WPS 中右键目录 →「更新域」。

## 目录结构

```
resume-breakdown-skill/
├── SKILL.md                  # Skill 指令
├── README.md
├── .gitignore
├── scripts/
│   ├── build_docx.py         # 文档生成器
│   ├── content_template.py   # 内容模块模板
│   └── update_toc.py         # 目录更新（Windows COM）
└── references/
    └── default-questions.md  # 通用基础问题库
```

## 隐私提醒

**不要提交任何真实简历或生成文档。** 本项目只包含通用脚本、模板和通用问题库：

- `content_<姓名>.py` 属于个人数据，请放在本仓库之外
- 生成的 `.docx`、简历 PDF、截图都不要上传
- `.gitignore` 已忽略 `*.docx`、`*.pdf`、`content_*.py`（模板除外）与临时目录

## License

未指定。使用或分发前请与作者确认。
