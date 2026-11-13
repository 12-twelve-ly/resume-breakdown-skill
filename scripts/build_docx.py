# -*- coding: utf-8 -*-
"""生成简历拆解完整版 docx：真实标题样式 + 目录域 + 问答空白区。

用法：
    python build_docx.py [内容模块名]
默认使用 content 模块；也可传入 content_xxx 生成其他简历的文档。
"""
import importlib
import importlib.util
import os
import sys

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def load_content(source: str) -> object:
    """按模块名或 .py 文件路径加载内容模块。"""
    if source.endswith(".py"):
        path = os.path.abspath(source)
        content_dir = os.path.dirname(path)
        if content_dir not in sys.path:
            sys.path.insert(0, content_dir)  # 允许内容模块导入同目录的公共数据
        spec = importlib.util.spec_from_file_location("resume_content", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    return importlib.import_module(source)


if len(sys.argv) < 2:
    sys.exit("用法: python build_docx.py <内容模块名或 content_xxx.py 路径>")

C = load_content(sys.argv[1])
OUT = getattr(C, "OUT_PATH", None)
if not OUT:
    sys.exit("内容模块必须在 OUT_PATH 中指定输出路径")
NAME = getattr(C, "NAME", "")
ROLE = getattr(C, "ROLE", "")
FONT = "微软雅黑"
DARK_BLUE = RGBColor(0x1F, 0x3A, 0x5F)
ACCENT = RGBColor(0x2E, 0x5C, 0x8A)
GRAY = RGBColor(0x88, 0x88, 0x88)

doc = Document()
doc.core_properties.title = "简历拆解 · 面试问答手册"

# ---------- 基础样式 ----------
normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
normal.paragraph_format.line_spacing = 1.35
normal.paragraph_format.space_after = Pt(4)

for sec in doc.sections:
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)


def style_heading(name, size, before, after, color):
    st = doc.styles[name]
    st.font.name = FONT
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = color
    st.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    st.paragraph_format.space_before = Pt(before)
    st.paragraph_format.space_after = Pt(after)
    st.paragraph_format.keep_with_next = True
    return st


style_heading("Heading 1", 15, 18, 10, DARK_BLUE)
style_heading("Heading 2", 13, 14, 8, DARK_BLUE)
style_heading("Heading 3", 11, 10, 4, ACCENT)
style_heading("Heading 4", 10.5, 8, 2, ACCENT)

# 非目录层级的分节标签：【原文】【拆解点】只保留视觉样式，不进入目录
section_style = doc.styles.add_style("SectionLabel", WD_STYLE_TYPE.PARAGRAPH)
section_style.base_style = doc.styles["Normal"]
section_style.font.name = FONT
section_style.font.size = Pt(11)
section_style.font.bold = True
section_style.font.color.rgb = ACCENT
section_style.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
section_style.paragraph_format.space_before = Pt(10)
section_style.paragraph_format.space_after = Pt(4)
section_style.paragraph_format.keep_with_next = True


def h3_plain(text):
    return doc.add_paragraph(text, style="SectionLabel")


def h1(text, page_break=True):
    p = doc.add_paragraph(text, style="Heading 1")
    if page_break:
        p.paragraph_format.page_break_before = True
    return p


def h2(text):
    return doc.add_paragraph(text, style="Heading 2")


def h3(text):
    return doc.add_paragraph(text, style="Heading 3")


def set_run(run, size=10.5, bold=False, color=None, italic=False):
    run.font.name = FONT
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def add_rich(p, text, size=10.5):
    """解析 **加粗** 标记"""
    for i, part in enumerate(text.split("**")):
        if part:
            r = p.add_run(part)
            set_run(r, size=size, bold=(i % 2 == 1))


def rich_para(text, indent=None, size=10.5):
    p = doc.add_paragraph()
    if indent is not None:
        p.paragraph_format.left_indent = Cm(indent)
    add_rich(p, text, size=size)
    return p


def label(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run(r, bold=True, color=ACCENT)
    return p


def shade(p, fill="EAF1FB"):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def add_field(paragraph, instr, placeholder=""):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr_el, sep, t, end):
        run._r.append(el)


def add_page_number_footer():
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("第 ")
    set_run(r1, size=9, color=GRAY)
    add_field(p, " PAGE ", "1")
    r2 = p.add_run(" 页")
    set_run(r2, size=9, color=GRAY)


def qa_placeholders(count):
    note = doc.add_paragraph()
    note.paragraph_format.space_after = Pt(4)
    r = note.add_run("（本区由你自行整理；复制下方 Q/A 段落即可扩充）")
    set_run(r, size=9, color=GRAY)
    for i in range(1, count + 1):
        q = doc.add_paragraph(style="Heading 4")  # Q 行进入目录层级
        q.paragraph_format.space_before = Pt(8)
        q.paragraph_format.space_after = Pt(3)
        q.paragraph_format.keep_with_next = True
        shade(q)
        r1 = q.add_run(f"Q{i}：")
        set_run(r1, bold=True, color=DARK_BLUE)
        r2 = q.add_run("（待填写问题）")
        set_run(r2, size=10, color=GRAY)
        a = doc.add_paragraph()
        a.paragraph_format.left_indent = Cm(0.4)
        a.paragraph_format.space_after = Pt(2)
        r3 = a.add_run("答：")
        set_run(r3, bold=True)
        r4 = a.add_run("（在此填写你的回答）")
        set_run(r4, size=10, color=GRAY)


def default_question(index, text):
    """默认基础问题：进入目录（标题 4）"""
    q = doc.add_paragraph(style="Heading 4")
    q.paragraph_format.space_before = Pt(8)
    q.paragraph_format.space_after = Pt(2)
    q.paragraph_format.keep_with_next = True
    shade(q)
    r1 = q.add_run(f"Q{index}：")
    set_run(r1, bold=True, color=DARK_BLUE)
    r2 = q.add_run(text)
    set_run(r2, bold=True, color=DARK_BLUE)
    return q


def example_answer(text):
    """示例回答：小字，不进入目录（文本自带「示例：」前缀）"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_run(r, size=9, color=RGBColor(0x59, 0x59, 0x59))
    return p


def category_label(text):
    """分组标签：小标题，不进入目录"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run(r, size=10.5, bold=True, color=DARK_BLUE)
    return p


def numbered_points(points):
    for i, (lead, rest) in enumerate(points, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.2)
        p.paragraph_format.space_after = Pt(3)
        r1 = p.add_run(f"{i}. ")
        set_run(r1)
        r2 = p.add_run(lead)
        set_run(r2, bold=True)
        r3 = p.add_run(rest)
        set_run(r3)


def numbered_rich(items, numbered=False):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(f"{i}. " if numbered else "• ")
        set_run(r)
        add_rich(p, item)


add_page_number_footer()

# ---------- 封面区 ----------
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("简历拆解 · 面试问答手册")
set_run(r, size=20, bold=True, color=DARK_BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(16)
r = p.add_run(f"{NAME} ｜ {ROLE}")
set_run(r, size=11, color=GRAY)

para = doc.add_paragraph()
r = para.add_run("使用说明：")
set_run(r, size=10, bold=True, color=ACCENT)
r = para.add_run("本文档对技能特长与项目经历逐条拆解（简历无技能栏时自动省略）；工作经历保留原文，并预置了该岗位的基础问题与示例回答（小字），可按需替换。"
                "加粗为关键词与得分点；各条目末尾的「预判问题与参考答案」为空白格式区，可直接粘贴你知识库中整理好的问题与答案。")
set_run(r, size=10, color=GRAY)

para = doc.add_paragraph()
r = para.add_run("目录：")
set_run(r, size=10, bold=True, color=ACCENT)
r = para.add_run("目录已展开到 Q 问题层级，可直接跳转到具体问题；新增或修改问题后，"
                "在 WPS 中右键目录 →「更新域」即可刷新；也可通过「视图 → 导航窗格」按标题定位。")
set_run(r, size=10, color=GRAY)

# ---------- 目录 ----------
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(10)
p.paragraph_format.space_after = Pt(6)
r = p.add_run("目  录")
set_run(r, size=16, bold=True, color=DARK_BLUE)

toc_p = doc.add_paragraph()
add_field(toc_p, 'TOC \\o "1-4" \\h \\z \\u', "（目录生成中：如未显示，右键此处 → 更新域）")

# ---------- 动态章节：技能 / 工作 / 项目 ----------
skills = getattr(C, "SKILLS", None) or []
work_list = getattr(C, "WORK_EXPERIENCES", None)
if not work_list and getattr(C, "WORK_EXPERIENCE", None):
    work_list = [C.WORK_EXPERIENCE]
work_list = work_list or []
projects = getattr(C, "PROJECTS", None) or []
qa_groups = getattr(C, "WORK_EXPERIENCE_DEFAULT_QA", None) or []

section_names = []
if skills:
    section_names.append("技能特长")
if work_list:
    section_names.append("工作经历")
if projects:
    section_names.append("项目经历")
CN_NUMS = "一二三四五六七八九十"
titles = {name: f"{CN_NUMS[i]}、{name}" for i, name in enumerate(section_names)}

if skills:
    h1(titles["技能特长"])
    for skill in skills:
        h2(skill["title"])
        h3_plain("【原文】")
        rich_para(skill["original"], indent=0.2)
        h3_plain("【拆解点】")
        numbered_points(skill["points"])
        h3("【预判问题与参考答案】")
        qa_placeholders(2)

if work_list:
    h1(titles["工作经历"])
    for we in work_list:
        h2(we["title"])
        p = doc.add_paragraph()
        r = p.add_run(we["meta"])
        set_run(r, size=10, color=GRAY)
        h3_plain("【原文】")
        numbered_rich(we["original_items"])

    # 工作经历下的默认基础问题
    if qa_groups:
        h3("【预判问题与参考答案】")
        qa_index = 0
        for group in qa_groups:
            category_label(group["category"])
            for q_text, example in group["items"]:
                qa_index += 1
                default_question(qa_index, q_text)
                example_answer(example)

if projects:
    h1(titles["项目经历"])
for proj in projects:
    h2(proj["title"])
    p = doc.add_paragraph()
    r = p.add_run(proj["meta"])
    set_run(r, size=10, color=GRAY)
    if proj.get("arch"):
        p = doc.add_paragraph()
        r = p.add_run(proj["arch"])
        set_run(r, size=10, color=ACCENT)

    h3_plain("【原文】")
    p = doc.add_paragraph()
    r = p.add_run(proj.get("content_label", "项目背景") + "：")
    set_run(r, bold=True)
    add_rich(p, proj["background"])

    if proj.get("duties"):
        label(proj.get("duties_label", "项目职责") + "：")
        numbered_rich(proj["duties"], numbered=True)

    if proj.get("results"):
        label(proj.get("results_label", "项目成效") + "：")
        numbered_rich(proj["results"], numbered=True)

    h3_plain("【拆解点】")
    numbered_points(proj["points"])
    h3("【预判问题与参考答案】")
    qa_placeholders(3)

doc.save(OUT)
print("saved:", OUT)
