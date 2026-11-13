# -*- coding: utf-8 -*-
"""简历拆解内容模板：复制本文件为 content_姓名.py 后填写。"""

NAME = "姓名"
ROLE = "目标岗位"
OUT_PATH = r"输出文档的完整路径.docx"

SKILLS = [
    {
        "title": "技能 01｜技能名称",
        "original": "简历原文，**加粗关键词**。",
        "points": [
            ("拆解点：", "该点可能被追问的内容。"),
        ],
    },
]

# 单段工作经历可用 WORK_EXPERIENCE = {...}；多段使用 WORK_EXPERIENCES 列表
WORK_EXPERIENCES = [
    {
        "title": "公司名称｜岗位",
        "meta": "起止时间",
        "original_items": ["工作内容原文，**加粗关键词**。"],
    },
]

WORK_EXPERIENCE_DEFAULT_QA = [
    {
        "category": "（一）公司组织架构相关信息（必会）",
        "items": [
            ("问题", "示例：示例回答。"),
        ],
    },
]

PROJECTS = [
    {
        "title": "项目 01｜项目名称",
        "meta": "起止时间 ｜ 公司",
        "arch": "技术架构/使用工具（可留空）",
        "content_label": "项目背景",
        "background": "项目背景概述。",
        "duties": ["项目内容/职责原文，**加粗关键词**。"],
        "duties_label": "项目职责",
        "results": ["项目成效原文。"],
        "results_label": "项目成效",
        "points": [
            ("拆解点：", "该点可能被追问的内容。"),
        ],
    },
]
