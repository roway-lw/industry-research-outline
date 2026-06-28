#!/usr/bin/env python3
"""
Build a delivery package for an industry-brain research outline.

This script intentionally uses only the Python standard library. It creates:
- one template-driven DOCX research outline
- multiple XLSX attachment forms
- one ZIP delivery package with a manifest

Input:
  python build_delivery_package.py --spec delivery_spec.json --out-dir output
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import time
import uuid
import zipfile
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


DEFAULT_TEMPLATE_NAME = "临西轴承产业大脑调研大纲V1.1.docx"

DEFAULT_ATTACHMENTS = [
    ("附件1_调研对象信息登记表.xlsx", "调研对象信息登记表", ["单位名称", "单位类型", "联系人", "职务", "联系方式", "调研时间", "参与人员", "主要职责/业务范围", "与目标产业的关系", "后续跟进事项"]),
    ("附件2_访谈纪要记录表.xlsx", "访谈纪要记录表", ["调研对象", "访谈主题", "主要观点", "关键事实", "问题痛点", "数据资源", "现有系统平台", "涉及政策/项目/企业", "需补充资料", "待核实事项", "初步场景机会"]),
    ("附件3_资料收集台账.xlsx", "资料收集台账", ["资料名称", "资料类型", "提供单位", "是否已获取", "文件格式", "时间范围", "更新频率", "是否涉密或敏感", "可否用于项目建设", "缺失说明", "跟进负责人"]),
    ("附件4_企业调研信息表.xlsx", "企业调研信息表", ["企业名称", "所属产业链环节", "主营产品/服务", "产值/营收区间", "员工规模", "上下游客户", "核心技术/资质", "主要诉求", "政策需求", "融资需求", "用工需求", "市场拓展需求", "数字化基础", "是否愿意接入产业大脑服务"]),
    ("附件5_产业链环节核实表.xlsx", "产业链环节核实表", ["产业链环节", "本地是否具备", "代表企业", "优势程度", "短板环节", "外部依赖区域", "招商补链方向", "现场确认来源", "备注"]),
    ("附件6_数据资源盘点表.xlsx", "数据资源盘点表", ["数据名称", "数据来源单位", "业务系统", "数据内容", "数据格式", "更新频率", "时间跨度", "数据质量", "共享条件", "接入难点", "可支撑场景"]),
    ("附件7_产业大脑场景线索表.xlsx", "产业大脑场景线索表", ["场景名称", "提出单位", "当前业务痛点", "涉及部门/企业", "现有系统", "所需数据", "数据来源", "建设价值", "实施难度", "优先级", "后续验证事项"]),
    ("附件8_问题假设验证清单.xlsx", "问题假设验证清单", ["假设问题", "来源依据", "需验证对象", "验证方式", "现场反馈", "是否成立", "对后续方案影响"]),
]

DEFAULT_ORGANIZATION_ATTACHMENTS = [
    ("附件1_单位基础信息记录表.xlsx", "单位基础信息记录表", ["单位名称", "单位类型", "联系人", "职务", "联系方式", "调研时间", "参与人员", "主营业务/核心职能", "服务对象", "与产业大脑关系"]),
    ("附件2_访谈纪要记录表.xlsx", "访谈纪要记录表", ["调研对象", "访谈主题", "主要观点", "关键事实", "痛点需求", "数据资源", "现有系统平台", "合作意向", "需补充资料", "后续跟进事项"]),
    ("附件3_合作机会梳理表.xlsx", "合作机会梳理表", ["合作方向", "合作内容", "涉及业务/服务", "产业大脑承接模块", "对单位价值", "对产业大脑价值", "落地难度", "优先级", "下一步动作"]),
    ("附件4_数据资源与系统对接盘点表.xlsx", "数据资源与系统对接盘点表", ["数据/系统名称", "所属业务", "数据内容", "数据格式", "更新频率", "系统接口情况", "共享条件", "数据安全要求", "可支撑场景", "待确认事项"]),
    ("附件5_资料收集与行动清单.xlsx", "资料收集与行动清单", ["事项类型", "资料/行动名称", "提供方", "用途", "当前状态", "负责人", "计划完成时间", "备注"]),
]


class BuildError(RuntimeError):
    pass


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def sanitize_filename(name: str, fallback: str = "output") -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip().strip(".")
    return name or fallback


def read_json(path: Path) -> Dict[str, Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise BuildError(f"Spec file not found: {path}") from exc
    except json.JSONDecodeError as exc:
        raise BuildError(f"Spec is not valid JSON: {path} ({exc})") from exc


def require_text(spec: Dict[str, Any], key: str) -> str:
    value = str(spec.get(key, "")).strip()
    if not value:
        raise BuildError(f"Missing required field: {key}")
    return value


def default_template_path(script_path: Path) -> Path:
    return script_path.resolve().parents[1] / "assets" / DEFAULT_TEMPLATE_NAME


StyleSample = Dict[str, Optional[ET.Element]]
StyleSamples = Dict[str, StyleSample]


def paragraph_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(f".//{qn(W_NS, 't')}")).strip()


def paragraph_style_id(p: ET.Element) -> str:
    pstyle = p.find(f"{qn(W_NS, 'pPr')}/{qn(W_NS, 'pStyle')}")
    return pstyle.get(qn(W_NS, "val"), "") if pstyle is not None else ""


def clone_paragraph_format(p: ET.Element) -> StyleSample:
    ppr = p.find(qn(W_NS, "pPr"))
    first_run = p.find(qn(W_NS, "r"))
    rpr = first_run.find(qn(W_NS, "rPr")) if first_run is not None else None
    return {
        "ppr": copy.deepcopy(ppr) if ppr is not None else None,
        "rpr": copy.deepcopy(rpr) if rpr is not None else None,
    }


def ensure_first_line_indent(ppr: Optional[ET.Element], first_line_chars: int = 200) -> ET.Element:
    ppr_el = copy.deepcopy(ppr) if ppr is not None else ET.Element(qn(W_NS, "pPr"))
    ind = ppr_el.find(qn(W_NS, "ind"))
    if ind is None:
        ind = ET.SubElement(ppr_el, qn(W_NS, "ind"))
    ind.set(qn(W_NS, "firstLineChars"), str(first_line_chars))
    return ppr_el


def extract_template_style_samples(document_root: ET.Element) -> StyleSamples:
    body = document_root.find(qn(W_NS, "body"))
    if body is None:
        raise BuildError("Template document.xml missing body")

    samples: StyleSamples = {}
    for p in body.findall(qn(W_NS, "p")):
        text = paragraph_text(p)
        if not text:
            continue

        style_id = paragraph_style_id(p)
        if "title" not in samples:
            samples["title"] = clone_paragraph_format(p)
        if style_id == "2" and "heading1" not in samples:
            samples["heading1"] = clone_paragraph_format(p)
        elif style_id == "3" and "heading2" not in samples:
            samples["heading2"] = clone_paragraph_format(p)

        ind = p.find(f"{qn(W_NS, 'pPr')}/{qn(W_NS, 'ind')}")
        if ind is not None and ind.get(qn(W_NS, "firstLineChars")) == "200" and "body" not in samples:
            samples["body"] = clone_paragraph_format(p)

    if "body" not in samples:
        for p in body.findall(qn(W_NS, "p")):
            if paragraph_text(p) and paragraph_style_id(p) not in {"2", "3"}:
                samples["body"] = clone_paragraph_format(p)
                break

    if "title" not in samples or "heading1" not in samples or "heading2" not in samples or "body" not in samples:
        missing = ", ".join(k for k in ("title", "heading1", "heading2", "body") if k not in samples)
        raise BuildError(f"Template missing required style samples: {missing}")

    samples["heading3"] = samples.get("heading3") or samples["heading2"]
    samples["question_group"] = samples["body"]
    samples["question"] = samples["body"]
    samples["paragraph"] = samples["body"]
    return samples


def make_paragraph(
    text: str,
    style_id: Optional[str] = None,
    ppr: Optional[ET.Element] = None,
    rpr: Optional[ET.Element] = None,
    first_line_chars: Optional[int] = None,
) -> ET.Element:
    p = ET.Element(qn(W_NS, "p"))
    if ppr is not None:
        p.append(ensure_first_line_indent(ppr, first_line_chars) if first_line_chars is not None else copy.deepcopy(ppr))
    elif style_id or first_line_chars is not None:
        ppr_el = ET.SubElement(p, qn(W_NS, "pPr"))
        if style_id:
            pstyle = ET.SubElement(ppr_el, qn(W_NS, "pStyle"))
            pstyle.set(qn(W_NS, "val"), style_id)
        if first_line_chars is not None:
            ind = ET.SubElement(ppr_el, qn(W_NS, "ind"))
            ind.set(qn(W_NS, "firstLineChars"), str(first_line_chars))

    r = ET.SubElement(p, qn(W_NS, "r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = ET.SubElement(r, qn(W_NS, "t"))
    t.set(qn(XML_NS, "space"), "preserve")
    t.text = text
    return p


def make_blank_paragraph() -> ET.Element:
    return ET.Element(qn(W_NS, "p"))


def sample_for(samples: StyleSamples, key: str) -> Tuple[Optional[ET.Element], Optional[ET.Element]]:
    sample = samples.get(key) or samples["paragraph"]
    return sample.get("ppr"), sample.get("rpr")


def block_to_paragraphs(block: Dict[str, Any], samples: StyleSamples) -> List[ET.Element]:
    typ = str(block.get("type", "paragraph")).strip().lower()
    text = str(block.get("text", "")).strip()

    if typ == "blank":
        return [make_blank_paragraph()]
    if typ == "title":
        ppr, rpr = sample_for(samples, "title")
        return [make_paragraph(text, ppr=ppr, rpr=rpr)]
    if typ == "heading1":
        ppr, rpr = sample_for(samples, "heading1")
        return [make_paragraph(text, ppr=ppr, rpr=rpr)]
    if typ == "heading2":
        ppr, rpr = sample_for(samples, "heading2")
        return [make_paragraph(text, ppr=ppr, rpr=rpr)]
    if typ == "heading3":
        ppr, rpr = sample_for(samples, "heading3")
        return [make_paragraph(text, ppr=ppr, rpr=rpr)]
    if typ == "question_group":
        ppr, rpr = sample_for(samples, "question_group")
        return [make_paragraph(text, ppr=ppr, rpr=rpr, first_line_chars=200)]
    if typ == "question":
        ppr, rpr = sample_for(samples, "question")
        return [make_paragraph(text, ppr=ppr, rpr=rpr, first_line_chars=200)]
    ppr, rpr = sample_for(samples, "paragraph")
    return [make_paragraph(text, ppr=ppr, rpr=rpr, first_line_chars=200)]


def default_outline_blocks(spec: Dict[str, Any]) -> List[Dict[str, str]]:
    region = require_text(spec, "region")
    industry = require_text(spec, "industry")
    title = spec.get("title") or f"{region}{industry}产业大脑调研大纲"
    return [
        {"type": "title", "text": title},
        {"type": "blank", "text": ""},
        {"type": "heading1", "text": "一、调研背景及目标"},
        {"type": "heading2", "text": "（一）调研背景"},
        {"type": "paragraph", "text": f"围绕{region}{industry}产业大脑建设前期工作，需通过资料分析和线下调研掌握本地产业基础、产业链结构、企业诉求、数据资源和场景机会。"},
        {"type": "heading2", "text": "（二）调研目标"},
        {"type": "paragraph", "text": "1.现状诊断：梳理产业基础、产业链关键环节、重点企业、政策资源和公共服务能力。"},
        {"type": "paragraph", "text": "2.需求挖掘：识别政府、园区、企业和服务机构对产业大脑的核心诉求。"},
        {"type": "paragraph", "text": "3.方案锚点：为后续产业大脑功能模块、数据资源目录和业务场景设计提供依据。"},
        {"type": "heading1", "text": "二、调研范围与重点"},
        {"type": "paragraph", "text": "重点围绕产业链结构、企业分布、公共服务平台、数据资源、业务系统、政策服务和产业大脑应用场景开展调研。"},
        {"type": "heading1", "text": "三、产业初步认知"},
        {"type": "paragraph", "text": "本章节应结合用户资料和公开信息形成初步判断，并将不确定内容转化为待现场核实问题。"},
        {"type": "heading1", "text": "四、调研对象及问题清单"},
        {"type": "heading2", "text": "（一）政府部门"},
        {"type": "paragraph", "text": "1.产业整体情况"},
        {"type": "paragraph", "text": "（1）本地产业企业数量、规上企业数量、产值、税收、就业、出口等核心数据情况如何？"},
        {"type": "heading2", "text": "（二）园区运营方与行业协会/产业联盟"},
        {"type": "paragraph", "text": "（1）本地产业链协同、公共服务平台、协会活动和企业共性诉求情况如何？"},
        {"type": "heading2", "text": "（三）重点企业"},
        {"type": "paragraph", "text": "（1）企业主营产品、上下游配套、数字化基础、融资人才市场诉求和产业大脑使用意愿如何？"},
        {"type": "heading2", "text": "（四）行业服务平台"},
        {"type": "paragraph", "text": "（1）检测认证、金融、物流、人才等服务机构可提供哪些能力和数据？"},
        {"type": "heading1", "text": "五、产业大脑场景挖掘重点"},
        {"type": "paragraph", "text": "重点挖掘产业监测、企业画像、产业链图谱、精准招商、政策服务、供需对接、数据资源盘点等场景。"},
        {"type": "heading1", "text": "六、资料收集清单"},
        {"type": "paragraph", "text": "围绕政策规划、企业名录、产业运行、项目招商、园区载体、业务系统和数据台账收集资料。"},
        {"type": "heading1", "text": "七、待现场核实问题"},
        {"type": "paragraph", "text": "将公开资料和前期资料中无法确认的内容列为现场核实问题，避免将调研假设直接作为结论。"},
    ]


def normalized_research_type(spec: Dict[str, Any]) -> str:
    value = str(spec.get("research_type") or spec.get("type") or "industry").strip().lower()
    if value in {"organization", "unit", "enterprise", "ecosystem", "institution"}:
        return "organization"
    return "industry"


def normalized_unit_category(spec: Dict[str, Any]) -> str:
    value = str(spec.get("unit_category") or spec.get("organization_category") or spec.get("unit_type") or "").strip().lower()
    enterprise_keys = ["企业", "龙头", "链主", "制造", "工厂", "company", "enterprise", "manufacturer"]
    return "enterprise" if any(key in value for key in enterprise_keys) else "ecosystem"


def require_organization_name(spec: Dict[str, Any]) -> str:
    value = str(spec.get("organization_name") or spec.get("unit_name") or spec.get("enterprise_name") or "").strip()
    if not value:
        raise BuildError("Missing required field for organization research: organization_name")
    return value


def organization_context(spec: Dict[str, Any]) -> Dict[str, str]:
    organization_name = require_organization_name(spec)
    unit_type = str(spec.get("unit_type") or spec.get("organization_type") or "待确认单位类型").strip()
    region = str(spec.get("region") or "").strip()
    industry = str(spec.get("industry") or spec.get("target_industry") or "目标产业").strip()
    project_name = str(spec.get("project_name") or f"{industry}产业大脑").strip()
    cooperation_goal = str(spec.get("cooperation_goal") or spec.get("research_goal") or "围绕产业大脑建设探索合作点").strip()
    return {
        "organization_name": organization_name,
        "unit_type": unit_type,
        "region": region,
        "industry": industry,
        "project_name": project_name,
        "cooperation_goal": cooperation_goal,
    }


def industry_chain_label(industry: str) -> str:
    industry = industry.strip()
    if industry.endswith("产业"):
        return f"{industry[:-2]}产业链"
    return f"{industry}产业链"


def default_enterprise_outline_blocks(spec: Dict[str, Any]) -> List[Dict[str, str]]:
    ctx = organization_context(spec)
    name = ctx["organization_name"]
    industry = ctx["industry"]
    chain_label = industry_chain_label(industry)
    project = ctx["project_name"]
    title = spec.get("title") or f"{name}产业大脑合作调研大纲"
    return [
        {"type": "title", "text": title},
        {"type": "blank", "text": ""},
        {"type": "heading1", "text": "一、交流背景及目标"},
        {"type": "heading2", "text": "（一）交流背景"},
        {"type": "paragraph", "text": f"围绕{project}建设，需要与{name}开展专题交流，重点了解企业经营现状、产业链位置、核心痛点、数字化基础和可参与的产业大脑合作场景。"},
        {"type": "heading2", "text": "（二）交流目标"},
        {"type": "paragraph", "text": f"1.识别{name}在{chain_label}中的角色，以及其对采购、销售、融资、用工、数字化和政策服务的真实需求。"},
        {"type": "paragraph", "text": "2.判断企业可作为产业大脑服务对象、场景试点对象、标杆案例、链主牵引方或数据需求来源的可能性。"},
        {"type": "paragraph", "text": "3.形成后续合作机会清单、需补充资料清单和现场核实问题。"},
        {"type": "heading1", "text": "二、企业基本情况"},
        {"type": "paragraph", "text": "重点了解企业名称、主营产品、产值规模、员工规模、所属产业链环节、主要客户和供应商、区域市场地位等基础情况。"},
        {"type": "heading1", "text": "三、企业经营与业务现状"},
        {"type": "heading2", "text": "（一）采购与供应链"},
        {"type": "paragraph", "text": "了解原材料或关键配套采购来源、采购渠道、价格波动影响、供应稳定性、是否存在集采或供应链协同需求。"},
        {"type": "heading2", "text": "（二）生产与交付"},
        {"type": "paragraph", "text": "了解订单组织、产能利用、设备能力、质量管理、能耗成本、交付周期和生产过程中的主要约束。"},
        {"type": "heading2", "text": "（三）销售与市场"},
        {"type": "paragraph", "text": "了解主要销售渠道、客户结构、获客方式、品牌使用、区域市场拓展和线上销售基础。"},
        {"type": "heading2", "text": "（四）数字化与数据基础"},
        {"type": "paragraph", "text": "了解企业是否使用ERP、MES、PLM、SCM、财务系统、电商平台或设备数据采集工具，以及可用于企业画像和场景服务的数据基础。"},
        {"type": "heading1", "text": "四、核心痛点与需求"},
        {"type": "paragraph", "text": "重点识别企业在降本、拓客、融资、用工、检测认证、物流、政策获取、数字化改造等方面的实际需求，并区分短期刚需和长期提升需求。"},
        {"type": "heading1", "text": "五、产业大脑合作方向"},
        {"type": "paragraph", "text": "围绕集采集销、供需对接、供应链金融、企业画像、政策服务、数字化改造、能耗优化、标杆示范和数据接入等方向，判断可落地的合作场景。"},
        {"type": "heading1", "text": "六、重点交流问题清单"},
        {"type": "paragraph", "text": "（1）企业当前最希望产业大脑帮助解决的三个问题是什么？"},
        {"type": "paragraph", "text": "（2）企业是否愿意参与区域统一品牌、集采集销、供需对接或标杆示范？"},
        {"type": "paragraph", "text": "（3）企业有哪些业务系统和数据可以支撑企业画像、供需匹配、金融服务或政策服务？"},
        {"type": "paragraph", "text": "（4）企业对数据共享、系统对接、服务收费和平台使用有哪些顾虑？"},
        {"type": "heading1", "text": "七、需补充资料清单"},
        {"type": "paragraph", "text": "建议收集企业介绍、产品目录、上下游合作情况、主要业务系统说明、数字化建设资料、政策项目和荣誉资质、可参与产业大脑合作的业务清单。"},
    ]


def default_ecosystem_outline_blocks(spec: Dict[str, Any]) -> List[Dict[str, str]]:
    ctx = organization_context(spec)
    name = ctx["organization_name"]
    unit_type = ctx["unit_type"]
    industry = ctx["industry"]
    project = ctx["project_name"]
    title = spec.get("title") or f"{name}产业大脑生态合作调研大纲"
    return [
        {"type": "title", "text": title},
        {"type": "blank", "text": ""},
        {"type": "heading1", "text": "一、交流背景及目标"},
        {"type": "heading2", "text": "（一）交流背景"},
        {"type": "paragraph", "text": f"围绕{project}建设，需要与{name}开展专题交流，重点了解该单位在{industry}领域的资源、数据、服务、平台和运营能力，判断双方可共建的合作场景。"},
        {"type": "heading2", "text": "（二）交流目标"},
        {"type": "paragraph", "text": f"1.识别{name}作为{unit_type}可为产业大脑提供的数据资源、服务能力、企业资源或运营支撑。"},
        {"type": "paragraph", "text": "2.梳理双方在数据共享、系统对接、服务专区、活动运营、资源导入和品牌推广方面的合作机会。"},
        {"type": "paragraph", "text": "3.形成可试点合作事项、资料清单和后续对接动作。"},
        {"type": "heading1", "text": "二、机构基本情况"},
        {"type": "paragraph", "text": "重点了解机构性质、主管或运营单位、服务对象、覆盖区域、行业范围、核心职能和主要业务成果。"},
        {"type": "heading1", "text": "三、机构资源与服务能力"},
        {"type": "heading2", "text": "（一）数据资源"},
        {"type": "paragraph", "text": "了解机构掌握的企业、会员、人才、活动、培训、检测、金融、物流、政策、项目或行业资讯等数据资源及其更新机制。"},
        {"type": "heading2", "text": "（二）服务能力"},
        {"type": "paragraph", "text": "了解机构面向产业企业提供的专业服务、服务产品、服务流程、收费模式、覆盖企业数量和典型服务案例。"},
        {"type": "heading2", "text": "（三）平台系统能力"},
        {"type": "paragraph", "text": "了解机构是否已有网站、业务系统、数据平台、小程序或接口能力，以及与产业大脑进行数据或服务对接的技术条件。"},
        {"type": "heading1", "text": "四、现有业务与运营模式"},
        {"type": "paragraph", "text": "重点了解机构当前业务流程、资源触达方式、会员或客户运营方式、活动组织机制、合作渠道和区域落地能力。"},
        {"type": "heading1", "text": "五、产业大脑合作方向"},
        {"type": "paragraph", "text": "围绕数据资源共享、系统接口对接、服务专区共建、活动联合运营、专家库或知识库共建、企业资源导入、联合推广和生态服务闭环等方向，判断可落地的合作点。"},
        {"type": "heading1", "text": "六、重点交流问题清单"},
        {"type": "paragraph", "text": "（1）该单位当前有哪些资源、数据或服务可与产业大脑形成互补？"},
        {"type": "paragraph", "text": "（2）是否具备数据共享、接口对接、服务入口嵌入或联合运营的条件？"},
        {"type": "paragraph", "text": "（3）双方共建专区、共办活动、共建数据库或共同服务企业的可行路径是什么？"},
        {"type": "paragraph", "text": "（4）合作推进中涉及的数据安全、权责边界、运营分工和收益模式需要如何约定？"},
        {"type": "heading1", "text": "七、需补充资料清单"},
        {"type": "paragraph", "text": "建议收集机构介绍、服务清单、数据目录、系统说明、活动计划、会员或服务对象情况、典型案例和可开放合作资源清单。"},
    ]


def default_organization_outline_blocks(spec: Dict[str, Any]) -> List[Dict[str, str]]:
    if normalized_unit_category(spec) == "enterprise":
        return default_enterprise_outline_blocks(spec)
    return default_ecosystem_outline_blocks(spec)


def build_docx(template: Path, out_docx: Path, spec: Dict[str, Any]) -> None:
    if not template.exists():
        raise BuildError(f"Template DOCX not found: {template}")
    if not zipfile.is_zipfile(template):
        raise BuildError(f"Template is not a valid DOCX zip: {template}")

    if spec.get("outline_blocks"):
        blocks = spec["outline_blocks"]
    elif normalized_research_type(spec) == "organization":
        blocks = default_organization_outline_blocks(spec)
    else:
        blocks = default_outline_blocks(spec)
    if not isinstance(blocks, list) or not blocks:
        raise BuildError("outline_blocks must be a non-empty list")

    with tempfile.TemporaryDirectory(prefix="docx_build_") as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(template, "r") as zin:
            zin.extractall(tmp_path)

        document_xml = tmp_path / "word" / "document.xml"
        if not document_xml.exists():
            raise BuildError("Template DOCX missing word/document.xml")

        root = ET.parse(document_xml).getroot()
        body = root.find(qn(W_NS, "body"))
        if body is None:
            raise BuildError("Template document.xml missing body")

        style_samples = extract_template_style_samples(root)
        sect_pr = body.find(qn(W_NS, "sectPr"))
        sect_pr_copy = copy.deepcopy(sect_pr) if sect_pr is not None else None

        for child in list(body):
            body.remove(child)

        for block in blocks:
            if not isinstance(block, dict):
                raise BuildError("Each outline block must be an object")
            for para in block_to_paragraphs(block, style_samples):
                body.append(para)

        if sect_pr_copy is not None:
            body.append(sect_pr_copy)

        ET.ElementTree(root).write(document_xml, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for file_path in tmp_path.rglob("*"):
                if file_path.is_file():
                    zout.write(file_path, file_path.relative_to(tmp_path).as_posix())

    validate_zip_member(out_docx, "word/document.xml", "DOCX")
    validate_docx_format_constraints(out_docx)


def validate_docx_format_constraints(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as z:
        root = ET.fromstring(z.read("word/document.xml"))

    body = root.find(qn(W_NS, "body"))
    if body is None:
        raise BuildError(f"DOCX missing body: {path}")

    paragraphs = [p for p in body.findall(qn(W_NS, "p")) if paragraph_text(p)]
    if not paragraphs:
        raise BuildError(f"DOCX has no text paragraphs: {path}")

    all_text = "\n".join(paragraph_text(p) for p in paragraphs)
    if "调研实施建议" in all_text:
        raise BuildError("DOCX must not contain the removed section: 调研实施建议")

    body_checked = 0
    for idx, p in enumerate(paragraphs):
        text = paragraph_text(p)
        style_id = paragraph_style_id(p)
        ind = p.find(f"{qn(W_NS, 'pPr')}/{qn(W_NS, 'ind')}")
        first_line_chars = ind.get(qn(W_NS, "firstLineChars")) if ind is not None else None

        is_title = idx == 0
        is_heading = style_id in {"1", "2", "3", "Heading1", "Heading2", "Heading3"} or re.match(r"^[一二三四五六七八九十]+、", text) or re.match(r"^（[一二三四五六七八九十]+）", text)
        if is_title or is_heading:
            if first_line_chars == "200":
                raise BuildError(f"Heading/title paragraph must not use first-line indent: {text[:30]}")
            continue

        body_checked += 1
        if first_line_chars != "200":
            raise BuildError(f"Body paragraph missing two-character first-line indent: {text[:30]}")

    if body_checked == 0:
        raise BuildError("DOCX format check found no body paragraphs")


def col_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def xml_escape(text: Any) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def sheet_xml(rows: List[List[str]], widths: List[int]) -> str:
    cols = []
    for idx, width in enumerate(widths, start=1):
        cols.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')

    row_xml = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{col_name(c_idx)}{r_idx}"
            style = ' s="1"' if r_idx == 1 else ""
            cells.append(
                f'<c r="{ref}" t="inlineStr"{style}><is><t>{xml_escape(value)}</t></is></c>'
            )
        row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <cols>{"".join(cols)}</cols>
 <sheetData>{"".join(row_xml)}</sheetData>
</worksheet>'''


def build_xlsx(path: Path, sheet_name: str, headers: List[str], blank_rows: int = 20) -> None:
    if not headers:
        raise BuildError(f"Attachment has no headers: {path.name}")
    sheet_name = str(sheet_name or "Sheet1")[:31]
    rows = [headers] + [["" for _ in headers] for _ in range(max(0, blank_rows))]
    widths = [min(max(len(h) + 4, 12), 30) for h in headers]

    workbook = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <sheets><sheet name="{xml_escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
 <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><name val="Microsoft YaHei"/></font></fonts>
 <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
 <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
 <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
 <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
 <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml(rows, widths))

    validate_zip_member(path, "xl/workbook.xml", "XLSX")


def validate_zip_member(path: Path, member: str, label: str) -> None:
    if not path.exists() or path.stat().st_size == 0:
        raise BuildError(f"{label} was not created or is empty: {path}")
    if not zipfile.is_zipfile(path):
        raise BuildError(f"{label} is not a valid zip container: {path}")
    with zipfile.ZipFile(path, "r") as z:
        names = set(z.namelist())
        if member not in names:
            raise BuildError(f"{label} missing required member {member}: {path}")
        bad = z.testzip()
        if bad:
            raise BuildError(f"{label} zip integrity check failed at {bad}: {path}")


def attachment_specs(spec: Dict[str, Any]) -> List[Dict[str, Any]]:
    raw = spec.get("attachments")
    if raw is None:
        default_attachments = DEFAULT_ORGANIZATION_ATTACHMENTS if normalized_research_type(spec) == "organization" else DEFAULT_ATTACHMENTS
        return [
            {"filename": filename, "sheet_name": sheet_name, "headers": headers, "blank_rows": 20}
            for filename, sheet_name, headers in default_attachments
        ]
    if not isinstance(raw, list) or not raw:
        raise BuildError("attachments must be a non-empty list when provided")
    return raw


def build_package(spec_path: Path, out_dir: Path, template: Optional[Path]) -> Path:
    spec = read_json(spec_path)
    research_type = normalized_research_type(spec)
    if research_type == "organization":
        ctx = organization_context(spec)
        region = ctx["region"]
        industry = ctx["industry"]
        organization_name = ctx["organization_name"]
        default_package_name = f"{organization_name}产业大脑合作调研交付包.zip"
        default_docx_name = f"{organization_name}产业大脑合作调研大纲.docx"
    else:
        region = require_text(spec, "region")
        industry = require_text(spec, "industry")
        organization_name = ""
        default_package_name = f"{region}{industry}产业大脑调研交付包.zip"
        default_docx_name = f"{region}{industry}产业大脑调研大纲.docx"

    package_name = sanitize_filename(
        spec.get("package_name") or default_package_name,
        "产业调研交付包.zip",
    )
    if not package_name.lower().endswith(".zip"):
        package_name += ".zip"

    title = sanitize_filename(spec.get("docx_filename") or default_docx_name)
    if not title.lower().endswith(".docx"):
        title += ".docx"

    template_path = template or default_template_path(Path(__file__))
    out_dir.mkdir(parents=True, exist_ok=True)

    staging = out_dir / f".building_{uuid.uuid4().hex}"
    staging.mkdir(parents=True, exist_ok=False)
    try:
        docx_path = staging / title
        build_docx(template_path, docx_path, spec)

        created_files = [docx_path]
        for item in attachment_specs(spec):
            filename = sanitize_filename(item.get("filename", "attachment.xlsx"))
            if not filename.lower().endswith(".xlsx"):
                filename += ".xlsx"
            headers = [str(h) for h in item.get("headers", [])]
            sheet_name = str(item.get("sheet_name") or Path(filename).stem)
            blank_rows = int(item.get("blank_rows", 20))
            xlsx_path = staging / filename
            build_xlsx(xlsx_path, sheet_name, headers, blank_rows)
            created_files.append(xlsx_path)

        manifest = {
            "package": package_name,
            "research_type": research_type,
            "region": region,
            "industry": industry,
            "organization_name": organization_name,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "template": str(template_path),
            "files": [p.name for p in created_files],
        }
        manifest_path = staging / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        note_path = staging / "生成说明.txt"
        note_path.write_text(
            "本交付包由 industry-research-outline Skill 生成。\n"
            "主文档基于参考模板生成，附件为线下调研记录表单。\n"
            "公开资料形成的判断仍需在线下调研中核实。\n",
            encoding="utf-8",
        )
        created_files.extend([manifest_path, note_path])

        package_path = out_dir / package_name
        temp_zip = out_dir / f".{package_name}.{uuid.uuid4().hex}.tmp"
        with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as z:
            for file_path in created_files:
                z.write(file_path, file_path.name)
        validate_zip_member(temp_zip, "manifest.json", "ZIP package")
        if package_path.exists():
            package_path.unlink()
        temp_zip.replace(package_path)
        return package_path
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Build an industry research delivery package.")
    parser.add_argument("--spec", required=True, help="Path to delivery_spec.json")
    parser.add_argument("--out-dir", required=True, help="Output directory")
    parser.add_argument("--template", help="Optional DOCX template path")
    args = parser.parse_args(list(argv) if argv is not None else None)

    try:
        package = build_package(
            spec_path=Path(args.spec).resolve(),
            out_dir=Path(args.out_dir).resolve(),
            template=Path(args.template).resolve() if args.template else None,
        )
    except BuildError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # Defensive final guard for user-facing reliability.
        print(f"[ERROR] Unexpected package build failure: {exc}", file=sys.stderr)
        return 3

    print(f"[OK] Delivery package created: {package}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
