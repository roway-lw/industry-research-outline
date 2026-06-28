# Delivery Spec Schema

Create a JSON file and pass it to `scripts/build_delivery_package.py`.

## Required Fields For Overall Industry Research

```json
{
  "research_type": "industry",
  "region": "陕西省宝鸡市岐山县",
  "industry": "擀面皮产业"
}
```

If `research_type` is omitted, the script treats the task as overall industry research.

## Required Fields For Single-Unit Cooperation Research

```json
{
  "research_type": "organization",
  "organization_name": "吉安人事人才网",
  "unit_category": "生态机构类",
  "unit_type": "人才服务平台",
  "industry": "电子电路产业",
  "project_name": "江西省电子电路产业大脑",
  "cooperation_goal": "探索人才招聘数据共享、行业招聘专区共建和线下活动合作"
}
```

For enterprise units, set `unit_category` or `unit_type` with an enterprise-related value, such as `企业类`, `链主企业`, `龙头企业`, `中小制造企业`, or `数字化标杆企业`.

## Recommended Fields

```json
{
  "title": "陕西省宝鸡市岐山县擀面皮产业大脑调研大纲",
  "docx_filename": "陕西省宝鸡市岐山县擀面皮产业大脑调研大纲.docx",
  "package_name": "陕西省宝鸡市岐山县擀面皮产业大脑调研交付包.zip",
  "outline_blocks": [
    {"type": "title", "text": "陕西省宝鸡市岐山县擀面皮产业大脑调研大纲"},
    {"type": "blank", "text": ""},
    {"type": "heading1", "text": "一、调研背景及目标"},
    {"type": "heading2", "text": "（一）调研背景"},
    {"type": "paragraph", "text": "正文内容"},
    {"type": "heading1", "text": "二、调研范围与重点"}
  ]
}
```

## Block Types

- `title`: document title, using the template title paragraph style
- `blank`: blank paragraph
- `heading1`: first-level heading such as `一、调研背景及目标`
- `heading2`: second-level heading such as `（一）调研背景`
- `heading3`: third-level heading
- `paragraph`: normal paragraph
- `question_group`: question group heading, such as `1.产业链结构`
- `question`: question item, such as `（1）...？`

## Main Outline Sections

For overall industry research, the main DOCX should normally use these sections:

1. 调研背景及目标
2. 调研范围与重点
3. 产业初步认知
4. 调研对象及问题清单
5. 产业大脑场景挖掘重点
6. 资料收集清单
7. 待现场核实问题

Do not create a `调研实施建议` chapter.

For single-unit cooperation research, use one of two structures:

Enterprise:

1. 交流背景及目标
2. 企业基本情况
3. 企业经营与业务现状
4. 核心痛点与需求
5. 产业大脑合作方向
6. 重点交流问题清单
7. 需补充资料清单

Ecosystem institution:

1. 交流背景及目标
2. 机构基本情况
3. 机构资源与服务能力
4. 现有业务与运营模式
5. 产业大脑合作方向
6. 重点交流问题清单
7. 需补充资料清单

## Attachments

If `attachments` is omitted, the script generates default attachments:

- Overall industry research: 8 fieldwork forms.
- Single-unit cooperation research: 5 lighter forms for unit information, interview notes, cooperation opportunities, data/system connection, and action tracking.

To customize:

```json
{
  "attachments": [
    {
      "filename": "附件1_调研对象信息登记表.xlsx",
      "sheet_name": "调研对象信息登记表",
      "headers": ["单位名称", "单位类型", "联系人"],
      "blank_rows": 20
    }
  ]
}
```
