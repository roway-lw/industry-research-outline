# Delivery Package

Use this guide when the user wants downloadable deliverables instead of a text-only draft.

## Package Contents

For overall industry research, the default package should contain:

- `XX地区XX产业大脑调研大纲.docx`
- `附件1_调研对象信息登记表.xlsx`
- `附件2_访谈纪要记录表.xlsx`
- `附件3_资料收集台账.xlsx`
- `附件4_企业调研信息表.xlsx`
- `附件5_产业链环节核实表.xlsx`
- `附件6_数据资源盘点表.xlsx`
- `附件7_产业大脑场景线索表.xlsx`
- `附件8_问题假设验证清单.xlsx`
- `manifest.json`
- `生成说明.txt`

For single-unit cooperation research, the default package should contain:

- `XX单位产业大脑合作调研大纲.docx`
- `附件1_单位基础信息记录表.xlsx`
- `附件2_访谈纪要记录表.xlsx`
- `附件3_合作机会梳理表.xlsx`
- `附件4_数据资源与系统对接盘点表.xlsx`
- `附件5_资料收集与行动清单.xlsx`
- `manifest.json`
- `生成说明.txt`

Do not include a separate `调研实施建议` chapter in the main DOCX.

## Generation Workflow

1. Generate or assemble `delivery_spec.json` according to `delivery-spec-schema.md`.
2. Run `scripts/build_delivery_package.py`.
3. Confirm that the script returns `[OK] Delivery package created: ...`.
4. Deliver the generated ZIP package path to the user.

## Reliability Requirements

The script must:

- use the Linxi DOCX template as the base for the main document
- create all attachment files
- create a manifest
- build a ZIP package
- validate that DOCX, XLSX, and ZIP files are valid zip containers
- fail with a clear error message instead of leaving an incomplete package

## User-Facing Language

Tell the user:

“我会生成一个可下载的调研交付包，里面包含正式调研大纲和线下调研记录附件。”

For single-unit cooperation research, say:

“我会生成一个可下载的单位合作调研交付包，里面包含正式调研大纲和合作记录附件。”

Do not discuss scripts, XML, ZIP internals, or technical implementation unless the user asks.
