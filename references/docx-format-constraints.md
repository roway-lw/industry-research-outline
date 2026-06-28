# DOCX Format Constraints

These constraints are mandatory for formal Word deliverables. They have higher priority than ordinary writing-style guidance.

## Template Is Mandatory

Use `assets/临西轴承产业大脑调研大纲V1.1.docx` as the base for every formal research-outline DOCX.

Do not create the formal DOCX from a blank document. Do not use Word default styling, generic AI report styling, decorative cover pages, cards, icons, gradients, or table-heavy report layouts unless the user explicitly asks for a redesign.

## Preserve Template Formatting

The output must preserve the template's:

- page settings and margins
- title position and title formatting
- heading levels and heading formatting
- font family, font size, color, and weight
- line spacing, paragraph spacing, alignment, and numbering feel
- restrained government/park project document tone

When inserting new content, copy the template's real paragraph properties and run properties. Do not merely write style names such as `Normal`, `Heading1`, or `Heading2` if doing so loses the observed template formatting.

## First-Line Indent Rule

First-line indent is an additive constraint, not a replacement for template formatting.

- Body paragraphs, question groups, and question items must inherit the template body style first, then apply a two-Chinese-character first-line indent.
- In WordprocessingML, the expected value is `w:firstLineChars="200"`.
- Keep any compatible template indentation values such as `w:firstLine` when inherited from the sample.
- Document title, chapter headings, section headings, and lower-level headings must not receive body first-line indentation.

## Script Guardrails

For downloadable packages, `scripts/build_delivery_package.py` must:

- extract title, heading, and body paragraph samples from the template document;
- reuse those samples when creating new paragraphs;
- add two-character first-line indentation only to body/question paragraphs;
- preserve the template section properties;
- validate that the generated DOCX is a valid DOCX zip;
- validate that body/question paragraphs contain `w:firstLineChars="200"`;
- validate that title and headings do not use body first-line indentation;
- validate that the removed section `调研实施建议` does not appear.

If any of these checks fail, stop and fix the generation logic before delivering the file.
