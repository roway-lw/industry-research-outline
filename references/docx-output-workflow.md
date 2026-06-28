# DOCX Output Workflow

Use this workflow whenever the user requests a formal Word document.

## Hard Rule

Do not create a formal DOCX from a blank document. Always use the reference template:

`assets/临西轴承产业大脑调研大纲V1.1.docx`

The template controls page settings, typography, title style, heading levels, paragraph rhythm, numbering feel, and overall document tone.

Before generating any formal DOCX, read `docx-format-constraints.md` and treat it as mandatory. The first-line indentation requirement is additive: preserve the template formatting first, then apply two-character first-line indentation only to body/question paragraphs.

## Required Process

For downloadable packages, prefer `scripts/build_delivery_package.py`, which implements this template-driven process and validates output files.

1. Copy the template DOCX to the intended output path.
2. Preserve the template's document styles, section settings, margins, fonts, colors, line spacing, and heading definitions.
3. Remove or replace the original Linxi-specific body content.
4. Insert the new outline content by reusing the template's real paragraph and run properties:
   - Title for the document title.
   - Heading 1 for `一、...`.
   - Heading 2 for `（一）...`.
   - Heading 3 for lower-level topic headings when needed.
   - Normal for正文 and question text.
   - Body text, question groups, and question items should use first-line indentation of two Chinese characters.
   - Do not merely assign style names if that loses the template's observed formatting.
5. Keep the historical outline's question-list feel:
   - Use Chinese section numbering such as `一、` and `（一）`.
   - Use numbered question groups such as `1.产业链结构`.
   - Use question items such as `（1）...？`.
6. Do not add decorative AI-style covers, cards, tables, gradients, icons, or generic report styling unless the user explicitly requests a redesign.

## Content Replacement Rules

When replacing template content:

- Keep the document title position and style.
- Keep the heading hierarchy simple and close to the sample.
- Keep paragraphs concise and operational.
- Do not convert the main outline into dense tables.
- Use tables only for attachments or repeated recording forms.

## Attachment Files

Attachments may be generated as separate DOCX or XLSX files. If generated as DOCX, they should use a restrained business form style and avoid generic AI report formatting.

For downloadable packages, generate attachments as XLSX files through `scripts/build_delivery_package.py`.

## Verification

Before delivering a DOCX:

- Confirm the output was created from the template, not from a blank document.
- Confirm title, Heading 1, Heading 2, Heading 3, and Normal styles are used.
- Confirm body text inherits the template body formatting and uses first-line indentation of two Chinese characters.
- Confirm titles and headings do not receive body first-line indentation.
- Confirm the result visually resembles the Linxi sample's formatting.
- If the document-generation environment supports rendering, render the DOCX and inspect pages for layout issues.

If rendering is unavailable, state that visual render QA was not completed.
