# Lessons and Provenance

## What This Skill Preserves From the Workflow

This skill distills a PDF-to-Word reconstruction workflow and its later user-requested repairs. It contains original operational guidance, not copies of private documents or other skill instruction bodies.

| Observed failure mode | Reusable rule |
| --- | --- |
| Equations look acceptable but are pictures | Require native OMML and verify editability separately from appearance |
| A sum index appears at the lower right | Inspect display limits and enforce under/over placement when the source requires it |
| Conversion loses part of an operator name | Visually verify operators such as determinant, not only object counts |
| A “three-line table” becomes a grid table | Verify table, row, cell, and inherited border rules |
| Chinese chapter numerals differ from the source | Treat numeral representation as configurable formatting |
| Headers are on the wrong side or too pale | Check odd/even section behavior and actual printed color |
| An attractive TOC has stale page numbers | Use linked Word TOC fields and refresh after pagination |
| The user edited the Word between turns | Patch the latest file surgically and audit unrelated content |
| Page counts change after matching a reference | Explain template-first reflow instead of promising identical pagination |

## Methodological Sources

The following resources informed the approach. They are not required to be installed, and their instruction bodies are not redistributed here.

- **OpenAI Documents skill methodology:** render document artifacts and inspect page images before delivery; separate successful creation from verified visual correctness. Official skills collection: <https://github.com/openai/skills>.
- **OpenAI PDF skill methodology:** combine text extraction with rendered-page inspection because PDF reading order and geometry matter. Official skills collection: <https://github.com/openai/skills>.
- **find-skills discovery methodology:** discover a specialized capability when an existing method does not meet requirements, then inspect its suitability rather than assuming its existence solves the task. Community skills discovery: <https://github.com/vercel-labs/skills>.
- **md2docx-math:** the LaTeX → MathML → OMML conversion path and the distinction between editable math and equation images. Project: <https://github.com/advancehs/md2docx-math>; skill entry: <https://github.com/advancehs/md2docx-math/blob/main/SKILL.md>.
- **Microsoft Office Math/Open XML:** Word's native equation structures and Office-hosted conversion support. Documentation: <https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-documents> and <https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math>.

These links identify provenance, not guarantees about third-party availability, exact versions, behavior, or license compatibility. Inspect the current source and license before copying or redistributing third-party code. The bundled conversion helpers are maintained independently of those skill bodies.

## Dependency and Redistribution Boundaries

Python packages used by the scripts remain third-party dependencies under their own licenses. Install them from their normal package distribution channels; do not embed an entire runtime or environment in the skill. Pin or record tested versions when reproducing a conversion.

Microsoft Word, Office `MML2OMML.XSL`, fonts, LibreOffice, and OCR engines are not included. Discover legally installed local components. Do not copy Office stylesheets into this repository or fetch unofficial copies to bypass a missing installation. A fallback converter is not a guarantee of identical rendering or complete mathematical coverage.

Input PDFs, sample DOCX files, screenshots, transcribed source text, personal paths, authentication data, and user-edited outputs are private task materials. Do not put them in fixtures, evals, commits, issue reports, or public uploads. Tests and examples must be synthetic or independently licensed for redistribution.

## Limits of the Term “Lossless”

PDF typically stores visual placement rather than a complete editable document model. Reconstructing paragraphs, table semantics, and equation syntax can require inference. OCR adds uncertainty, and different layout engines or fonts can change line breaks. Therefore this skill aims for faithful content and verified editable formatting within a declared contract; it does not claim universal, automatic, pixel-identical lossless conversion.
