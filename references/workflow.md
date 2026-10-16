# Conversion Workflow

## Contents

1. Contract and evidence
2. Reconstruction
3. Editable equations and tables
4. Layout, numbering, TOC, and headers
5. Preserve later edits
6. Verification and delivery

## 1. Contract and Evidence

Inspect inputs without reading unrelated workspace code. Identify the newest user-supplied DOCX and distinguish a formatting reference from an already edited output. Preserve input hashes and copies outside the deliverable filename.

Use the helpers' `--help` as the command-line source of truth:

```text
python scripts/bootstrap.py --help
python scripts/inspect_inputs.py --help
python scripts/build_docx.py --help
python scripts/render_verify.py --help
python scripts/audit_docx.py --help
```

Typical invocation, using task-local paths rather than hard-coded personal paths:

```text
python scripts/bootstrap.py
python scripts/bootstrap.py --install
python scripts/inspect_inputs.py source.pdf --out inspection --template reference.docx --render
python scripts/build_docx.py --plan plan.json --output result.docx --template reference.docx
python scripts/render_verify.py result.docx --outdir qa
python scripts/audit_docx.py result.docx --expected plan.json
```

The install step is conditional, not required if dependencies are already present. Omit `--template` without a reference. Rendering defaults to preserving the input; use `--update-fields` deliberately for a newly built document requiring final TOC pagination, not automatically for surgical corrections.

The initial report should establish text-layer quality, scan likelihood, page count, equation/table inventory, and required fonts/renderers. PDFs with a few searchable pages can still contain scanned or image-only regions; sample across the entire document.

For DOCX references, inspect `word/document.xml`, styles, settings, and relevant header/footer parts. Effective formatting may be inherited or directly applied. Resolve the hierarchy, not merely `paragraph.style.name`. A document's first table might be a code block, not its academic table template.

Record priorities explicitly:

| Question | Decision |
| --- | --- |
| What governs typography? | PDF, reference DOCX, or written requirements |
| What governs pagination? | Original PDF only if compatible; otherwise reflow |
| What is editable? | Text, OMML equations, actual table cells |
| What can be an image? | Legitimate figures or explicitly accepted exceptions |
| Which numbering changes? | Chapter, subheading, caption, equation, cross-reference |
| Which existing file is authoritative? | Latest user-edited DOCX for revisions |

Do not ask about preferences already clear from the user's request. Ask when fidelity and template constraints materially conflict or a critical symbol is unreadable.

## 2. Reconstruction

Use PDF extraction to collect text, character positions, page geometry, and table candidates. Render pages to inspect the actual reading order. The model then creates a structured content plan using the shipped builder schema/examples. The helper does not replace this reconstruction judgment.

### Builder Plan Contract

The AI writes a JSON object using the builder's supported keys. Read the implementation/help if a requested feature exceeds this contract; do not invent keys and assume they are honored.

- `template_roles`: optional role-to-paragraph-index map for `body`, `chapter`, `h2`, `h3`, `caption`. Indices refer to paragraphs actually inspected in this particular reference, not universal sample indices.
- `format`: optional role formatting. Supported role values can include `font`, `latin`, `size`, `bold`, `align`, `line_spacing`, `space_before`, `space_after`, `first_line_indent`.
- `page`: optional `width`, `height`, `top`, `bottom`, `left`, `right`, in points.
- `headers`: optional `odd` (`"chapter"` or literal text), `even` (literal), `first_page_blank`, and nested `format` containing `font`, `latin`, `size`, `color`, `bold`. Include only when requested.
- `toc`: optional `layout` (`"two-column"` or `"single"`), `title`, `split_after`, `size`, `line`, `gap`. Include only when requested; inspect generated field structure and final pagination.
- `blocks`: ordered semantic content, with `type` identifying a `heading`, `paragraph`, `equation`, `table`, or `image`.

Block content:

| Type | Content fields |
| --- | --- |
| `heading` | `role`: `chapter`, `h2`, or `h3`; `text` |
| `paragraph` | `text`, or `runs` containing text runs with optional `bold`, `italic`, `subscript`, `superscript`, or `{ "math": "LaTeX" }` |
| `equation` | `latex`, `number`, optional `force_below: true` |
| `table` | `caption`, `rows` as string cell arrays, optional `widths` in points, `three_line: true` |
| `image` | `path`, `width`, optional `caption`; use for legitimate figures, not equations/tables |

Keep a separate coverage/uncertainty ledger when fields are not part of the builder schema. This ledger records source page/region, output element ID, element kind, and verification status. Every equation and table should have a stable ID independent of its displayed chapter number. This makes renumbering and targeted repairs safer.

Classify source lines into running furniture, heading levels, paragraphs, captions, equations, tables, and references. Join line wraps conservatively. Preserve real paragraph breaks, punctuation, Latin words, mathematical variables, footnotes, and lists. Repeated page headers should not become body paragraphs, but repeated body text must not be silently deduplicated.

When source extraction fails, inspect cropped regions or full pages. Use OCR only if a suitable engine is available. A doubtful negative sign, Greek letter, subscript, determinant, or index is a content risk, not a typography detail. Keep uncertainty notes and ask for a clearer source where necessary.

## 3. Editable Equations and Tables

### Equations

Transcribe visible math into LaTeX, convert it to MathML, then to native OMML using the available backend. On licensed Office installations, discover the local `MML2OMML.XSL`; never vendor or download a copy as part of this skill. On other systems, use the supported fallback backend and verify its supported syntax.

After conversion, examine both structure and rendered output:

- `w:drawing` is not proof of an equation; an equation image is not editable OMML.
- `m:oMath` / `m:oMathPara` confirm native objects but not semantic correctness.
- Display formulas need display layout, suitable size, and proper limit placement. `m:limLoc` and the relevant under/over structures matter; a lower-right sum index can be wrong even though the index text exists.
- Ensure `det`, `min`, `max`, function names, operators, and Greek letters survive conversion. Check each visible letter, not just a normalized text count.
- Inspect delimiter extent, fraction nesting, integral bounds, accents, and matrix dimensions.
- Use a stable equation-number arrangement with the formula centered and the number at the right. Avoid fragile runs of spaces. Ensure numbering does not alter equation semantics or cause overflow.
- Keep a transcript-to-source mapping and compare all formulas at readable resolution.

Do not use screenshots as an unnoticed fallback. If a formula cannot be reconstructed with sufficient confidence, disclose it and resolve the uncertainty before calling the output complete.

### Tables

Reconstruct cells, row/column grouping, merged cells, units, and notes. Use genuine Word table XML. Widths must fit the actual text area, not blindly reuse PDF coordinates.

For a true three-line table, apply a top border to the table, a bottom border to its header row, and a bottom border to the table. Suppress left/right/inside vertical borders and ordinary internal horizontal borders, including cell-level overrides and inherited table-style borders. Repeat header rows on later pages when appropriate. Keep captions separate from tables and preserve their numbering and spacing.

Render every table to check border visibility, line wrapping, split rows, and cells that lost text. Structural border flags alone cannot prove the visible result.

## 4. Layout, Numbering, TOC, and Headers

Use the reference's actual page size, margins, fonts for separate scripts, font sizes, line spacing, before/after spacing, first-line indentation, and heading hierarchy. Preserve editable styles where practical. Match caption spacing explicitly: matching font and alignment does not imply matching paragraph spacing.

Renumber using a structural map. A change from chapters 1/2 to 6/7 must not rewrite years, table data, or mathematical constants. Search for old chapter-number patterns afterward and review every remaining occurrence. Preserve or update bookmarks and cross-references with the relevant target.

A PDF-style directory may require two columns, hierarchical indentation, dot leaders, right-aligned page numbers, linked heading entries, and a different font scale from the body. The builder's single-column mode uses a native Word TOC field. Its two-column mode uses a borderless layout table, heading hyperlinks, and dynamic PAGEREF fields; it does not automatically add or rename entries when headings change. Rebuild those entries deliberately after structural changes, or implement a native TOC in isolated two-column sections if the user specifically requires automatically expandable entries. Never hard-code page numbers or let TOC columns leak into body layout. Check both displayed field results and underlying hyperlinks. Refresh after final pagination, then inspect every line for truncation, wrapping, or wrong pages.

For requested mirrored running headers:

- Enable different odd and even page headers.
- Set odd/default chapter text right-aligned and even book title left-aligned.
- Handle chapter section starts and `linkToPrevious` deliberately.
- Check whether the first page intentionally differs and whether parity refers to displayed page numbering or physical sheet order.
- Use actual requested color and weight. Inactive headers may look gray in Word even when their printed color is black.

Do not add mirrored headers, Arabic chapter labels, or two-column TOCs to every conversion merely because they were useful in one historical example.

## 5. Preserve Later Edits

For requests such as "only fix equations 6.16 and 6.21," inventory the latest DOCX and create a backup. Find the exact target paragraphs using their native equation numbers and source context, not paragraph indices from an older file.

Prefer ZIP-preserving edits: copy untouched archive entries verbatim; change only targeted XML elements and genuinely necessary relationships. Never reconstruct the whole document from the original PDF when the user has edited the Word output.

Create an allowlist of expected changed parts/nodes. Compare before/after body text, OMML, tables, section properties, styles, relationships, and headers outside the target. Preserve comments, tracked changes, drawings, hyperlinks, and content controls. If an Office save rewrites metadata or field results, inspect that separately and avoid adopting unrelated changes automatically.

Render a disposable copy when possible. The final authoritative file must contain the requested patch without surprise updates. Explain any unavoidable renderer normalization rather than claiming byte-for-byte preservation.

## 6. Verification and Delivery

Use two independent checks:

1. **Structural/semantic:** content coverage, native equation count, equation labels, actual tables/cells, number consistency, field presence, header settings, and preservation of unrequested content.
2. **Visual:** rendered pages, all equations/tables at full resolution, all pages via legible contact sheets, plus TOC, chapter starts, parity pairs, dense pages, and every changed page.

Report which renderer was used. Word and LibreOffice can paginate differently. If neither is available, state that visual verification remains incomplete; do not treat a structural audit as a substitute.

Compare the source and result according to the contract. Template-first conversion may legitimately have different page counts. Editable Word equations need not occupy identical pixels to PDF glyphs, but must retain mathematical meaning, requested limit placement, readable alignment, and correct labels.

The final response should link the DOCX and briefly state what was verified, any known uncertainty, and deliberate differences. Keep private source documents, extracted text, screenshots, backups, and conversion logs out of a public skill package.
