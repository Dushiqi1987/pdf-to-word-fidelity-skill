# Validation Record

Validated on 2026-10-16 using Python 3.12 and Microsoft Word 16 on Windows.

- 16 unit/regression tests passed, covering native equation backends, determinants,
  lower/upper limits, Word-normalized numbering, unresolved fields, template
  cloning, editable tables, TOC structure, and overwrite refusal.
- The public synthetic plan built a three-page DOCX. Local Word finalization and
  PyMuPDF rendering completed. Four equations (three display, one inline), one
  editable data table plus a two-column directory, no embedded image objects,
  resolved page fields, and mirrored headers were checked.
- An independent agent, given only the synthetic PDF and this skill, reconstructed
  one content page. It produced three native equations and a real three-line
  table, rendered the DOCX in Word, and documented small spacing differences.
- A no-skill baseline also completed that bounded page successfully. This is a
  usability/regression check, not evidence that the skill outperforms every
  baseline or that arbitrary PDFs convert without loss.

The four prompts in `evals/evals.json` are reusable acceptance scenarios, not a
claim that all four have completed full blinded end-to-end evaluation. Scanned
documents, unavailable fonts, complex merged tables and arbitrary publishing
layouts remain dependent on the source and the executing agent's verification.

Run `python -m unittest discover -s tests -v` for structural regression checks.
Rendering and source-to-output semantic comparison remain separate gates.

No private source PDF, template, converted manuscript, local audit log, or account
credential is included in this repository.
