#!/usr/bin/env python3
"""Render DOCX locally without changing its bytes, unless field updates are explicitly requested."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

from bootstrap import doctor


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def word_export(docx, pdf, update=False):
    def quote(value):
        return "'" + str(value).replace("'", "''") + "'"
    script = """$ErrorActionPreference='Stop'
$w = $null
$d = $null
try {
 $w = New-Object -ComObject Word.Application
 $w.Visible = $false
 $w.DisplayAlerts = 0
 $w.AutomationSecurity = 3
 $d = $w.Documents.Open(INPUT, $false, READONLY, $false)
 UPDATE
 $d.ExportAsFixedFormat(OUTPUT, 17)
} finally {
 if ($null -ne $d) { $d.Close(0) }
 if ($null -ne $w) { $w.Quit() }
}
""".replace("INPUT", quote(docx)).replace("OUTPUT", quote(pdf)).replace("READONLY", "$false" if update else "$true")
    update_code = """$d.Fields.Update() | Out-Null
 foreach ($toc in $d.TablesOfContents) { $toc.Update() }
 $d.Repaginate()
 $d.Fields.Update() | Out-Null
 foreach ($section in $d.Sections) {
  foreach ($header in $section.Headers) { if ($header.Exists) { $header.Range.Fields.Update() | Out-Null } }
  foreach ($footer in $section.Footers) { if ($footer.Exists) { $footer.Range.Fields.Update() | Out-Null } }
 }
 $d.Save()""" if update else ""
    script = script.replace("UPDATE", update_code)
    with tempfile.TemporaryDirectory(prefix="docx-word-render-") as tmp:
        ps = Path(tmp) / "export.ps1"
        ps.write_text(script, encoding="utf-8-sig")
        # PowerShell 7 avoids stale Office PIA casts seen in some WinPS 5.1 setups.
        shell = shutil.which("pwsh") or shutil.which("powershell")
        if not shell:
            raise RuntimeError("PowerShell is required for Word COM export on Windows.")
        subprocess.run([shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(ps)],
                       check=True, timeout=300)


def libreoffice_export(docx, out, executable):
    with tempfile.TemporaryDirectory(prefix="docx-lo-render-") as tmp:
        root = Path(tmp)
        source = root / docx.name
        shutil.copy2(docx, source)
        profile = (root / "profile").as_uri()
        subprocess.run([executable, "-env:UserInstallation=" + profile, "--headless", "--convert-to", "pdf",
                        "--outdir", str(root), str(source)], check=True, timeout=300)
        result = root / (source.stem + ".pdf")
        if not result.is_file():
            raise RuntimeError("LibreOffice did not produce the expected PDF.")
        shutil.copy2(result, out)


def render_pdf(pdf, out, scale):
    try:
        import pymupdf
        doc = pymupdf.open(str(pdf))
        count = len(doc)
        texts = []
        for i, page in enumerate(doc, 1):
            page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False).save(str(out / f"page-{i:03d}.png"))
            texts.append({"page": i, "text": page.get_text()})
        doc.close()
        return count, "PyMuPDF", texts
    except ImportError:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(str(pdf))
        count = len(doc)
        for i in range(count):
            page = doc[i]
            bitmap = page.render(scale=scale)
            bitmap.to_pil().save(out / f"page-{i+1:03d}.png")
            bitmap.close()
            page.close()
        doc.close()
        return count, "PDFium fallback: visually verify all glyphs", []


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--outdir", required=True, type=Path)
    parser.add_argument("--renderer", choices=["auto", "word", "libreoffice"], default="auto")
    parser.add_argument("--scale", type=float, default=1.5)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--update-fields", action="store_true", help="Explicitly update fields and SAVE the input in Word before rendering")
    group.add_argument("--finalized-docx", type=Path, help="Copy input to this different path, update fields and save the copy in Word")
    args = parser.parse_args()
    source = args.docx.resolve(strict=True)
    original_hash = digest(source)
    args.outdir.mkdir(parents=True, exist_ok=True)
    target = source
    if args.finalized_docx:
        target = args.finalized_docx.resolve()
        if target == source:
            parser.error("--finalized-docx must differ from the input path")
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            parser.error("--finalized-docx already exists; use a new path to avoid overwriting user changes")
        shutil.copy2(source, target)
    update = args.update_fields or args.finalized_docx is not None
    info = doctor()
    renderer = args.renderer
    if renderer == "auto":
        renderer = "word" if os.name == "nt" and info["word_com_registered"] else "libreoffice"
    pdf = args.outdir.resolve() / "rendered.pdf"
    if renderer == "word":
        if os.name != "nt" or not info["word_com_registered"]:
            raise RuntimeError("Word COM is unavailable. Install/configure a local renderer yourself; no cloud fallback is used.")
        word_export(target, pdf, update)
    else:
        if update:
            raise RuntimeError("Saving updated field caches requires Word COM in this workflow. Read-only LibreOffice rendering remains available.")
        if not info["libreoffice"]:
            raise RuntimeError("No local Word or LibreOffice renderer is available. Do not claim visual verification.")
        libreoffice_export(target, pdf, info["libreoffice"])
    if not args.update_fields and digest(source) != original_hash:
        raise RuntimeError("Input DOCX bytes changed unexpectedly during rendering.")
    count, image_renderer, texts = render_pdf(pdf, args.outdir, args.scale)
    report = {"input": str(source), "rendered_docx": str(target), "pdf": str(pdf), "pages": count,
              "docx_renderer": renderer, "image_renderer": image_renderer,
              "input_preserved": digest(source) == original_hash, "fields_explicitly_updated": update,
              "visual_qa": "NOT AUTOMATED: inspect every page PNG, including formulas, tables, TOC, and odd/even headers."}
    (args.outdir / "render-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    (args.outdir / "rendered-text.json").write_text(json.dumps(texts, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
