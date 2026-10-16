#!/usr/bin/env python3
"""Audit DOCX structural evidence. This is not visual or semantic validation."""
import argparse
import json
from pathlib import Path
import re
import zipfile

from lxml import etree

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def equation_labels(root):
    labels = set()
    for equation in root.xpath("//m:oMath", namespaces=NS):
        for delimiter in equation.xpath(".//m:d", namespaces=NS):
            begin = delimiter.xpath("./m:dPr/m:begChr/@m:val", namespaces=NS) or ["("]
            end = delimiter.xpath("./m:dPr/m:endChr/@m:val", namespaces=NS) or [")"]
            value = "".join(delimiter.xpath("./m:e//m:t/text()", namespaces=NS)).strip()
            if begin in (["("], ["（"]) and end in ([")"], ["）"]) and re.fullmatch(r"\d+(?:\.\d+)+", value):
                labels.add(value)
        paragraphs = equation.xpath("ancestor::w:p[1]", namespaces=NS)
        if paragraphs:
            value = "".join(paragraphs[0].xpath(".//w:t/text() | .//m:t/text()", namespaces=NS))
            suffix = re.search(r"[（(]\s*(\d+(?:\.\d+)+)\s*[）)]\s*$", value)
            if suffix:
                labels.add(suffix.group(1))
    return labels


def unresolved_pagerefs(root):
    paragraphs = root.xpath("//w:instrText[contains(., 'PAGEREF')]/ancestor::w:p[1]", namespaces=NS)
    count = sum("?" in "".join(p.xpath(".//w:t/text()", namespaces=NS)) for p in paragraphs)
    for field in root.xpath("//w:fldSimple[contains(@w:instr, 'PAGEREF')]", namespaces=NS):
        value = "".join(field.xpath(".//w:t/text()", namespaces=NS)).strip()
        if not value or "?" in value:
            count += 1
    return count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--expected", type=Path, help="Builder plan JSON to check equation numbers and content table count")
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    with zipfile.ZipFile(args.docx) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        header_text = {}
        for name in archive.namelist():
            if re.match(r"word/header\d+\.xml$", name):
                header_text[name] = "".join(etree.fromstring(archive.read(name)).xpath("//w:t/text()", namespaces=NS))
        image_parts = [n for n in archive.namelist() if n.startswith("word/media/")]
    math = root.xpath("//m:oMath", namespaces=NS)
    display = root.xpath("//m:oMathPara/m:oMath", namespaces=NS)
    tables = root.xpath("//w:tbl[not(ancestor::w:tbl)]", namespaces=NS)
    def istoc(t):
        fields = " ".join(t.xpath(".//w:instrText/text() | .//w:fldSimple/@w:instr", namespaces=NS))
        styles = t.xpath(".//w:pStyle/@w:val", namespaces=NS)
        return bool(re.search(r"\b(?:TOC|PAGEREF)\b", fields, re.I)) or any(s.lower().startswith("toc") for s in styles)
    toc_tables = [t for t in tables if istoc(t)]
    text = "".join(root.xpath("//w:t/text() | //m:t/text()", namespaces=NS))
    plain = "\n".join(root.xpath("//w:t/text()", namespaces=NS))
    raw_latex = re.findall(r"\\(?:frac|sum|prod|sqrt|begin|end|mathrm|operatorname|alpha|beta|theta|sigma)\b|\$\$", plain)
    empty_math = [i + 1 for i, m in enumerate(math) if not "".join(m.xpath(".//m:t/text()", namespaces=NS)).strip()]
    errors, warnings = [], []
    if any(re.search(r"Error!|错误[!！]", value, re.I) for value in header_text.values()):
        errors.append("A header field contains a Word error result; verify STYLEREF style names.")
    if raw_latex:
        errors.append("Raw LaTeX commands or display delimiters remain outside native equations.")
    if empty_math:
        errors.append("Empty native equation objects found.")
    unresolved_fields = unresolved_pagerefs(root)
    if unresolved_fields:
        errors.append("Unresolved PAGEREF cached page numbers remain; finalize fields before delivery.")
    expected_numbers = []
    if args.expected:
        plan = json.loads(args.expected.read_text(encoding="utf-8-sig"))
        blocks = plan.get("blocks", [])
        expected_numbers = [str(b["number"]) for b in blocks if b.get("type") == "equation" and b.get("number") is not None]
        expected_equations = sum(b.get("type") == "equation" for b in blocks)
        if len(math) < expected_equations:
            errors.append(f"Native equation count {len(math)} is below expected {expected_equations}.")
        found_numbers = equation_labels(root)
        for n in expected_numbers:
            if n not in found_numbers:
                errors.append(f"Missing equation number ({n}).")
        expected_tables = sum(b.get("type") == "table" for b in blocks)
        if len(tables) - len(toc_tables) != expected_tables:
            errors.append(f"Content table count {len(tables)-len(toc_tables)} differs from expected {expected_tables}.")
        content_tables = [t for t in tables if not istoc(t)]
        for index, b in enumerate(b for b in blocks if b.get("type") == "table"):
            if b.get("three_line") and index < len(content_tables):
                t = content_tables[index]
                visible_vertical = t.xpath(".//w:tblBorders/w:left | .//w:tblBorders/w:right | .//w:tblBorders/w:insideV | .//w:tcBorders/w:left | .//w:tcBorders/w:right | .//w:tcBorders/w:start | .//w:tcBorders/w:end", namespaces=NS)
                if any(n.get("{"+NS["w"]+"}val", "single") not in ("nil", "none") for n in visible_vertical):
                    errors.append(f"Table {index+1} requested three-line borders but has direct visible vertical borders.")
                warnings.append(f"Table {index+1}: verify top/header/bottom rules visually, including inherited style borders.")
    warnings.append("Counts do not prove correct transcription, complete coverage, true three-line appearance, or pixel-level fidelity. Inspect all rendered pages against the PDF and reference.")
    report = {"docx": str(args.docx.resolve()), "native_equations": len(math), "display_equations": len(display),
              "inline_equations": len(math)-len(display), "top_level_tables": len(tables), "toc_tables": len(toc_tables),
              "content_tables": len(tables)-len(toc_tables), "drawing_objects": len(root.xpath("//w:drawing | //w:pict", namespaces=NS)),
              "embedded_media_parts": len(image_parts), "headers": header_text, "raw_latex_hits": raw_latex,
              "empty_equations": empty_math, "expected_numbers": expected_numbers,
              "unresolved_pageref_paragraphs": unresolved_fields, "errors": errors, "warnings": warnings,
              "passed_structural_checks": not errors}
    value = json.dumps(report, ensure_ascii=False, indent=2)
    print(value)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(value, encoding="utf-8")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
