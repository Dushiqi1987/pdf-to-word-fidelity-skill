#!/usr/bin/env python3
"""Inspect local conversion dependencies; install Python dependencies only on request."""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

PACKAGES = {
    "docx": "python-docx", "lxml": "lxml", "pdfplumber": "pdfplumber",
    "pypdfium2": "pypdfium2", "latex2mathml": "latex2mathml",
    "mathml2omml": "mathml2omml", "PIL": "Pillow", "pymupdf": "PyMuPDF",
}


def office_paths():
    roots = [Path(os.environ.get(k, "")) for k in ("ProgramFiles", "ProgramFiles(x86)") if os.environ.get(k)]
    xsls = []
    offices = []
    for root in roots:
        for suffix in ("Microsoft Office/root/Office16", "Microsoft Office/Office16", "Microsoft Office/Office15"):
            base = root / suffix
            if (base / "MML2OMML.XSL").is_file():
                xsls.append(str(base / "MML2OMML.XSL"))
            if (base / "WINWORD.EXE").is_file():
                offices.append(str(base / "WINWORD.EXE"))
    return offices, xsls


def doctor():
    offices, xsls = office_paths()
    registered = False
    if os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
                registered = True
        except OSError:
            pass
    lo = shutil.which("soffice") or shutil.which("libreoffice")
    if not lo:
        for key in ("ProgramFiles", "ProgramFiles(x86)"):
            candidate = Path(os.environ.get(key, "")) / "LibreOffice/program/soffice.exe"
            if candidate.is_file():
                lo = str(candidate)
                break
    modules = {package: importlib.util.find_spec(module) is not None for module, package in PACKAGES.items()}
    return {"python": sys.executable, "python_version": sys.version.split()[0], "packages": modules,
            "missing_packages": [k for k, v in modules.items() if not v], "word_com_registered": registered,
            "word_executables": offices, "libreoffice": lo, "mml2omml_xsl": xsls,
            "renderer_available": registered or bool(lo), "notes": [
                "Office is not installed automatically. A registered Word application may still fail to launch.",
                "Use the same Python interpreter for installation and all scripts. No cloud service is used."]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--install", action="store_true", help="Install missing Python packages with this interpreter's pip")
    parser.add_argument("--out", type=Path, help="Also save the dependency report as JSON")
    args = parser.parse_args()
    report = doctor()
    if args.install and report["missing_packages"]:
        constraints = Path(__file__).resolve().parents[1] / 'requirements.txt'
        command = [sys.executable, "-m", "pip", "install"]
        if constraints.is_file():
            command.extend(['--constraint', str(constraints)])
        subprocess.run([*command, *report["missing_packages"]], check=True)
        importlib.invalidate_caches()
        report = doctor()
    value = json.dumps(report, ensure_ascii=False, indent=2)
    print(value)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(value, encoding="utf-8")
    return 1 if report["missing_packages"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
