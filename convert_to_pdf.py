"""
DOCX → PDF Converter (batch)
==============================
Converts all .docx files in the output/ folder to PDF using Microsoft Word.
Formatting is perfectly preserved since Word itself does the rendering.

Usage:
    python convert_to_pdf.py
"""

import os
import sys
import time
from pathlib import Path

import win32com.client

# ── Configuration ────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
INPUT_DIR  = BASE_DIR / "output"
PDF_DIR    = BASE_DIR / "output_pdf"

# Word SaveAs PDF format constant
WD_FORMAT_PDF = 17


def convert_all():
    # Gather all .docx files
    docx_files = sorted(INPUT_DIR.glob("*.docx"))
    if not docx_files:
        print("⚠️  No .docx files found in output/. Run generate_docs.py first.")
        return

    print(f"📄  Found {len(docx_files)} DOCX files to convert.\n")

    # Ensure output directory exists
    PDF_DIR.mkdir(parents=True, exist_ok=True)

    # Launch Word (hidden)
    print("🚀  Starting Microsoft Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    converted = 0
    failed = 0

    try:
        for i, docx_path in enumerate(docx_files, start=1):
            pdf_name = docx_path.stem + ".pdf"
            pdf_path = PDF_DIR / pdf_name

            try:
                doc = word.Documents.Open(str(docx_path))
                doc.SaveAs(str(pdf_path), FileFormat=WD_FORMAT_PDF)
                doc.Close(False)
                converted += 1
                print(f"  ✅  [{i}/{len(docx_files)}]  {pdf_name}")
            except Exception as e:
                failed += 1
                print(f"  ❌  [{i}/{len(docx_files)}]  {docx_path.name}  —  {e}")
    finally:
        word.Quit()

    print(f"\n🎉  Done!  {converted} PDFs saved to: {PDF_DIR}")
    if failed:
        print(f"⚠️   {failed} file(s) failed to convert.")


if __name__ == "__main__":
    convert_all()
