#!/usr/bin/env python3
"""
File Converter Script

Supported operations:
1. PDF → DOCX
2. DOCX → PDF
3. Remove PDF password
4. Remove Office (DOCX/DOC/XLSX/etc.) password

Usage:
  python converter.py <input_file> <output_file> <conversion_type> [password]

Conversion types:
  pdf2doc                 Convert PDF → DOCX
  doc2pdf                 Convert DOCX → PDF
  remove_pdf_password      Remove PDF password (requires password)
  remove_office_password   Remove Office file password (requires password)

Examples:
  python converter.py sample.pdf output.docx pdf2doc
  python converter.py report.docx report.pdf doc2pdf
  python converter.py locked.pdf unlocked.pdf remove_pdf_password mypass
  python converter.py locked.docx unlocked.docx remove_office_password mypass
"""

import sys
import os
import shutil
import subprocess

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None

try:
    import pikepdf
except ImportError:
    pikepdf = None

try:
    import msoffcrypto
except ImportError:
    msoffcrypto = None


def pdf_to_docx(infile, outfile):
    if not Converter:
        raise RuntimeError("Missing library: install with `pip install pdf2docx`")
    cv = Converter(infile)
    cv.convert(outfile, start=0, end=None)
    cv.close()
    print(f"[✔] Converted PDF → DOCX: {outfile}")


def docx_to_pdf(infile, outfile):
    if docx2pdf_convert:
        docx2pdf_convert(infile, outfile)
        print(f"[✔] Converted DOCX → PDF (docx2pdf): {outfile}")
        return

    soffice = shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice not found. Install or use docx2pdf.")
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir",
           os.path.dirname(outfile) or ".", os.path.abspath(infile)]
    subprocess.run(cmd, check=True)
    produced = os.path.join(os.path.dirname(outfile), os.path.splitext(os.path.basename(infile))[0] + ".pdf")
    if os.path.exists(produced) and produced != outfile:
        os.replace(produced, outfile)
    print(f"[✔] Converted DOCX → PDF (LibreOffice): {outfile}")


def remove_pdf_password(infile, outfile, password):
    if not pikepdf:
        raise RuntimeError("Missing library: install with `pip install pikepdf`")
    with pikepdf.open(infile, password=password) as pdf:
        pdf.save(outfile)
    print(f"[✔] Removed PDF password → {outfile}")


def remove_office_password(infile, outfile, password):
    if not msoffcrypto:
        raise RuntimeError("Missing library: install with `pip install msoffcrypto-tool`")
    with open(infile, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        if not office.is_encrypted():
            shutil.copy(infile, outfile)
            print(f"[ℹ] File not encrypted, copied to {outfile}")
            return
        office.load_key(password=password)
        with open(outfile, "wb") as out:
            office.decrypt(out)
    print(f"[✔] Removed Office password → {outfile}")


def main():
    if len(sys.argv) < 4 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)

    infile = sys.argv[1]
    outfile = sys.argv[2]
    conversion = sys.argv[3].lower()
    password = sys.argv[4] if len(sys.argv) > 4 else None

    if not os.path.exists(infile):
        print(f"❌ Input file not found: {infile}")
        sys.exit(1)

    try:
        if conversion == "pdf2doc":
            pdf_to_docx(infile, outfile)
        elif conversion == "doc2pdf":
            docx_to_pdf(infile, outfile)
        elif conversion == "remove_pdf_password":
            if not password:
                raise RuntimeError("Please provide password for PDF.")
            remove_pdf_password(infile, outfile, password)
        elif conversion == "remove_office_password":
            if not password:
                raise RuntimeError("Please provide password for Office file.")
            remove_office_password(infile, outfile, password)
        else:
            print("❌ Unknown conversion type. Use one of:")
            print("   pdf2doc | doc2pdf | remove_pdf_password | remove_office_password")
            print("\nFor detailed usage, run:")
            print("   python converter.py --help")
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
