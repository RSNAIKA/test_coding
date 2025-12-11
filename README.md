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
