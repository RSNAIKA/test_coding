SHORT SUMMARY
-------------
A single-file utility to convert and merge images and PDFs, with advanced layout,
streaming (constant-memory) support, per-image/page controls and CSV/inline mappings.

FEATURE SET (quick)
 - image2pdf        : single image -> PDF
 - images2pdf       : multiple images -> single merged PDF (streaming or in-memory)
 - pdfmerge         : merge multiple PDF files into one
 - pdf2doc          : PDF -> DOCX (requires pdf2docx)
 - doc2pdf          : DOCX -> PDF (requires docx2pdf or libreoffice)
 - remove_pdf_password
 - remove_office_password
Advanced layout & processing:
 - Streaming constant-memory writing via reportlab (--streaming)
 - Progress (tqdm fallback) (--progress)
 - Page size control (named or custom WIDTHxHEIGHT in mm) (--page-size)
 - Per-page custom sizes mapping (--per-page-sizes)
 - Per-image margins (top,right,bottom,left in mm) (--per-image-margins)
 - Per-image rotation override (0|90|180|270) (--per-image-rotation)
 - EXIF auto-rotation (--autorotate)
 - Automatic page orientation selection per page (--auto-orient)
 - Alignment inside content area (--align-h, --align-v)
 - Scaling modes: fit | fill | stretch | original
 - DPI control for mm→pixel/point conversions (--dpi)

USAGE (COMMAND SYNOPSIS)
------------------------
python document_converter.py <infile> <outfile> <conversion> [options]

Where <conversion> is one of:
  image2pdf | images2pdf | pdfmerge | pdf2doc | doc2pdf
  remove_pdf_password | remove_office_password

EXAMPLES (COMMON)
-----------------
# 1) Basic single image -> PDF (A4, 300 DPI, 10mm margins)
python document_converter.py photo.jpg photo.pdf image2pdf

# 2) Directory -> single PDF (streaming, progress)
python document_converter.py ./scans out.pdf images2pdf --streaming --progress --page-size A4 --dpi 300

# 3) Merge PDFs
python document_converter.py "a.pdf,b.pdf" merged.pdf pdfmerge

# 4) Per-image margins and rotation (inline mapping)
python document_converter.py "./scans" out.pdf images2pdf --streaming \
  --per-image-margins "scan1.jpg:10,scan2.jpg:8x12x8x12" \
  --per-image-rotation "scan2.jpg:90" --per-page-sizes "scan1.jpg:210x297,scan2.jpg:297x210" \
  --autorotate --auto-orient --align-h center --align-v center --progress

MAPPING FORMATS (CSV or INLINE)
-------------------------------
Per-page sizes:  filename,210x297  or  filename:210x297
Per-image margins: filename,10  or  filename:8x12x8x12  (top,right,bottom,left)
Per-image rotation: filename,90  (0|90|180|270)
CSV files accept comments (#) and blank lines.

INSTALL (recommended)
---------------------
Recommended (all features):
  pip install pillow reportlab PyPDF2 tqdm pikepdf pdf2docx docx2pdf msoffcrypto-tool

Minimal for streaming images->PDF:
  pip install pillow reportlab

Minimal for basic merging without streaming:
  pip install pillow PyPDF2

