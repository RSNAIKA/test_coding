#!/usr/bin/env python3
"""
Document Converter Script — Full Documentation Header

Author: Ramakrishna Shankara Naika
Email: ramatth78@gmail.com
License: MIT License
Version: 2025-12-11

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

LICENSE (MIT) - include LICENSE file in project root
----------------------------------------------------
MIT License

Copyright (c) 2025 Ramakrishna Shankara Naika

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

NOTES FOR INSTITUTIONS / FORM / DEPLOYMENT
-----------------------------------------
- This script runs locally; no external network calls are made.
- For CI or server deployments, install only needed packages to reduce attack surface.
- Use `--streaming` for large-scale batch jobs (constant memory).
- Use CSV mapping files to standardize per-project layout settings.
- Recommended workflow for scanned document archiving:
  1. Place raw scans in a directory.
  2. Prepare per-image CSVs (sizes/margins/rotations) if needed.
  3. Run streaming conversion with `--progress` and `--auto-orient`.

END OF HEADER
"""

# ---------------------------------------------------------------------------
# Implementation follows (full script)
# ---------------------------------------------------------------------------

import os
import sys
import shutil
import subprocess
import argparse
import io
from math import isfinite
from PIL import Image, ExifTags

# Optional libraries
try:
    import PyPDF2
except Exception:
    PyPDF2 = None

try:
    import pikepdf
except Exception:
    pikepdf = None

try:
    from pdf2docx import Converter
except Exception:
    Converter = None

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# streaming PDF library (reportlab)
try:
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.lib.units import mm as RL_MM
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# progress bar
try:
    from tqdm import tqdm
except Exception:
    tqdm = None

# constants
PAGE_SIZES_MM = {
    "A4": (210.0, 297.0),
    "LETTER": (216.0, 279.0),
    "A3": (297.0, 420.0),
    "A5": (148.0, 210.0),
}

# -----------------------
# Basic helpers
# -----------------------
def mm_to_pixels(mm, dpi):
    inches = mm / 25.4
    return int(round(inches * dpi))

def mm_to_points(mm):
    # reportlab uses points; RL_MM is points-per-mm
    return mm * RL_MM

def parse_page_size(value):
    if not value:
        return None
    v = str(value).strip().upper()
    if v in PAGE_SIZES_MM:
        return PAGE_SIZES_MM[v]
    if "X" in v:
        parts = v.split("X")
        if len(parts) == 2:
            try:
                w = float(parts[0]); h = float(parts[1])
                return (w, h)
            except Exception:
                raise argparse.ArgumentTypeError(f"Invalid page size numbers: {value}")
    raise argparse.ArgumentTypeError(f"Unknown page size: {value}")

def parse_mapping(value, parse_value_func):
    """
    Generic parser for per-image mappings.
    value can be:
     - path to CSV file (lines like: filename,VALUE or filename:VALUE) - comments (#) allowed
     - inline mapping "img1.jpg:210x297,img2.png:12"
    parse_value_func: takes string -> parsed value (or raises)
    returns dict: {basename: parsed_value}
    """
    mapping = {}
    if not value:
        return mapping
    value = str(value).strip()
    if os.path.exists(value) and os.path.isfile(value):
        with open(value, "r", encoding="utf-8") as f:
            for ln in f:
                ln = ln.strip()
                if not ln or ln.startswith("#"):
                    continue
                # accept separators ':' or ','
                if ":" in ln:
                    fname, sval = ln.split(":", 1)
                elif "," in ln:
                    fname, sval = ln.split(",", 1)
                else:
                    continue
                fname = os.path.basename(fname.strip())
                try:
                    mapping[fname] = parse_value_func(sval.strip())
                except Exception:
                    continue
        return mapping
    # inline mapping
    for pair in [p.strip() for p in value.split(",") if p.strip()]:
        if ":" in pair:
            fname, sval = pair.split(":", 1)
        elif "," in pair:
            fname, sval = pair.split(",", 1)
        else:
            continue
        fname = os.path.basename(fname.strip())
        try:
            mapping[fname] = parse_value_func(sval.strip())
        except Exception:
            continue
    return mapping

def parse_margin_value(s):
    # margin can be a single mm value or four values (top,right,bottom,left) separated by 'x' or ','
    s = s.strip()
    if "X" in s.upper():
        parts = [p.strip() for p in s.upper().split("X")]
    elif "," in s:
        parts = [p.strip() for p in s.split(",")]
    else:
        parts = [s]
    vals = []
    for p in parts:
        try:
            v = float(p)
            if not isfinite(v):
                raise ValueError()
            vals.append(v)
        except Exception:
            raise argparse.ArgumentTypeError(f"Invalid margin number: {p}")
    if len(vals) == 1:
        return (vals[0], vals[0], vals[0], vals[0])  # top,right,bottom,left
    if len(vals) == 2:
        return (vals[0], vals[1], vals[0], vals[1])
    if len(vals) == 4:
        return tuple(vals[:4])
    raise argparse.ArgumentTypeError("Margins must be 1,2 or 4 numbers (mm)")

def parse_rotation_value(s):
    # expect 0,90,180,270 (degrees)
    try:
        v = int(s)
        if v % 90 != 0:
            raise argparse.ArgumentTypeError("Rotation must be multiple of 90")
        v = v % 360
        return v
    except Exception:
        raise argparse.ArgumentTypeError("Invalid rotation value; must be integer degrees (0|90|180|270)")

def autorotate_image_if_needed(img):
    """Apply EXIF orientation; returns image (may be same object or rotated copy)."""
    try:
        exif = img._getexif()
        if not exif:
            return img
        orientation_key = next((k for k, v in ExifTags.TAGS.items() if v == 'Orientation'), None)
        if not orientation_key:
            return img
        orientation = exif.get(orientation_key, None)
        if orientation == 3:
            return img.rotate(180, expand=True)
        if orientation == 6:
            return img.rotate(270, expand=True)
        if orientation == 8:
            return img.rotate(90, expand=True)
    except Exception:
        pass
    return img

def get_progress(iterable, total=None, show=True, desc=None):
    if not show:
        return iterable
    if tqdm:
        return tqdm(iterable, total=total, desc=desc)
    def gen():
        i = 0
        for item in iterable:
            i += 1
            print(f"[{desc or 'Progress'}] {i}/{total or '?'}")
            yield item
    return gen()

# -----------------------
# Streaming implementation (reportlab)
# -----------------------
def streaming_images_to_pdf(infiles, outfile, default_page_size_mm=(210,297), dpi=300,
                            default_margin_mm=10, scaling="fit", show_progress=False, sort=False,
                            align_h="center", align_v="center", per_page_sizes=None,
                            per_image_margins=None, per_image_rotation=None, autorotate=False,
                            auto_orient=False):
    """
    Streaming (reportlab) constant-memory writer.
    per_page_sizes: dict basename -> (w_mm,h_mm)
    per_image_margins: dict basename -> (top,right,bottom,left) in mm
    per_image_rotation: dict basename -> degrees (0/90/180/270)
    auto_orient: if True, swap page width/height when image aspect better fits rotated page
    """
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is required for streaming mode: pip install reportlab")

    # build list
    files = []
    if isinstance(infiles, str) and os.path.isdir(infiles):
        entries = sorted(os.listdir(infiles)) if sort else os.listdir(infiles)
        for fn in entries:
            full = os.path.join(infiles, fn)
            if os.path.isfile(full) and fn.lower().split(".")[-1] in ("jpg","jpeg","png","tiff","bmp","webp"):
                files.append(full)
    else:
        parts = [p.strip() for p in str(infiles).split(",") if p.strip()]
        for p in parts:
            if not os.path.exists(p):
                raise FileNotFoundError(p)
            files.append(p)

    if not files:
        raise RuntimeError("No images found to convert/merge.")

    canvas = Canvas(outfile)
    total = len(files)
    iterator = get_progress(files, total=total, show=show_progress, desc="Images")

    for path in iterator:
        basename = os.path.basename(path)
        page_size_mm = per_page_sizes.get(basename, default_page_size_mm) if per_page_sizes else default_page_size_mm
        margin_vals_mm = per_image_margins.get(basename, None) if per_image_margins else None
        if not margin_vals_mm:
            # uniform margin
            margin_vals_mm = (default_margin_mm, default_margin_mm, default_margin_mm, default_margin_mm)
        # rotation override (applied to image pixels)
        rotation_override = per_image_rotation.get(basename, None) if per_image_rotation else None

        # open image and apply autorotate / override rotation
        with Image.open(path) as pil_img:
            if autorotate:
                pil_img = autorotate_image_if_needed(pil_img)
            if rotation_override is not None:
                if rotation_override != 0:
                    pil_img = pil_img.rotate(-rotation_override, expand=True)  # negative because PIL rotates counter-clockwise
            img_w_px, img_h_px = pil_img.size

            # auto-orient: if requested and rotating page makes fit better, swap page dimensions
            page_w_mm, page_h_mm = page_size_mm
            if auto_orient:
                img_ratio = img_w_px / img_h_px if img_h_px != 0 else 1.0
                page_ratio = page_w_mm / page_h_mm if page_h_mm != 0 else 1.0
                # if image and page aspect disagree, swap
                if (img_ratio > 1 and page_ratio < 1) or (img_ratio < 1 and page_ratio > 1):
                    page_w_mm, page_h_mm = page_h_mm, page_w_mm

            page_w_pt = mm_to_points(page_w_mm)
            page_h_pt = mm_to_points(page_h_mm)

            # content area after per-side margins
            top_mm, right_mm, bottom_mm, left_mm = margin_vals_mm
            content_w_pt = page_w_pt - (mm_to_points(left_mm) + mm_to_points(right_mm))
            content_h_pt = page_h_pt - (mm_to_points(top_mm) + mm_to_points(bottom_mm))
            if content_w_pt <= 0 or content_h_pt <= 0:
                raise RuntimeError(f"Margins too large for page size for {basename}")

            # convert image pixel size to points using DPI
            img_w_pt = (img_w_px / dpi) * 72.0
            img_h_pt = (img_h_px / dpi) * 72.0

            # determine target size in points according to scaling
            if scaling == "original":
                target_w_pt, target_h_pt = img_w_pt, img_h_pt
            elif scaling == "stretch":
                target_w_pt, target_h_pt = content_w_pt, content_h_pt
            else:
                ratio_img = img_w_pt / img_h_pt if img_h_pt != 0 else 1.0
                ratio_content = content_w_pt / content_h_pt if content_h_pt != 0 else 1.0
                if (ratio_img > ratio_content and scaling == "fit") or (ratio_img < ratio_content and scaling == "fill"):
                    if scaling == "fit":
                        target_w_pt = content_w_pt
                        target_h_pt = target_w_pt / ratio_img
                    else:  # fill
                        target_h_pt = content_h_pt
                        target_w_pt = target_h_pt * ratio_img
                else:
                    if scaling == "fit":
                        target_h_pt = content_h_pt
                        target_w_pt = target_h_pt * ratio_img
                    else:
                        target_w_pt = content_w_pt
                        target_h_pt = target_w_pt / ratio_img

            # compute position according to alignment and margins
            if align_h == "left":
                x_pt = mm_to_points(left_mm)
            elif align_h == "right":
                x_pt = page_w_pt - mm_to_points(right_mm) - target_w_pt
            else:  # center
                x_pt = mm_to_points(left_mm) + (content_w_pt - target_w_pt)/2.0

            if align_v == "top":
                y_pt = page_h_pt - mm_to_points(top_mm) - target_h_pt
            elif align_v == "bottom":
                y_pt = mm_to_points(bottom_mm)
            else:  # center
                y_pt = mm_to_points(bottom_mm) + (content_h_pt - target_h_pt)/2.0

            # set page size and draw
            canvas.setPageSize((page_w_pt, page_h_pt))
            # Use ImageReader with PIL image object (in-memory) to avoid temp files
            # Convert PIL to RGB if needed
            if pil_img.mode not in ("RGB", "RGBA"):
                pil_img = pil_img.convert("RGB")
            img_reader = ImageReader(pil_img)
            canvas.drawImage(img_reader, x_pt, y_pt, width=target_w_pt, height=target_h_pt, preserveAspectRatio=False, anchor='sw')
            canvas.showPage()

    canvas.save()
    print(f"[✔] Streaming (constant memory) wrote {total} images -> {outfile}")

# -----------------------
# Non-streaming implementation (Pillow)
# -----------------------
def images_to_pdf_non_streaming(infiles, outfile, default_page_size_mm=(210,297), dpi=300,
                                default_margin_mm=10, scaling="fit", show_progress=False, sort=False,
                                align_h="center", align_v="center", per_page_sizes=None,
                                per_image_margins=None, per_image_rotation=None, autorotate=False,
                                auto_orient=False):
    # build file list
    files = []
    if isinstance(infiles, str) and os.path.isdir(infiles):
        entries = sorted(os.listdir(infiles)) if sort else os.listdir(infiles)
        for fn in entries:
            full = os.path.join(infiles, fn)
            if os.path.isfile(full) and fn.lower().split(".")[-1] in ("jpg","jpeg","png","tiff","bmp","webp"):
                files.append(full)
    else:
        parts = [p.strip() for p in str(infiles).split(",") if p.strip()]
        for p in parts:
            if not os.path.exists(p):
                raise FileNotFoundError(p)
            files.append(p)

    if not files:
        raise RuntimeError("No images found to convert/merge.")

    images_pages = []
    iterator = get_progress(files, total=len(files), show=show_progress, desc="Images")

    for path in iterator:
        basename = os.path.basename(path)
        page_size_mm = per_page_sizes.get(basename, default_page_size_mm) if per_page_sizes else default_page_size_mm
        margin_vals_mm = per_image_margins.get(basename, None) if per_image_margins else None
        if not margin_vals_mm:
            margin_vals_mm = (default_margin_mm, default_margin_mm, default_margin_mm, default_margin_mm)
        rotation_override = per_image_rotation.get(basename, None) if per_image_rotation else None

        # possibly auto-swap page orientation later; first open image
        with Image.open(path) as im:
            if autorotate:
                im = autorotate_image_if_needed(im)
            if rotation_override is not None and rotation_override != 0:
                im = im.rotate(-rotation_override, expand=True)
            img_w_px, img_h_px = im.size

            page_w_mm, page_h_mm = page_size_mm
            if auto_orient:
                img_ratio = img_w_px / img_h_px if img_h_px != 0 else 1.0
                page_ratio = page_w_mm / page_h_mm if page_h_mm != 0 else 1.0
                if (img_ratio > 1 and page_ratio < 1) or (img_ratio < 1 and page_ratio > 1):
                    page_w_mm, page_h_mm = page_h_mm, page_w_mm

            pw_px = mm_to_pixels(page_w_mm, dpi)
            ph_px = mm_to_pixels(page_h_mm, dpi)
            top_mm, right_mm, bottom_mm, left_mm = margin_vals_mm
            top_px = mm_to_pixels(top_mm, dpi)
            right_px = mm_to_pixels(right_mm, dpi)
            bottom_px = mm_to_pixels(bottom_mm, dpi)
            left_px = mm_to_pixels(left_mm, dpi)
            content_w = max(1, pw_px - (left_px + right_px))
            content_h = max(1, ph_px - (top_px + bottom_px))

            # scaling logic
            img_w, img_h = im.size
            if scaling == "original":
                target_w, target_h = img_w, img_h
            elif scaling == "stretch":
                target_w, target_h = content_w, content_h
            else:
                ratio_img = img_w / img_h if img_h != 0 else 1.0
                ratio_content = content_w / content_h if content_h != 0 else 1.0
                if (ratio_img > ratio_content and scaling == "fit") or (ratio_img < ratio_content and scaling == "fill"):
                    if scaling == "fit":
                        target_w = content_w
                        target_h = int(round(target_w / ratio_img))
                    else:
                        target_h = content_h
                        target_w = int(round(target_h * ratio_img))
                else:
                    if scaling == "fit":
                        target_h = content_h
                        target_w = int(round(target_h * ratio_img))
                    else:
                        target_w = content_w
                        target_h = int(round(target_w / ratio_img))

            # resize if needed
            if scaling != "original":
                resized = im.resize((max(1, int(target_w)), max(1, int(target_h))), Image.LANCZOS)
            else:
                resized = im

            # create page and paste according to alignment + margins
            page = Image.new("RGB", (pw_px, ph_px), (255,255,255))
            if align_h == "left":
                x = left_px
            elif align_h == "right":
                x = pw_px - right_px - resized.width
            else:
                x = left_px + (content_w - resized.width)//2

            if align_v == "top":
                y = top_px + (content_h - resized.height)
            elif align_v == "bottom":
                y = bottom_px
            else:
                y = bottom_px + (content_h - resized.height)//2

            page.paste(resized.convert("RGB"), (int(x), int(y)))
            images_pages.append(page.copy())

    # save pages as PDF
    first, rest = images_pages[0], images_pages[1:]
    first.save(outfile, "PDF", resolution=dpi, save_all=True, append_images=rest)
    for im in images_pages:
        try: im.close()
        except Exception: pass
    print(f"[✔] Merged {len(images_pages)} images -> {outfile} (non-streaming)")

# -----------------------
# Other helpers (pdf merge, pdf->docx, docx->pdf)
# -----------------------
def pdf_merge(infiles, outfile):
    files = [p.strip() for p in str(infiles).split(",") if p.strip()]
    if not files:
        raise RuntimeError("No PDF files provided to merge.")
    for p in files:
        if not os.path.exists(p):
            raise FileNotFoundError(p)

    if PyPDF2:
        merger = PyPDF2.PdfMerger()
        try:
            for p in files:
                merger.append(p)
            with open(outfile, "wb") as out_f:
                merger.write(out_f)
            print(f"[✔] Merged {len(files)} PDFs -> {outfile} (PyPDF2)")
        finally:
            try:
                merger.close()
            except Exception:
                pass
        return

    if pikepdf:
        merged = pikepdf.Pdf.new()
        for p in files:
            src = pikepdf.Pdf.open(p)
            merged.pages.extend(src.pages)
        merged.save(outfile)
        print(f"[✔] Merged {len(files)} PDFs -> {outfile} (pikepdf)")
        return

    raise RuntimeError("Install PyPDF2 or pikepdf to enable pdf merging: `pip install PyPDF2 pikepdf`")

def pdf_to_docx(infile, outfile):
    if not Converter:
        raise RuntimeError("Missing library: install with `pip install pdf2docx`")
    cv = Converter(infile)
    cv.convert(outfile, start=0, end=None)
    cv.close()
    print(f"[✔] Converted PDF -> DOCX: {outfile}")

def docx_to_pdf(infile, outfile):
    if docx2pdf_convert:
        docx2pdf_convert(infile, outfile)
        print(f"[✔] Converted DOCX -> PDF (docx2pdf): {outfile}")
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
    print(f"[✔] Converted DOCX -> PDF (LibreOffice): {outfile}")

# -----------------------
# CLI
# -----------------------
def build_parser():
    p = argparse.ArgumentParser(prog="document_converter.py",
                                description="Convert/merge documents and images to PDF and vice-versa.")
    p.add_argument("infile", help="Input file(s). For multiple, use comma-separated list or a directory for images.")
    p.add_argument("outfile", help="Output file path.")
    p.add_argument("conversion", help="Conversion type: image2pdf | images2pdf | pdfmerge | pdf2doc | doc2pdf | remove_pdf_password | remove_office_password",
                   type=str)
    # page options
    p.add_argument("--page-size", "-P", default="A4", type=parse_page_size,
                   help="Page size name (A4, LETTER) or WIDTHxHEIGHT in mm (e.g. 210x297). Default A4.")
    p.add_argument("--dpi", "-d", default=300, type=int, help="DPI used for mm->pixels conversion. Default 300.")
    # margins
    p.add_argument("--margin-mm", "-m", default=10.0, type=float, help="Default margin (mm). Applies to all sides if not overridden.")
    p.add_argument("--per-image-margins", default=None,
                   help="Per-image margins mapping. CSV/inline mapping. Values: one number or 2 or 4 numbers (mm). Example: 'a.jpg:10,b.jpg:5x5x5x5'")
    # scaling & alignment
    p.add_argument("--scaling", "-s", default="fit", choices=["fit","fill","stretch","original"],
                   help="Image scaling mode. Default 'fit'.")
    p.add_argument("--align-h", choices=["left","center","right"], default="center", help="Horizontal alignment inside content area.")
    p.add_argument("--align-v", choices=["top","center","bottom"], default="center", help="Vertical alignment inside content area.")
    # rotation / autorotate
    p.add_argument("--per-image-rotation", default=None,
                   help="Per-image rotation override mapping (degrees 0|90|180|270). CSV or inline. e.g. 'a.jpg:90,b.jpg:180'")
    p.add_argument("--autorotate", action="store_true", help="Auto-rotate images based on EXIF orientation before placing.")
    p.add_argument("--auto-orient", action="store_true", help="Automatically choose page orientation per page (portrait/landscape) so image fits better.")
    # per-page sizes
    p.add_argument("--per-page-sizes", default=None,
                   help="Per-image page sizes mapping. CSV/inline mapping. e.g. 'a.jpg:210x297,b.jpg:148x210'")
    # streaming, progress, sorting
    p.add_argument("--streaming", action="store_true", help="Use streaming (reportlab) constant-memory PDF creation.")
    p.add_argument("--progress", action="store_true", help="Show progress output for large image sets (uses tqdm if available).")
    p.add_argument("--sort", action="store_true", help="Sort directory image listing alphabetically before merging.")
    p.add_argument("--password", "-p", default=None, help="Password for password-removal commands.")
    return p

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    parser = build_parser()
    args = parser.parse_args()

    infile = args.infile
    outfile = args.outfile
    conversion = args.conversion.lower()
    page_size = args.page_size or PAGE_SIZES_MM["A4"]
    dpi = args.dpi
    margin = args.margin_mm
    scaling = args.scaling
    align_h = args.align_h
    align_v = args.align_v
    show_progress = args.progress
    sort = args.sort
    password = args.password
    streaming = args.streaming
    autorotate = args.autorotate
    auto_orient = args.auto_orient

    per_page_sizes = {}
    per_image_margins = {}
    per_image_rotation = {}
    if args.per_page_sizes:
        per_page_sizes = parse_mapping(args.per_page_sizes, parse_page_size)
    if args.per_image_margins:
        per_image_margins = parse_mapping(args.per_image_margins, parse_margin_value)
    if args.per_image_rotation:
        per_image_rotation = parse_mapping(args.per_image_rotation, parse_rotation_value)

    try:
        if conversion in ("image2pdf", "image_to_pdf", "img2pdf"):
            # single image
            if streaming:
                streaming_images_to_pdf(infile, outfile, default_page_size_mm=page_size, dpi=dpi,
                                        default_margin_mm=margin, scaling=scaling, show_progress=show_progress,
                                        sort=sort, align_h=align_h, align_v=align_v,
                                        per_page_sizes=per_page_sizes, per_image_margins=per_image_margins,
                                        per_image_rotation=per_image_rotation, autorotate=autorotate,
                                        auto_orient=auto_orient)
            else:
                images_to_pdf_non_streaming(infile, outfile, default_page_size_mm=page_size, dpi=dpi,
                                            default_margin_mm=margin, scaling=scaling, show_progress=show_progress,
                                            sort=sort, align_h=align_h, align_v=align_v,
                                            per_page_sizes=per_page_sizes, per_image_margins=per_image_margins,
                                            per_image_rotation=per_image_rotation, autorotate=autorotate,
                                            auto_orient=auto_orient)
        elif conversion in ("images2pdf", "images_to_pdf", "imgmergepdf", "image_merge_pdf"):
            if streaming:
                streaming_images_to_pdf(infile, outfile, default_page_size_mm=page_size, dpi=dpi,
                                        default_margin_mm=margin, scaling=scaling, show_progress=show_progress,
                                        sort=sort, align_h=align_h, align_v=align_v,
                                        per_page_sizes=per_page_sizes, per_image_margins=per_image_margins,
                                        per_image_rotation=per_image_rotation, autorotate=autorotate,
                                        auto_orient=auto_orient)
            else:
                images_to_pdf_non_streaming(infile, outfile, default_page_size_mm=page_size, dpi=dpi,
                                            default_margin_mm=margin, scaling=scaling, show_progress=show_progress,
                                            sort=sort, align_h=align_h, align_v=align_v,
                                            per_page_sizes=per_page_sizes, per_image_margins=per_image_margins,
                                            per_image_rotation=per_image_rotation, autorotate=autorotate,
                                            auto_orient=auto_orient)
        elif conversion in ("pdfmerge", "mergepdf", "pdf_merge"):
            pdf_merge(infile, outfile)
        else:
            # other legacy conversions
            if conversion == "pdf2doc":
                pdf_to_docx(infile, outfile)
            elif conversion == "doc2pdf":
                docx_to_pdf(infile, outfile)
            elif conversion == "remove_pdf_password":
                if not password:
                    raise RuntimeError("Please provide password for PDF using --password or -p")
                if not pikepdf:
                    raise RuntimeError("pikepdf required for removing PDF password: pip install pikepdf")
                with pikepdf.open(infile, password=password) as pdf:
                    pdf.save(outfile)
                print(f"[✔] Removed PDF password -> {outfile}")
            elif conversion == "remove_office_password":
                try:
                    import msoffcrypto
                except Exception:
                    raise RuntimeError("Install msoffcrypto: pip install msoffcrypto-tool")
                with open(infile, "rb") as f:
                    office = msoffcrypto.OfficeFile(f)
                    if not office.is_encrypted():
                        shutil.copy(infile, outfile)
                        print(f"[ℹ] File not encrypted, copied to {outfile}")
                    else:
                        office.load_key(password=password)
                        with open(outfile, "wb") as out:
                            office.decrypt(out)
                        print(f"[✔] Removed Office password -> {outfile}")
            else:
                print("❌ Unknown conversion type. Use one of:")
                print("   pdf2doc | doc2pdf | remove_pdf_password | remove_office_password")
                print("   image2pdf | images2pdf | pdfmerge")
                print("\nFor detailed usage, run:")
                print("   python document_converter.py --help")
                sys.exit(1)
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
