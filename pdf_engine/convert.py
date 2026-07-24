"""
PDF ⇄ DOCX 统一转换入口 — auto-routes by file extension.

Usage:
  python convert.py input.pdf  output.docx     PDF → DOCX
  python convert.py input.docx output.pdf      DOCX → PDF (Windows COM)
  python convert.py input.pdf                  auto-names output
"""

import sys
import os

ROUTE_MAP = {
    ('.pdf', '.docx'): 'pdf_to_docx',
    ('.docx', '.pdf'): 'docx_to_pdf',
    ('.html', '.pdf'): 'html_to_pdf',
    ('.htm', '.pdf'): 'html_to_pdf',
}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(src):
        print(f"[ERROR] File not found: {src}", file=sys.stderr)
        sys.exit(1)

    src_ext = os.path.splitext(src)[1].lower()
    dst_ext = os.path.splitext(dst)[1].lower() if dst else None

    # Auto-detect output extension
    if dst_ext is None:
        dst_ext = '.docx' if src_ext == '.pdf' else '.pdf'
        dst = src.rsplit('.', 1)[0] + dst_ext

    route = ROUTE_MAP.get((src_ext, dst_ext))
    if not route:
        print(f"[ERROR] Unsupported: {src_ext} → {dst_ext}", file=sys.stderr)
        sys.exit(1)

    if route == 'pdf_to_docx':
        from universal_pdf_converter import UniversalPDFConverter
        try:
            out = UniversalPDFConverter().convert(src, dst)
            print(out)
            sys.exit(0)
        except Exception as e:
            print(f"[ERROR] {e}", file=sys.stderr)
            sys.exit(1)

    elif route == 'docx_to_pdf':
        from docx_to_pdf import DocxToPdfConverter
        try:
            out = DocxToPdfConverter().convert(src, dst)
            print(out)
            sys.exit(0)
        except Exception as e:
            print(f"[ERROR] {e}", file=sys.stderr)
            sys.exit(1)

    elif route == 'html_to_pdf':
        from html_to_pdf import convert as html_to_pdf_convert
        try:
            out = html_to_pdf_convert(src, dst)
            print(out)
            sys.exit(0)
        except Exception as e:
            print(f"[ERROR] {e}", file=sys.stderr)
            sys.exit(1)


if __name__ == '__main__':
    main()
