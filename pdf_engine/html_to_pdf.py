"""
Windows-native HTML → PDF converter via WPS / Word COM automation.
Word/WPS can natively open .html files and save as PDF.

Usage:
  python html_to_pdf.py input.html [output.pdf]
  Exit code 0 = success, 1 = failure
"""

import sys
import os
import logging
import time

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger("HtmlToPdf")

PDF_FORMAT = 17  # wdFormatPDF


def convert(html_path: str, pdf_path: str = None) -> str:
    """
    Convert HTML to PDF via WPS or Word COM.
    Returns the output PDF path.
    """
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"HTML not found: {html_path}")

    if pdf_path is None:
        pdf_path = html_path.rsplit(".", 1)[0] + ".pdf"

    html_abs = os.path.abspath(html_path)
    pdf_abs = os.path.abspath(pdf_path)

    # Delete existing output to avoid overwrite prompt
    if os.path.exists(pdf_abs):
        try:
            os.remove(pdf_abs)
        except Exception:
            pass

    # Try WPS first, then Word
    for prog_id, name in [("Wps.Application", "WPS"), ("Word.Application", "Word")]:
        logger.info(f"Trying {name}: {html_abs} → {pdf_abs}")
        app = None
        doc = None
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            from win32com.client import Dispatch
            app = Dispatch(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            try:
                app.AutomationSecurity = 3
            except Exception:
                pass

            doc = app.Documents.Open(html_abs, ReadOnly=True)
            doc.SaveAs(pdf_abs, FileFormat=PDF_FORMAT)
            doc.Close()
            doc = None
            app.Quit()
            app = None

            if os.path.exists(pdf_abs) and os.path.getsize(pdf_abs) > 0:
                logger.info(f"[{name}] Conversion OK → {pdf_abs}")
                return pdf_abs

        except Exception as e:
            logger.warning(f"[{name}] Failed: {e}")
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            time.sleep(0.1)

    raise RuntimeError("Both WPS and Word failed. Is either installed?")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Windows-native HTML → PDF (WPS/Word COM)")
    parser.add_argument("input", help="Input .html path")
    parser.add_argument("output", nargs="?", help="Output .pdf path")
    args = parser.parse_args()
    try:
        out = convert(args.input, args.output)
        print(out)
        sys.exit(0)
    except Exception as e:
        logger.error(str(e))
        sys.exit(1)
