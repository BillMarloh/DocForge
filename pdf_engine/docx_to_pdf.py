"""
Windows-native DOCX → PDF converter via WPS / Word COM automation.

Strategy: WPS first, Word fallback. Silent background mode. No Chrome/Puppeteer needed.

Usage:
  python docx_to_pdf.py input.docx [output.pdf]
  Exit code 0 = success, 1 = failure
"""

import sys
import os
import logging
import time

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
logger = logging.getLogger("DocxToPdf")


class DocxToPdfConverter:
    """WPS/Word dual-engine converter with forced cleanup"""

    # WPS Format: 17 = wdFormatPDF
    # Word Format: 17 = wdFormatPDF
    PDF_FORMAT = 17

    def __init__(self, timeout: int = 60):
        self.timeout = timeout
        self.app = None
        self.doc = None

    def convert(self, docx_path: str, pdf_path: str = None) -> str:
        """
        Convert DOCX to PDF. Tries WPS first, falls back to Word.

        Returns the output PDF path.
        Raises RuntimeError if both engines fail.
        """
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"DOCX not found: {docx_path}")

        if pdf_path is None:
            pdf_path = docx_path.rsplit(".", 1)[0] + ".pdf"

        docx_abs = os.path.abspath(docx_path)
        pdf_abs = os.path.abspath(pdf_path)

        # Try WPS first
        logger.info(f"Trying WPS: {docx_abs} → {pdf_abs}")
        if self._convert_with(docx_abs, pdf_abs, "Wps.Application", "WPS"):
            return pdf_abs

        # Fall back to Word
        logger.info(f"Trying Word: {docx_abs} → {pdf_abs}")
        if self._convert_with(docx_abs, pdf_abs, "Word.Application", "Word"):
            return pdf_abs

        raise RuntimeError("Both WPS and Word failed. Is either installed?")

    def _convert_with(self, docx_abs: str, pdf_abs: str,
                      prog_id: str, name: str) -> bool:
        """Attempt conversion with a specific COM application."""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

        try:
            from win32com.client import Dispatch
            self.app = Dispatch(prog_id)
            self.app.Visible = False
            self.app.DisplayAlerts = False

            # Suppress any "save as" prompts
            try:
                self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass

            self.doc = self.app.Documents.Open(docx_abs, ReadOnly=True)

            # Delete existing output to avoid overwrite prompt
            if os.path.exists(pdf_abs):
                try:
                    os.remove(pdf_abs)
                except Exception:
                    pass

            self.doc.SaveAs(pdf_abs, FileFormat=self.PDF_FORMAT)
            self.doc.Close()
            self.doc = None

            self.app.Quit()
            self.app = None

            if os.path.exists(pdf_abs) and os.path.getsize(pdf_abs) > 0:
                logger.info(f"[{name}] Conversion OK → {pdf_abs}")
                return True
            else:
                logger.warning(f"[{name}] Output file missing or empty")
                return False

        except Exception as e:
            logger.warning(f"[{name}] Failed: {e}")
            return False

        finally:
            self._force_cleanup()

    def _force_cleanup(self):
        """Ensure COM objects are released no matter what."""
        if self.doc is not None:
            try:
                self.doc.Close(SaveChanges=False)
            except Exception:
                pass
            self.doc = None

        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None

        # Release COM reference
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

        # Allow COM to release before process exits
        time.sleep(0.1)


# ============================================================
# CLI
# ============================================================
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Windows-native DOCX → PDF (WPS/Word COM)")
    parser.add_argument("input", help="Input .docx path")
    parser.add_argument("output", nargs="?", help="Output .pdf path (default: same name)")
    parser.add_argument("--timeout", type=int, default=60, help="Seconds before timeout")
    args = parser.parse_args()

    try:
        converter = DocxToPdfConverter(timeout=args.timeout)
        out = converter.convert(args.input, args.output)
        print(out)
        sys.exit(0)
    except Exception as e:
        logger.error(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
