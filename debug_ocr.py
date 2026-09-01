"""
GMP Automation - Offline OCR debug tool
Menampilkan teks mentah hasil Offline OCR untuk satu file PDF.

Usage:
    python debug_ocr.py <path_pdf>
"""

import sys
from deepseek_ocr.client import ocr_pdf

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    pdf_path = sys.argv[1]
    pages = ocr_pdf(pdf_path)
    for i, page in enumerate(pages, start=1):
        print(f'==================== PAGE {i} ====================')
        print(page)
