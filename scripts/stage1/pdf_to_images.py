import fitz
import sys
import os
from pathlib import Path

# Provide access to pptx_utils
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

try:
    from pptx_utils import configure_utf8_output
except ImportError:
    def configure_utf8_output(): pass

def convert_pdf_to_images(pdf_path, output_dir, prefix):
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Increase resolution
        output_path = os.path.join(output_dir, f"{prefix}-{i+1:03d}.jpg")
        pix.save(output_path)
    doc.close()

if __name__ == "__main__":
    configure_utf8_output()
    pdf_path = sys.argv[1]
    output_dir = sys.argv[2]
    prefix = sys.argv[3]
    convert_pdf_to_images(pdf_path, output_dir, prefix)
