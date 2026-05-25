import sys
import zipfile
import os
from pathlib import Path

# Provide access to pptx_utils
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

try:
    from pptx_utils import configure_utf8_output
except ImportError:
    def configure_utf8_output(): pass

def unpack(pptx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

if __name__ == "__main__":
    configure_utf8_output()
    if len(sys.argv) < 3:
        print("Usage: python unpack_pptx.py <pptx_path> <output_dir>")
        sys.exit(1)
    unpack(sys.argv[1], sys.argv[2])
