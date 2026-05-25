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

def pack(unpacked_dir, output_pptx):
    # Pack unpacked_dir back into output_pptx
    with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, dirs, files in os.walk(unpacked_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Calculate path within zip
                arcname = os.path.relpath(file_path, unpacked_dir)
                zip_ref.write(file_path, arcname)

if __name__ == "__main__":
    configure_utf8_output()
    if len(sys.argv) < 3:
        print("Usage: python pack_pptx.py <unpacked_dir> <output_pptx>")
        sys.exit(1)
        
    pack(sys.argv[1], sys.argv[2])
