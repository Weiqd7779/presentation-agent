"""
inject_content_template.py
Template script for Stage 2 Content Injection.
Copy this file to your output directory as `inject_content.py` and fill in the 
SLIDE_ORDER and CONTENT_MAP.
"""
import sys
import os
from pathlib import Path

# Add skill root to path for pptx_utils import
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import (
    configure_utf8_output,
    load_slide, save_slide,
    set_textbox_text,
    reorder_slides
)

# 1. Windows UTF-8 initialization
configure_utf8_output()

# 2. Slide Structure Definition
# List of 1-based indices from the ORIGINAL template.
# Repetition = duplication; omission = deletion.
# Set to None to use all slides in original order.
SLIDE_ORDER = None  # Example: [1, 5, 5, 2]

# 3. Content Mapping Definition
# Key: 1-based slide index in the FINAL output presentation.
# Value: Dictionary of { "TextBox Name or Placeholder ID": "Text Content" }
CONTENT_MAP = {
    1: {
        "TextBox 1": "My Presentation Title",
        "TextBox 2": "Subtitle or Department Name"
    },
    # 2: { ... },
}

def inject(unpacked_dir):
    """
    Perform slide reordering and content injection.
    
    Args:
        unpacked_dir: Path to the directory containing unpacked PPTX files.
    """
    unpacked_path = Path(unpacked_dir)
    
    # --- Phase A: Structure Modification ---
    if SLIDE_ORDER:
        print(f"==> Reordering slides: {SLIDE_ORDER}")
        reorder_slides(unpacked_path, SLIDE_ORDER)
    else:
        print("==> Keeping original slide order.")

    # --- Phase B: Content Injection ---
    print("==> Injecting content into slides...")
    for slide_num, boxes in CONTENT_MAP.items():
        print(f"  Processing Slide {slide_num}...")
        try:
            tree, root = load_slide(unpacked_path, slide_num)
        except Exception as e:
            print(f"  [ERROR] Cannot load slide {slide_num}: {e}")
            continue
            
        for box_name, text in boxes.items():
            # Check if text is a list (for multiple paragraphs/bullets)
            if isinstance(text, list):
                if text:
                    set_textbox_text(root, box_name, text[0], extra_paras=text[1:])
                else:
                    set_textbox_text(root, box_name, "")
            else:
                set_textbox_text(root, box_name, str(text))
        
        save_slide(tree, unpacked_path, slide_num)
        
    print("Injection process complete.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python inject_content.py <unpacked_dir>")
        sys.exit(1)
        
    unpacked_path = sys.argv[1]
    inject(unpacked_path)
