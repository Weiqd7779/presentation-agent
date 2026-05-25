"""
extract_placeholders_v2.py — Parse PPTX XML to list all available text boxes.
Improved Version: Filters out empty textboxes and skips empty slides.
"""

import sys
import os
import re
from pathlib import Path

# Provide access to pptx_utils
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import configure_utf8_output, list_pictures, list_textboxes


SEMANTIC_NAME_RE = re.compile(
    r"^(S\d{2}_[A-Z0-9]+(?:_[A-Z0-9]+)*|[A-Za-z][A-Za-z0-9]*(?:_[A-Za-z0-9]+)*)$"
)
GENERIC_NAME_RE = re.compile(r"^(TextBox|Text|Picture|Image|Shape|Rectangle|Oval|Group)\s*\d*$", re.IGNORECASE)


def is_generic_name(name: str) -> bool:
    return bool(GENERIC_NAME_RE.match(name.strip()))


def is_semantic_name(name: str) -> bool:
    name = name.strip()
    return bool(SEMANTIC_NAME_RE.match(name)) and not is_generic_name(name)

def extract_all(unpacked_dir):
    unpacked_path = Path(unpacked_dir)
    slides = list(unpacked_path.glob("ppt/slides/slide*.xml"))
    
    # Sort slides correctly by their numerical slide number
    def extract_slide_num(p):
        match = re.search(r"slide(\d+)\.xml", p.name)
        return int(match.group(1)) if match else 9999
        
    slides.sort(key=extract_slide_num)
    
    print(f"--- Extracted Placeholders from {unpacked_dir} ---")
    print("Use these EXACT 'Name' values in your CONTENT_MAP.\n")
    
    found_any = False
    for slide_path in slides:
        num = extract_slide_num(slide_path)
        all_boxes = list_textboxes(slide_path)
        all_pictures = list_pictures(slide_path)
        
        # Filter for non-empty ones
        valid_boxes = []
        for box in all_boxes:
            text = box.get("current_text", "").strip()
            # Only keep if there's actual alphanumeric text or meaningful symbols
            # Skips things like single dots, spaces, or placeholder characters if needed
            if text:
                valid_boxes.append(box)
        
        if not valid_boxes and not all_pictures:
            continue
            
        found_any = True
        print(f"## Slide {num}")
        if valid_boxes:
            print("  Text boxes:")
            for box in valid_boxes:
                print(f"  - Name: \"{box['name']}\"")
                print(f"    Current Text: {box['current_text']}")
                if not is_semantic_name(box["name"]):
                    print("    Name Warning: generic or non-semantic shape name; consider renaming for stable mapping.")
        if all_pictures:
            print("  Pictures:")
            for picture in all_pictures:
                print(f"  - Picture Name: \"{picture['name']}\"")
                if picture.get("placeholder_idx"):
                    print(f"    Placeholder Idx: {picture['placeholder_idx']}")
                if not is_semantic_name(picture["name"]):
                    print("    Name Warning: generic or non-semantic picture name; consider renaming for stable image_map.")
        print()

    if not found_any:
        print("  (No non-empty textboxes found in the entire presentation)\n")

if __name__ == "__main__":
    configure_utf8_output()
    if len(sys.argv) < 2:
        print("Usage: python extract_placeholders_v2.py <unpacked_dir>")
        sys.exit(1)
        
    unpacked_path = sys.argv[1]
    if not os.path.isdir(unpacked_path):
        print(f"Error: Directory not found: {unpacked_path}")
        sys.exit(1)
        
    extract_all(unpacked_path)
