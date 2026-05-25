import sys
import os
import re
from pathlib import Path
from lxml import etree

# Access pptx_utils
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))
from pptx_utils import configure_utf8_output, NS_P, NS_A

def audit_xml(unpacked_dir):
    unpacked_dir = Path(unpacked_dir)
    slides_dir = unpacked_dir / "ppt" / "slides"
    
    # Robustly find all slide XMLs
    slide_files = list(slides_dir.glob("slide*.xml"))
    
    # Sort files naturally by numeric value found in name
    def slide_sort_key(p):
        nums = re.findall(r"\d+", p.name)
        return int(nums[0]) if nums else 0
    
    slide_files.sort(key=slide_sort_key)
    
    issues = []
    critical_issues = []
    
    # Common placeholder texts to flag
    PLACEHOLDER_REGEX = [
        r"\[.*?\]",           # [Title], [Text]
        r"\{\{.*?\}\}",       # {{Name}}
        r"Click to add",      # PPT default
        r"Lorem ipsum",       # Dummy text
        r"Enter text"
    ]
    
    print(f"\n=== Stage 2 Structural Audit Report ===")
    
    for slide_path in slide_files:
        # Extract ID or use filename as slide ID
        match = re.search(r"slide(?:_gen_)?(\d+)", slide_path.name)
        slide_id = match.group(1) if match else slide_path.stem
        try:
            tree = etree.parse(str(slide_path))
            root = tree.getroot()
        except Exception as e:
            critical_issues.append(f"Malformed slide XML: {slide_path.name} ({e})")
            continue
        
        slide_issues = []
        
        for sp in root.iter(f"{{{NS_P}}}sp"):
            cNvPr = sp.find(f".//{{{NS_P}}}cNvPr")
            name = cNvPr.get("name", "Unknown") if cNvPr is not None else "Unknown"
            
            # 1. Check for residual placeholder text
            texts = [t.text for t in sp.iter(f"{{{NS_A}}}t") if t.text]
            full_text = " ".join(texts)
            
            for pattern in PLACEHOLDER_REGEX:
                if re.search(pattern, full_text, re.IGNORECASE):
                    slide_issues.append(f"Residual Placeholder: '{full_text}' found in shape '{name}'")
                    break
            
            # 2. Text Density / Overflow Check (Simple Heuristic)
            xfrm = sp.find(f".//{{{NS_A}}}xfrm")
            if xfrm is not None:
                ext = xfrm.find(f"{{{NS_A}}}ext")
                if ext is not None:
                    try:
                        cx = int(ext.get("cx")) / 12700.0 # pt
                        cy = int(ext.get("cy")) / 12700.0 # pt
                    except Exception:
                        critical_issues.append(f"Invalid geometry values in {slide_path.name} shape '{name}'")
                        continue
                    
                    if cx > 0 and cy > 0:
                        char_count = len(full_text)
                        # Assume 12pt font avg, ~6pt width
                        # density = chars / (area / font_area)
                        density = char_count / ((cx * cy) / 144) 
                        if density > 15.0 and char_count > 50: # Arbitrary threshold for "too much text"
                            slide_issues.append(f"Potential Overflow: High text density ({density:.1f}) in '{name}'")
        
        if slide_issues:
            print(f"\n## Slide {slide_id}")
            for issue in slide_issues:
                print(f"  [!] {issue}")
                issues.append(issue)
            
    if critical_issues:
        print("\n[FAIL] Critical structural issues found:")
        for issue in critical_issues:
            print(f"  [X] {issue}")
        sys.exit(1)

    if not issues:
        print("\n[PASS] No obvious structural or content issues found in XML.")
    else:
        print(f"\n[FAIL] Found {len(issues)} potential issues. Please review.")

if __name__ == "__main__":
    configure_utf8_output()
    if len(sys.argv) < 2:
        print("Usage: python audit_injection.py <unpacked_dir>")
        sys.exit(1)
    audit_xml(sys.argv[1])
