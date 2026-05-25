"""
pptx_utils.py — Shared utilities for presentation-agent build scripts.

Usage in any build script:
    import sys
    from pathlib import Path
    
    SKILL_DIR = Path(__file__).resolve().parent.parent.parent
    sys.path.insert(0, str(SKILL_DIR))
    
    from pptx_utils import (
        configure_utf8_output, NS_P, NS_A, NS_R, NS_XML,
        load_slide, save_slide,
        set_textbox_text, list_textboxes,
        delete_shape_by_name,
        detect_template_type, is_canva_export,
        pdf_to_images,
        unpack_pptx, pack_pptx, convert_pptx_to_pdf,
        reorder_slides,
    )
"""

import sys
import copy
import re
import zipfile
import math
from pathlib import Path
from typing import Optional, Union

# ── Windows UTF-8 fix ──────────────────────────────────────────────────────
# Must be called at the top of any script that prints emoji on Windows.

def configure_utf8_output() -> None:
    """
    Force stdout/stderr to UTF-8 on Windows to prevent Mojibake.
    
    This is especially critical on Chinese Windows locales (cp950/cp936) 
    where the default console encoding cannot handle various Unicode symbols.
    """
    import os
    
    # Always set environment variable for child processes
    os.environ["PYTHONIOENCODING"] = "utf-8"

    # Reconfigure streams if they exist and support reconfiguration (Python 3.7+)
    for stream_name in ("stdout", "stderr", "stdin"):
        stream = getattr(sys, stream_name, None)
        if stream and hasattr(stream, "reconfigure"):
            try:
                # Check current encoding; if it's not utf-8, force it.
                # On Windows, 'utf-8' might sometimes be 'cp65001', which is also fine.
                curr_enc = getattr(stream, "encoding", "").lower()
                if curr_enc not in ("utf-8", "utf8", "cp65001"):
                    stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass  # Fallback if stream is redirected or doesn't support it


# ── XML namespace constants ────────────────────────────────────────────────

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"


# ── Slide XML load / save ──────────────────────────────────────────────────

def load_slide(unpacked_dir, slide_number: int):
    """
    Load a slide XML from an unpacked PPTX directory.

    Returns:
        (ElementTree, root_element) tuple
    """
    from lxml import etree
    path = Path(unpacked_dir) / "ppt" / "slides" / f"slide{slide_number}.xml"
    tree = etree.parse(str(path))
    return tree, tree.getroot()


def save_slide(tree, unpacked_dir, slide_number: int) -> None:
    """Save a slide XML tree back to the unpacked directory."""
    path = Path(unpacked_dir) / "ppt" / "slides" / f"slide{slide_number}.xml"
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    print(f"  [OK] slide{slide_number}.xml saved")


# ── TextBox inventory ──────────────────────────────────────────────────────

def list_textboxes(slide_xml_path) -> list:
    """
    List all named TextBoxes in a slide and their current text content.

    Useful before editing to build the name → content mapping.

    Returns:
        List of {"name": str, "current_text": str}
    """
    from lxml import etree
    root = etree.parse(str(slide_xml_path)).getroot()
    boxes = []
    for sp in root.iter(f"{{{NS_P}}}sp"):
        cNvPr = sp.find(f".//{{{NS_P}}}cNvPr")
        name = cNvPr.get("name", "?") if cNvPr is not None else "?"
        texts = [t.text for t in sp.iter(f"{{{NS_A}}}t")
                 if t.text and t.text.strip()]
        boxes.append({"name": name, "current_text": " ".join(texts)})
    return boxes


def list_pictures(slide_xml_path) -> list:
    """
    List all named pictures in a slide.

    Useful for building image_map entries. Returns:
        List of {"name": str, "placeholder_idx": str | None}
    """
    from lxml import etree
    root = etree.parse(str(slide_xml_path)).getroot()
    pictures = []
    for pic in root.iter(f"{{{NS_P}}}pic"):
        cNvPr = pic.find(f".//{{{NS_P}}}cNvPr")
        name = cNvPr.get("name", "?") if cNvPr is not None else "?"
        ph = pic.find(f".//{{{NS_P}}}ph")
        idx = ph.get("idx") if ph is not None else None
        pictures.append({"name": name, "placeholder_idx": idx})
    return pictures


# ── TextBox text replacement (Path A-TX) ──────────────────────────────────

def set_textbox_text(
    slide_root,
    box_name: str,
    new_text: str,
    extra_paras: Optional[list] = None,
    **kwargs
) -> float:
    """
    Replace ALL text in a named TextBox while preserving run + paragraph formatting.

    Rules:
    - Never adds new shapes or TextBoxes.
    - Preserves the <a:rPr> of the first run (font, size, bold, colour).
    - Preserves the <a:pPr> of the first paragraph (alignment, spacing).
    - Each string in extra_paras becomes a separate <a:p> (for bullets).

    Args:
        slide_root:  Root lxml element from load_slide()
        box_name:    Exact name of the TextBox (e.g. "TextBox 7")
        new_text:    Text for the first paragraph
        extra_paras: Optional list of strings for additional paragraphs

    Returns:
        True if found and updated, False if not found (prints [WARN]).
    """
    from lxml import etree

    for sp in slide_root.iter(f"{{{NS_P}}}sp"):
        cNvPr = sp.find(f".//{{{NS_P}}}cNvPr")
        name = cNvPr.get("name") if cNvPr is not None else None
        
        # Check if it has a placeholder idx
        ph = sp.find(f".//{{{NS_P}}}nvPr/{{{NS_P}}}ph")
        if ph is None:
            # Maybe inside grpSp or direct
            ph = sp.find(f".//{{{NS_P}}}ph")
        idx = ph.get("idx") if ph is not None else None
        
        # Match either by exact Name or by Placeholder ID
        if str(name) != str(box_name) and str(idx) != str(box_name):
            continue

        # Guard: never modify decorative placeholders regardless of what was requested.
        # These are template-managed elements — touching them corrupts slide numbering,
        # footer text, and date fields that PowerPoint controls automatically.
        _DECORATIVE_PH_TYPES = {"dt", "sldNum", "ftr", "hdr", "date", "footer"}
        if ph is not None and ph.get("type", "") in _DECORATIVE_PH_TYPES and not kwargs.get("allow_decorative"):
            print(f"  [GUARD] '{box_name}' is a decorative placeholder ({ph.get('type')}) — skipping")
            return True

        txBody = sp.find(f".//{{{NS_P}}}txBody")
        if txBody is None:
            return True  # shape found, but no text body — skip silently

        paras = txBody.findall(f"{{{NS_A}}}p")
        if not paras:
            return True

        # Extract formatting from the first run of the first paragraph
        first_rPr = None
        for r in paras[0].iter(f"{{{NS_A}}}r"):
            rpr = r.find(f"{{{NS_A}}}rPr")
            if rpr is not None:
                first_rPr = copy.deepcopy(rpr)
                break

        first_pPr = paras[0].find(f"{{{NS_A}}}pPr")
        first_pPr = copy.deepcopy(first_pPr) if first_pPr is not None else None

        # --- SCALING LOGIC ---
        # 1. Get container dimensions (EMUs)
        xfrm = sp.find(f".//{{{NS_A}}}xfrm")
        ext = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
        cx = int(ext.get("cx")) if ext is not None else 0
        cy = int(ext.get("cy")) if ext is not None else 0
        
        # EMU to PT conversion (1 pt = 12700 EMU)
        box_w_pt = cx / 12700.0
        box_h_pt = cy / 12700.0
        
        # 2. Extract original font size (sz)
        orig_sz = 2400  # Default 24pt
        if first_rPr is not None:
            sz_attr = first_rPr.get("sz")
            if sz_attr:
                orig_sz = int(sz_attr)

        # 3. Heuristic scaling
        # Total text lines estimation
        all_lines = [new_text] + (extra_paras or [])
        char_count = sum(len(l) for l in all_lines)
        
        if char_count > 0 and box_w_pt > 0 and box_h_pt > 0:
            font_size_pt = orig_sz / 100.0
            est_width_per_char = font_size_pt * 0.55 
            chars_per_line = max(1, box_w_pt / est_width_per_char)
            
            # Sum up estimated lines for each paragraph
            est_lines = sum(math.ceil(len(l) / chars_per_line) if l else 1 for l in all_lines)
            est_height = est_lines * (font_size_pt * 1.3) # 1.3 for line spacing
            
            if est_height > box_h_pt:
                # Scale down font size to fit height
                scale = box_h_pt / est_height
                # Add a small safety buffer (0.95)
                scale = scale * 0.95
            else:
                scale = 1.0
        else:
            scale = 1.0
                
        # If an external scale or fixed font size is provided, override local calculation.
        fixed_font_size = kwargs.get("font_size")
        min_font_size = int(kwargs.get("min_font_size", 1000))
        max_font_size = kwargs.get("max_font_size")
        final_scale = kwargs.get("force_scale", scale)
        if fixed_font_size is not None:
            new_sz = int(float(fixed_font_size) * 100)
            if first_rPr is not None and not kwargs.get("dry_run"):
                first_rPr.set("sz", str(new_sz))
            if new_sz != orig_sz:
                print(f"  [font] '{box_name}' font {orig_sz}->{new_sz} (fixed)")
        else:
            if final_scale < 1.0 or max_font_size is not None:
                new_sz = max(min_font_size, int(orig_sz * final_scale))
                if max_font_size is not None:
                    new_sz = min(new_sz, int(float(max_font_size) * 100))
                if new_sz < orig_sz:
                    print(f"  [scale] '{box_name}' font {orig_sz}->{new_sz} (scale={final_scale:.2f})")
                    if first_rPr is not None and not kwargs.get("dry_run"):
                        first_rPr.set("sz", str(new_sz))

        # --- END SCALING LOGIC ---

        # Remove all existing paragraphs
        for p in paras:
            txBody.remove(p)

        def _make_para(text: str):
            if kwargs.get("dry_run"):
                return None
            p_el = etree.SubElement(txBody, f"{{{NS_A}}}p")
            if first_pPr is not None:
                p_el.insert(0, copy.deepcopy(first_pPr))
            if text:
                r_el = etree.SubElement(p_el, f"{{{NS_A}}}r")
                if first_rPr is not None:
                    r_el.append(copy.deepcopy(first_rPr))
                t_el = etree.SubElement(r_el, f"{{{NS_A}}}t")
                t_el.text = text
                t_el.set(f"{{{NS_XML}}}space", "preserve")
            return p_el

        _make_para(new_text)
        for ep in (extra_paras or []):
            _make_para(ep)
        return final_scale

    print(f"  [WARN] TextBox '{box_name}' not found — skipping")
    return 1.0


def replace_picture_image(
    unpacked_dir,
    slide_number: int,
    picture_name: str,
    image_path,
    *,
    fill: str = "contain",
    background: str = "white",
) -> bool:
    """
    Replace an existing picture's media file while preserving the slide layout.

    The replacement is rendered into the existing media part's file extension and
    aspect ratio, so PowerPoint relationships remain stable and rerunnable.
    """
    from lxml import etree
    import defusedxml.minidom
    from PIL import Image

    unpacked_path = Path(unpacked_dir)
    slide_path = unpacked_path / "ppt" / "slides" / f"slide{slide_number}.xml"
    rels_path = unpacked_path / "ppt" / "slides" / "_rels" / f"slide{slide_number}.xml.rels"
    source_path = Path(image_path)

    if not source_path.exists():
        print(f"  [WARN] Image source not found: {source_path}")
        return False

    tree = etree.parse(str(slide_path))
    root = tree.getroot()
    target_pic = None
    for pic in root.iter(f"{{{NS_P}}}pic"):
        cNvPr = pic.find(f".//{{{NS_P}}}cNvPr")
        name = cNvPr.get("name") if cNvPr is not None else None
        ph = pic.find(f".//{{{NS_P}}}ph")
        idx = ph.get("idx") if ph is not None else None
        if str(name) == str(picture_name) or str(idx) == str(picture_name):
            target_pic = pic
            break

    if target_pic is None:
        print(f"  [WARN] Picture '{picture_name}' not found — skipping")
        return False

    blip = target_pic.find(f".//{{{NS_A}}}blip")
    if blip is None:
        print(f"  [WARN] Picture '{picture_name}' has no image relationship — skipping")
        return False

    rid = blip.get(f"{{{NS_R}}}embed")
    if not rid:
        print(f"  [WARN] Picture '{picture_name}' has no embedded image id — skipping")
        return False

    rels_dom = defusedxml.minidom.parse(str(rels_path))
    target = None
    for rel in rels_dom.getElementsByTagName("Relationship"):
        if rel.getAttribute("Id") == rid:
            target = rel.getAttribute("Target")
            break
    if not target:
        print(f"  [WARN] Image relationship '{rid}' not found — skipping")
        return False

    media_path = (slide_path.parent / target).resolve()
    xfrm = target_pic.find(f".//{{{NS_A}}}xfrm")
    ext = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
    if ext is not None:
        cx = max(1, int(ext.get("cx", "1")))
        cy = max(1, int(ext.get("cy", "1")))
        ratio = cx / cy
    else:
        ratio = 1.0

    out_w = 1600
    out_h = max(1, int(out_w / ratio))
    src = Image.open(source_path).convert("RGBA")
    canvas = Image.new("RGBA", (out_w, out_h), background)

    if fill == "cover":
        scale = max(out_w / src.width, out_h / src.height)
    else:
        scale = min(out_w / src.width, out_h / src.height)
    new_size = (max(1, int(src.width * scale)), max(1, int(src.height * scale)))
    resized = src.resize(new_size, Image.Resampling.LANCZOS)

    if fill == "cover":
        left = max(0, (resized.width - out_w) // 2)
        top = max(0, (resized.height - out_h) // 2)
        resized = resized.crop((left, top, left + out_w, top + out_h))
        canvas.alpha_composite(resized, (0, 0))
    else:
        canvas.alpha_composite(resized, ((out_w - resized.width) // 2, (out_h - resized.height) // 2))

    suffix = media_path.suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        canvas.convert("RGB").save(media_path, quality=95)
    else:
        canvas.save(media_path)
    print(f"  [image] '{picture_name}' replaced with {source_path.name}")
    return True


def _find_graphic_frame_by_name(slide_root, frame_name: str):
    for frame in slide_root.iter(f"{{{NS_P}}}graphicFrame"):
        cNvPr = frame.find(f".//{{{NS_P}}}cNvPr")
        name = cNvPr.get("name") if cNvPr is not None else None
        if str(name) == str(frame_name):
            return frame
    return None


def set_table_data(slide_root, table_name: str, options: dict) -> bool:
    """
    Replace text in a PowerPoint table (<p:graphicFrame>/<a:tbl>) and optionally
    resize the table grid. Existing cell formatting is preserved per cell.
    """
    from lxml import etree

    frame = _find_graphic_frame_by_name(slide_root, table_name)
    if frame is None:
        print(f"  [WARN] Table '{table_name}' not found — skipping")
        return False

    table = frame.find(f".//{{{NS_A}}}tbl")
    if table is None:
        print(f"  [WARN] Graphic frame '{table_name}' has no table — skipping")
        return False

    rows = table.findall(f"./{{{NS_A}}}tr")
    grid_cols = table.findall(f"./{{{NS_A}}}tblGrid/{{{NS_A}}}gridCol")

    new_rows = options.get("rows")
    if new_rows is not None:
        for row_index, row_values in enumerate(new_rows):
            if row_index >= len(rows):
                break
            cells = rows[row_index].findall(f"./{{{NS_A}}}tc")
            for col_index, value in enumerate(row_values):
                if col_index >= len(cells):
                    break
                _set_table_cell_text(cells[col_index], str(value), options)

    cells_map = options.get("cells", {})
    if isinstance(cells_map, dict):
        for key, value in cells_map.items():
            try:
                row_index, col_index = [int(part) for part in str(key).split(",", 1)]
            except ValueError:
                print(f"  [WARN] Invalid table cell key '{key}' for '{table_name}'")
                continue
            if row_index < 0 or row_index >= len(rows):
                continue
            cells = rows[row_index].findall(f"./{{{NS_A}}}tc")
            if col_index < 0 or col_index >= len(cells):
                continue
            _set_table_cell_text(cells[col_index], str(value), options)

    size = options.get("size")
    if isinstance(size, dict):
        _resize_table_frame(frame, rows, grid_cols, size)

    return True


def _set_table_cell_text(cell, text: str, options: dict) -> None:
    from lxml import etree

    tx_body = cell.find(f"./{{{NS_A}}}txBody")
    if tx_body is None:
        tx_body = etree.SubElement(cell, f"{{{NS_A}}}txBody")
        etree.SubElement(tx_body, f"{{{NS_A}}}bodyPr")
        etree.SubElement(tx_body, f"{{{NS_A}}}lstStyle")

    paragraphs = tx_body.findall(f"./{{{NS_A}}}p")
    if not paragraphs:
        paragraphs = [etree.SubElement(tx_body, f"{{{NS_A}}}p")]

    first_p = paragraphs[0]
    first_p_pr = first_p.find(f"./{{{NS_A}}}pPr")
    first_p_pr = copy.deepcopy(first_p_pr) if first_p_pr is not None else None
    first_r_pr = None
    for run in first_p.iter(f"{{{NS_A}}}r"):
        r_pr = run.find(f"./{{{NS_A}}}rPr")
        if r_pr is not None:
            first_r_pr = copy.deepcopy(r_pr)
            break

    font_size = options.get("font_size")
    if font_size is not None and first_r_pr is not None:
        first_r_pr.set("sz", str(int(float(font_size) * 100)))

    for paragraph in paragraphs:
        tx_body.remove(paragraph)

    for index, line in enumerate(text.split("\n") or [""]):
        p_el = etree.SubElement(tx_body, f"{{{NS_A}}}p")
        if first_p_pr is not None:
            p_el.insert(0, copy.deepcopy(first_p_pr))
        if line:
            r_el = etree.SubElement(p_el, f"{{{NS_A}}}r")
            if first_r_pr is not None:
                r_el.append(copy.deepcopy(first_r_pr))
            t_el = etree.SubElement(r_el, f"{{{NS_A}}}t")
            t_el.text = line
            t_el.set(f"{{{NS_XML}}}space", "preserve")


def _resize_table_frame(frame, rows, grid_cols, size: dict) -> None:
    unit = size.get("unit", "in")
    if unit != "in":
        raise ValueError("table size currently supports unit='in' only")
    width = size.get("w")
    height = size.get("h")
    if width is None and height is None:
        return

    xfrm = frame.find(f"./{{{NS_P}}}xfrm")
    ext = xfrm.find(f"{{{NS_A}}}ext") if xfrm is not None else None
    if ext is None:
        return

    if width is not None:
        target_w = int(float(width) * 914400)
        old_total = sum(int(col.get("w", "0")) for col in grid_cols) or int(ext.get("cx", "1"))
        for col in grid_cols:
            old_w = int(col.get("w", "0"))
            col.set("w", str(max(1, int(old_w * target_w / old_total))))
        ext.set("cx", str(target_w))

    if height is not None:
        target_h = int(float(height) * 914400)
        old_total = sum(int(row.get("h", "0")) for row in rows) or int(ext.get("cy", "1"))
        for row in rows:
            old_h = int(row.get("h", "0"))
            row.set("h", str(max(1, int(old_h * target_h / old_total))))
        ext.set("cy", str(target_h))

# ── Shape Deletion ─────────────────────────────────────────────────────────

def delete_shape_by_name(slide_root, target_name: str) -> bool:
    """
    Delete a shape (<p:sp> or <p:pic>) by its name attribute in <p:cNvPr>.
    
    Args:
        slide_root:  Root lxml element from load_slide()
        target_name: Exact name of the shape (e.g., "Picture 2")
        
    Returns:
        True if found and deleted, False otherwise.
    """
    # Look for sp (shape) and pic (picture)
    for tag in [f"{{{NS_P}}}sp", f"{{{NS_P}}}pic", f"{{{NS_P}}}grpSp"]:
        for el in slide_root.iter(tag):
            cNvPr = el.find(f".//{{{NS_P}}}cNvPr")
            name = cNvPr.get("name") if cNvPr is not None else None
            
            ph = el.find(f".//{{{NS_P}}}nvPr/{{{NS_P}}}ph")
            if ph is None:
                ph = el.find(f".//{{{NS_P}}}ph")
            idx = ph.get("idx") if ph is not None else None
            
            if str(name) == str(target_name) or str(idx) == str(target_name):
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
                    return True
    
    print(f"  [WARN] Shape '{target_name}' not found for deletion — skipping")
    return False


# ── Template type detection ────────────────────────────────────────────────

def is_canva_export(pptx_path) -> bool:
    """
    Check whether a PPTX was exported from Canva.

    Canva writes its application name into docProps/app.xml.
    This is a stronger signal than the absence of <p:ph> tags alone.

    Returns:
        True if the file was generated by Canva, False otherwise.
    """
    with zipfile.ZipFile(pptx_path) as zf:
        if "docProps/app.xml" in zf.namelist():
            content = zf.read("docProps/app.xml").decode("utf-8", errors="ignore")
            if "canva" in content.lower():
                return True
    return False


def detect_template_type(pptx_path) -> str:
    """
    Detect whether a PPTX template uses formal placeholders or TextBoxes.

    Canva, Freepik, Slidesgo exports typically use only TextBoxes.
    Native PowerPoint templates use formal <p:ph> placeholders.

    Checks up to the first 3 slides (not just slide 1) to avoid
    misclassifying templates whose cover slide happens to be image-only.

    Returns:
        'PLACEHOLDER_TEMPLATE' or 'TEXTBOX_TEMPLATE'
    """
    with zipfile.ZipFile(pptx_path) as zf:
        names = zf.namelist()
        slide_files = sorted(
            f for f in names if re.match(r"ppt/slides/slide\d+\.xml$", f)
        )[:3]  # check up to 3 slides
        if not slide_files:
            return "TEXTBOX_TEMPLATE"
        for sf in slide_files:
            content = zf.read(sf).decode("utf-8")
            if re.search(r"<p:ph\b", content):
                return "PLACEHOLDER_TEMPLATE"
    return "TEXTBOX_TEMPLATE"


# ── PDF → JPEG conversion ──────────────────────────────────────────────────

def pdf_to_images(pdf_path, output_dir, prefix: str = "slide", dpi: int = 150) -> list:
    """
    Convert a PDF to per-page JPEG images using pymupdf.

    No external binary (pdftoppm / Poppler) required.

    Args:
        pdf_path:   Path to PDF file
        output_dir: Directory where JPEGs will be saved (created if needed)
        prefix:     Filename prefix (default "slide" → slide_001.jpg, slide_002.jpg …)
        dpi:        Resolution (default 150; use 200+ for QA, 100 for thumbnails)

    Returns:
        Sorted list of Path objects for the generated JPEG files

    Raises:
        RuntimeError: if pymupdf is not installed
    """
    try:
        import fitz
    except ImportError:
        raise RuntimeError(
            "pymupdf is not installed.\n"
            "Install it with:  uv add pymupdf"
        )

    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    doc = fitz.open(str(pdf_path))
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        out = output_dir / f"{prefix}_{i + 1:03d}.jpg"
        pix.save(str(out))
        paths.append(out)
    doc.close()
    return sorted(paths)


# ── Pack / unpack helpers ──────────────────────────────────────────────────

def unpack_pptx(pptx_path, output_dir, skill_pptx_scripts: str) -> None:
    """Unpack a PPTX to a directory using the skills-pptx unpack.py script."""
    import subprocess
    result = subprocess.run(
        [sys.executable,
         str(Path(skill_pptx_scripts) / "office" / "unpack.py"),
         str(pptx_path), str(output_dir)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"unpack failed: {result.stderr}")
    print(f"  [OK] unpacked to {output_dir}")


def pack_pptx(unpacked_dir, output_pptx, original_pptx, skill_pptx_scripts: str) -> None:
    """Pack an unpacked directory back to PPTX using skills-pptx pack.py script."""
    import subprocess
    result = subprocess.run(
        [sys.executable,
         str(Path(skill_pptx_scripts) / "office" / "pack.py"),
         str(unpacked_dir), str(output_pptx),
         "--original", str(original_pptx)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pack failed: {result.stderr}")
    print(f"  [OK] packed to {output_pptx}")


def convert_pptx_to_pdf(pptx_path: Union[str, Path], output_dir: Union[str, Path], _unused_scripts_dir=None) -> Path:
    """
    Convert PPTX to PDF using LibreOffice (soffice).
    
    Robustly finds soffice on Windows if not in PATH.
    """
    import subprocess
    import shutil
    from pathlib import Path

    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)

    # 1. Find soffice binary
    soffice_cmd = "soffice"
    if not shutil.which(soffice_cmd):
        # Common Windows installation path
        windows_path = Path("C:/Program Files/LibreOffice/program/soffice.exe")
        if windows_path.exists():
            soffice_cmd = str(windows_path)
        else:
            raise RuntimeError(
                "LibreOffice (soffice) not found in PATH or at C:/Program Files/LibreOffice/.\n"
                "Please install LibreOffice to enable PDF/Image export."
            )

    # 2. Run conversion
    print(f"  [soffice] Converting {pptx_path.name} to PDF...")
    result = subprocess.run(
        [soffice_cmd, "--headless", "--convert-to", "pdf",
         str(pptx_path), "--outdir", str(output_dir)],
        capture_output=True, text=True, encoding="utf-8"
    )
    
    if result.returncode != 0:
        raise RuntimeError(f"soffice conversion failed (code {result.returncode}): {result.stderr}")

    pdf_path = output_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"PDF not created at expected location: {pdf_path}")
    
    return pdf_path


# ── Slide Reordering / Deletion ────────────────────────────────────────────

def _serialize_xml_bytes(dom) -> bytes:
    """Serialize minidom DOM to bytes with standalone='yes' in the XML declaration."""
    raw = dom.toxml(encoding="UTF-8")
    return raw.replace(
        b'<?xml version="1.0" encoding="UTF-8"?>',
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        1,
    )


def reorder_slides(unpacked_dir, slide_order: list) -> None:
    """
    Physically reorder, duplicate, or delete slides in an unpacked PPTX.

    Operates at the file-system level so that:
    - Duplicated slides become fully independent copies (no shared-file corruption).
    - Deleted slides are removed from disk.
    - After this call slide files on disk are slide1.xml … slideN.xml in the new
      sequence — CONTENT_MAP should therefore use 1-based positions in the FINAL
      output order (applied after this call).

    Uses minidom for .rels and [Content_Types].xml to preserve namespace prefixes.
    Uses lxml for presentation.xml (consistent with the rest of pptx_utils).

    Args:
        unpacked_dir: Path to the unpacked PPTX directory.
        slide_order:  List of 1-based slide indices from the ORIGINAL template.
                      Repetition = duplication; omission = deletion.
                      Example: [1, 3, 3, 2] keeps slide 1, duplicates slide 3 twice,
                      keeps slide 2; slide 4+ are deleted.
    """
    import shutil
    import defusedxml.minidom
    from lxml import etree

    unpacked_path   = Path(unpacked_dir)
    slides_dir      = unpacked_path / "ppt" / "slides"
    slides_rels_dir = slides_dir / "_rels"
    pres_xml_path   = unpacked_path / "ppt" / "presentation.xml"
    pres_rels_path  = unpacked_path / "ppt" / "_rels" / "presentation.xml.rels"
    ct_path         = unpacked_path / "[Content_Types].xml"

    if not pres_xml_path.exists():
        print(f"  [ERROR] presentation.xml not found at {pres_xml_path}")
        return

    # ── Phase 1: Parse presentation.xml (lxml) ────────────────────────────
    tree = etree.parse(str(pres_xml_path))
    root = tree.getroot()

    sldIdLst = root.find(f".//{{{NS_P}}}sldIdLst")
    if sldIdLst is None:
        print("  [WARN] <p:sldIdLst> not found — cannot reorder")
        return

    original_sldIds = list(sldIdLst.findall(f"{{{NS_P}}}sldId"))
    n_original = len(original_sldIds)
    print(f"  [reorder] Original slide count: {n_original}")

    # ── Phase 2: Parse presentation.xml.rels (minidom) ────────────────────
    # Build rId → slide filename map and collect non-slide relationships.
    rels_dom = defusedxml.minidom.parse(str(pres_rels_path))
    rid_to_file: dict[str, str] = {}   # e.g. {"rId2": "slide1.xml"}
    non_slide_rels = []                # minidom nodes for non-slide entries

    SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"

    for rel in rels_dom.getElementsByTagName("Relationship"):
        target   = rel.getAttribute("Target")
        rid      = rel.getAttribute("Id")
        rel_type = rel.getAttribute("Type")
        if rel_type == SLIDE_REL_TYPE and target.startswith("slides/"):
            rid_to_file[rid] = target.split("/")[-1]   # "slide1.xml"
        else:
            non_slide_rels.append(rel)

    # Map each original 0-based index → source filename on disk.
    idx_to_source: dict[int, str] = {}
    for i, sld_el in enumerate(original_sldIds):
        rid = sld_el.get(f"{{{NS_R}}}id")
        if rid and rid in rid_to_file:
            idx_to_source[i] = rid_to_file[rid]
        else:
            print(f"  [WARN] rId '{rid}' for original slide {i+1} not in rels — will skip")

    # Validate requested indices.
    valid_order: list[int] = []
    for idx_1based in slide_order:
        try:
            idx_0based = int(idx_1based) - 1
            if 0 <= idx_0based < n_original:
                valid_order.append(idx_0based)
            else:
                print(f"  [WARN] Slide index {idx_1based} out of range (1-{n_original}) — skipping")
        except (ValueError, TypeError):
            print(f"  [WARN] Invalid index '{idx_1based}' — skipping")

    if not valid_order:
        print("  [WARN] No valid slide indices — aborting reorder")
        return

    # ── Phase 3: Copy source files to temp names ──────────────────────────
    # Temp names use "_reorder_tmp_" prefix — they cannot clash with slide\d+.xml.
    # Each position gets its own copy, so duplicates become independent files.
    # When a slide with diagrams/SmartArt is duplicated (same source appears 2+
    # times), we deep-copy all diagram files so both slides own independent copies.
    slides_rels_dir.mkdir(parents=True, exist_ok=True)
    tmp_pairs: list[tuple[Path, Path]] = []  # (tmp_slide, tmp_rels)

    # Track which source slides we've already processed to detect duplicates.
    seen_sources: dict[int, int] = {}   # src_idx → count of times seen so far

    # Relationship types that require file-level deep copy when duplicating.
    DIAGRAM_REL_TYPES = {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramDrawing",
        # Microsoft-specific namespace for SmartArt drawing canvas
        "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
    }

    def _deep_copy_slide_resources(tmp_rels_path: Path, unpacked: Path) -> None:
        """For a duplicated slide's .rels, copy every diagram file to a new unique
        name and update the .rels targets in-place so the copy is fully independent."""
        import defusedxml.minidom as _dm

        rels_dom = _dm.parse(str(tmp_rels_path))
        changed  = False

        for rel in rels_dom.getElementsByTagName("Relationship"):
            rel_type = rel.getAttribute("Type")
            if rel_type not in DIAGRAM_REL_TYPES:
                continue

            target = rel.getAttribute("Target")   # e.g. "../diagrams/data1.xml"
            if not target:
                continue

            # Resolve the absolute path of the original diagram file.
            # Targets in .rels are relative to the SLIDE XML part, e.g.:
            #   slide at ppt/slides/slide4.xml
            #   target "../diagrams/data1.xml" → ppt/diagrams/data1.xml
            # tmp_rels is at ppt/slides/_rels/_reorder_tmp_N.xml.rels,
            # so slide_dir = tmp_rels_path.parent.parent = ppt/slides/
            slide_dir  = tmp_rels_path.parent.parent   # ppt/slides/
            diag_path  = (slide_dir / target).resolve()

            if not diag_path.exists():
                continue

            # Find a unique destination name by incrementing the trailing number.
            stem   = diag_path.stem    # e.g. "data1"
            suffix = diag_path.suffix  # e.g. ".xml"
            parent = diag_path.parent

            # Strip trailing digits and increment.
            m = re.match(r"^(.*?)(\d+)$", stem)
            if m:
                base_name, num = m.group(1), int(m.group(2))
            else:
                base_name, num = stem, 1

            while True:
                num += 1
                new_name = f"{base_name}{num}{suffix}"
                new_path = parent / new_name
                if not new_path.exists():
                    break

            shutil.copy2(str(diag_path), str(new_path))

            # Rewrite target to point at new filename (keep the ../ prefix).
            new_target = target.rsplit("/", 1)[0] + "/" + new_name
            rel.setAttribute("Target", new_target)
            changed = True

        if changed:
            with open(tmp_rels_path, "wb") as _f:
                _f.write(_serialize_xml_bytes(rels_dom))

    for new_pos, src_idx in enumerate(valid_order):
        src_name = idx_to_source.get(src_idx)
        if src_name is None:
            print(f"  [WARN] No source file for original slide {src_idx+1} — skipping position {new_pos+1}")
            continue

        src_slide = slides_dir / src_name
        src_rels  = slides_rels_dir / f"{src_name}.rels"

        if not src_slide.exists():
            print(f"  [ERROR] Source slide not found: {src_slide} — skipping")
            continue

        tmp_slide = slides_dir      / f"_reorder_tmp_{new_pos + 1}.xml"
        tmp_rels  = slides_rels_dir / f"_reorder_tmp_{new_pos + 1}.xml.rels"

        shutil.copy2(str(src_slide), str(tmp_slide))
        if src_rels.exists():
            shutil.copy2(str(src_rels), str(tmp_rels))
        else:
            # Create a minimal valid empty rels so pack_pptx never trips.
            tmp_rels.write_bytes(
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>\n'
            )

        # If this source was already used before, deep-copy its diagram files
        # so this copy becomes fully independent (shared diagram files corrupt PPTX).
        if src_idx in seen_sources:
            _deep_copy_slide_resources(tmp_rels, unpacked_path)

        seen_sources[src_idx] = seen_sources.get(src_idx, 0) + 1
        tmp_pairs.append((tmp_slide, tmp_rels))

    # ── Phase 4: Delete all original slide files ──────────────────────────
    # Only files originally referenced in presentation.xml.rels are removed.
    for fname in rid_to_file.values():
        for p in [slides_dir / fname, slides_rels_dir / f"{fname}.rels"]:
            if p.exists():
                p.unlink()

    # ── Phase 5: Rename temp files → slide1.xml, slide2.xml … ─────────────
    final_names: list[str] = []
    for new_pos, (tmp_slide, tmp_rels) in enumerate(tmp_pairs):
        final_name = f"slide{new_pos + 1}.xml"
        final_slide = slides_dir      / final_name
        final_rels  = slides_rels_dir / f"{final_name}.rels"
        tmp_slide.rename(final_slide)
        tmp_rels.rename(final_rels)
        final_names.append(final_name)

    n_new = len(final_names)

    # ── Phase 6: Rebuild presentation.xml.rels (minidom) ──────────────────
    max_existing_rid = 0
    for rel in non_slide_rels:
        m = re.match(r"rId(\d+)$", rel.getAttribute("Id"))
        if m:
            max_existing_rid = max(max_existing_rid, int(m.group(1)))

    new_rels_dom = defusedxml.minidom.parseString(
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )
    rels_node = new_rels_dom.documentElement

    for rel in non_slide_rels:
        nr = new_rels_dom.createElement("Relationship")
        nr.setAttribute("Id",     rel.getAttribute("Id"))
        nr.setAttribute("Type",   rel.getAttribute("Type"))
        nr.setAttribute("Target", rel.getAttribute("Target"))
        rels_node.appendChild(nr)

    new_slide_rids: list[str] = []
    for i, fname in enumerate(final_names):
        rid = f"rId{max_existing_rid + 1 + i}"
        new_slide_rids.append(rid)
        nr = new_rels_dom.createElement("Relationship")
        nr.setAttribute("Id",     rid)
        nr.setAttribute("Type",   SLIDE_REL_TYPE)
        nr.setAttribute("Target", f"slides/{fname}")
        rels_node.appendChild(nr)

    with open(pres_rels_path, "wb") as f:
        f.write(_serialize_xml_bytes(new_rels_dom))

    # ── Phase 7: Rebuild presentation.xml sldIdLst (lxml) ─────────────────
    # Only scan existing <p:sldId> elements — NOT sldMasterId (which uses 2147483648+,
    # outside the valid slide ID range of 256-2147483647).
    existing_sld_ids = [
        int(el.get("id"))
        for el in sldIdLst.findall(f"{{{NS_P}}}sldId")
        if el.get("id") and str(el.get("id")).isdigit()
    ]
    max_sld_id = max(existing_sld_ids) if existing_sld_ids else 255

    for el in list(sldIdLst):
        sldIdLst.remove(el)

    for i, rid in enumerate(new_slide_rids):
        new_el = etree.SubElement(sldIdLst, f"{{{NS_P}}}sldId")
        new_el.set("id", str(max_sld_id + 1 + i))
        new_el.set(f"{{{NS_R}}}id", rid)

    tree.write(str(pres_xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── Phase 8: Update [Content_Types].xml (minidom) ─────────────────────
    SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    NOTES_CT = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"

    ct_dom   = defusedxml.minidom.parse(str(ct_path))
    ct_types = ct_dom.documentElement

    for override in list(ct_dom.getElementsByTagName("Override")):
        part = override.getAttribute("PartName")
        # Remove old slide overrides and notesSlide overrides (notes are rebuilt below)
        if re.match(r"^/ppt/slides/slide\d+\.xml$", part) or \
           re.match(r"^/ppt/notesSlides/notesSlide\d+\.xml$", part):
            if override.parentNode:
                override.parentNode.removeChild(override)

    for fname in final_names:
        ov = ct_dom.createElement("Override")
        ov.setAttribute("PartName",    f"/ppt/slides/{fname}")
        ov.setAttribute("ContentType", SLIDE_CT)
        ct_types.appendChild(ov)

    with open(ct_path, "wb") as f:
        f.write(_serialize_xml_bytes(ct_dom))

    # ── Phase 9: Remove orphaned notesSlides (their slide refs become stale) ──
    # After renaming slides, every notesSlide .rels pointing to an old slide
    # filename becomes a broken reference. The simplest correct fix is to
    # drop all notesSlide files and strip notesSlide refs from slide .rels.
    # Notes are template content that the injection pipeline does not use.
    NOTES_REL_TYPE = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    )
    notes_dir      = unpacked_path / "ppt" / "notesSlides"
    notes_rels_dir = notes_dir / "_rels"

    # Delete all notesSlide XML files and their .rels
    if notes_dir.exists():
        for fp in notes_dir.glob("notesSlide*.xml"):
            fp.unlink()
        if notes_rels_dir.exists():
            for fp in notes_rels_dir.glob("notesSlide*.xml.rels"):
                fp.unlink()

    # Strip notesSlide Relationship entries from each new slide .rels
    for final_name in final_names:
        rels_path = slides_rels_dir / f"{final_name}.rels"
        if not rels_path.exists():
            continue
        slide_rels_dom = defusedxml.minidom.parse(str(rels_path))
        changed = False
        for rel in list(slide_rels_dom.getElementsByTagName("Relationship")):
            if rel.getAttribute("Type") == NOTES_REL_TYPE:
                if rel.parentNode:
                    rel.parentNode.removeChild(rel)
                    changed = True
        if changed:
            with open(rels_path, "wb") as f:
                f.write(_serialize_xml_bytes(slide_rels_dom))

    print(f"  [OK] Slides reordered. New count: {n_new}")
