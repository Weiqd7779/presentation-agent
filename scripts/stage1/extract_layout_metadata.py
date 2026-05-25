import json
import re
import sys
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont

SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import NS_A, NS_P, configure_utf8_output


EMU_PER_INCH = 914400
DEFAULT_SLIDE_W = 12192000
DEFAULT_SLIDE_H = 6858000


def slide_sort_key(path: Path) -> int:
    match = re.search(r"slide(\d+)\.xml$", path.name)
    return int(match.group(1)) if match else 999999


def text_of_shape(el) -> str:
    return " ".join(
        node.text.strip()
        for node in el.iter(f"{{{NS_A}}}t")
        if node.text and node.text.strip()
    )


def raw_shape_geometry(el) -> dict | None:
    xfrm = el.find(f".//{{{NS_A}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if off is None or ext is None:
        return None
    try:
        x = int(off.get("x", "0"))
        y = int(off.get("y", "0"))
        w = int(ext.get("cx", "0"))
        h = int(ext.get("cy", "0"))
    except ValueError:
        return None
    if w <= 0 or h <= 0:
        return None
    return {"x": x, "y": y, "w": w, "h": h}


def raw_graphic_frame_geometry(el) -> dict | None:
    xfrm = el.find(f"./{{{NS_P}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    if off is None or ext is None:
        return None
    try:
        x = int(off.get("x", "0"))
        y = int(off.get("y", "0"))
        w = int(ext.get("cx", "0"))
        h = int(ext.get("cy", "0"))
    except ValueError:
        return None

    table = el.find(f".//{{{NS_A}}}tbl")
    if table is not None:
        grid_w = sum(
            int(col.get("w", "0"))
            for col in table.findall(f".//{{{NS_A}}}tblGrid/{{{NS_A}}}gridCol")
        )
        rows_h = sum(
            int(row.get("h", "0"))
            for row in table.findall(f"./{{{NS_A}}}tr")
        )
        w = max(w, grid_w)
        h = max(h, rows_h)

    if w <= 0 or h <= 0:
        return None
    return {"x": x, "y": y, "w": w, "h": h}


def transform_geometry(geom: dict, transform: dict) -> dict:
    x_scale = transform["w"] / transform["ch_w"] if transform["ch_w"] else 1
    y_scale = transform["h"] / transform["ch_h"] if transform["ch_h"] else 1
    return {
        "x": int(transform["x"] + (geom["x"] - transform["ch_x"]) * x_scale),
        "y": int(transform["y"] + (geom["y"] - transform["ch_y"]) * y_scale),
        "w": int(geom["w"] * x_scale),
        "h": int(geom["h"] * y_scale),
    }


def combine_group_transform(parent: dict, group_el) -> dict:
    xfrm = group_el.find(f"./{{{NS_P}}}grpSpPr/{{{NS_A}}}xfrm")
    if xfrm is None:
        return parent
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    ch_off = xfrm.find(f"{{{NS_A}}}chOff")
    ch_ext = xfrm.find(f"{{{NS_A}}}chExt")
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return parent

    local = {
        "x": int(off.get("x", "0")),
        "y": int(off.get("y", "0")),
        "w": int(ext.get("cx", "1")),
        "h": int(ext.get("cy", "1")),
    }
    absolute = transform_geometry(local, parent)
    return {
        "x": absolute["x"],
        "y": absolute["y"],
        "w": absolute["w"],
        "h": absolute["h"],
        "ch_x": int(ch_off.get("x", "0")),
        "ch_y": int(ch_off.get("y", "0")),
        "ch_w": int(ch_ext.get("cx", "1")),
        "ch_h": int(ch_ext.get("cy", "1")),
    }


def visible_ratio(geom: dict, slide_w: int, slide_h: int) -> float:
    x0 = max(0, geom["x"])
    y0 = max(0, geom["y"])
    x1 = min(slide_w, geom["x"] + geom["w"])
    y1 = min(slide_h, geom["y"] + geom["h"])
    visible_w = max(0, x1 - x0)
    visible_h = max(0, y1 - y0)
    return (visible_w * visible_h) / max(1, geom["w"] * geom["h"])


def is_meaningful_container(geom: dict, slide_w: int, slide_h: int) -> bool:
    area_ratio = (geom["w"] * geom["h"]) / max(1, slide_w * slide_h)
    width_ratio = geom["w"] / max(1, slide_w)
    height_ratio = geom["h"] / max(1, slide_h)
    return (
        area_ratio >= 0.08
        and width_ratio >= 0.18
        and height_ratio >= 0.18
        and visible_ratio(geom, slide_w, slide_h) >= 0.85
    )


def region_for_box(box: dict, slide_w: int, slide_h: int) -> str:
    cx = (box["x"] + box["w"] / 2) / slide_w
    cy = (box["y"] + box["h"] / 2) / slide_h
    horizontal = "left" if cx < 0.33 else "right" if cx > 0.67 else "center"
    vertical = "top" if cy < 0.33 else "bottom" if cy > 0.67 else "middle"
    return f"{vertical}-{horizontal}"


def likely_role(name: str, shape_type: str, text: str, geom: dict, slide_w: int, slide_h: int) -> str:
    lname = name.lower()
    ltext = text.lower()
    area_ratio = (geom["w"] * geom["h"]) / max(1, slide_w * slide_h)
    y_ratio = geom["y"] / max(1, slide_h)

    if shape_type == "picture":
        return "image"
    if shape_type == "table":
        return "table"
    if shape_type == "shape":
        return "container" if area_ratio > 0.03 else "decorative_shape"
    if "title" in lname or "title" in ltext or y_ratio < 0.18:
        return "title"
    if "subtitle" in lname:
        return "subtitle"
    if "metric" in lname or re.search(r"\d+%|\d+x|\d+\s*/\s*\d+", text):
        return "metric"
    if "header" in lname:
        return "table_header"
    if "row" in lname:
        return "table_cell"
    if "body" in lname or area_ratio > 0.04:
        return "body"
    if "contact" in lname or "@" in text:
        return "contact"
    return "label"


def iter_slide_shapes(parent, transform: dict, slide_w: int, slide_h: int):
    for child in parent:
        if child.tag == f"{{{NS_P}}}grpSp":
            group_transform = combine_group_transform(transform, child)
            yield from iter_slide_shapes(child, group_transform, slide_w, slide_h)
            continue
        if child.tag == f"{{{NS_P}}}graphicFrame":
            c_nv_pr = child.find(f".//{{{NS_P}}}cNvPr")
            name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""
            raw_geom = raw_graphic_frame_geometry(child)
            text = text_of_shape(child)
            if name and raw_geom is not None and text:
                yield name, "table", text, transform_geometry(raw_geom, transform)
            continue
        if child.tag not in {f"{{{NS_P}}}sp", f"{{{NS_P}}}pic"}:
            continue
        c_nv_pr = child.find(f".//{{{NS_P}}}cNvPr")
        name = c_nv_pr.get("name", "") if c_nv_pr is not None else ""
        raw_geom = raw_shape_geometry(child)
        if not name or raw_geom is None:
            continue
        geom = transform_geometry(raw_geom, transform)
        if geom["w"] <= 0 or geom["h"] <= 0:
            continue
        if child.tag == f"{{{NS_P}}}pic":
            yield name, "picture", "", geom
            continue
        text = text_of_shape(child)
        if text:
            yield name, "text", text, geom


def read_slide_size(unpacked_dir: Path) -> tuple[int, int]:
    presentation_xml = unpacked_dir / "ppt" / "presentation.xml"
    if not presentation_xml.exists():
        return DEFAULT_SLIDE_W, DEFAULT_SLIDE_H
    root = etree.parse(str(presentation_xml)).getroot()
    sld_sz = root.find(f".//{{{NS_P}}}sldSz")
    if sld_sz is None:
        return DEFAULT_SLIDE_W, DEFAULT_SLIDE_H
    return int(sld_sz.get("cx", DEFAULT_SLIDE_W)), int(sld_sz.get("cy", DEFAULT_SLIDE_H))


def find_slide_images(output_dir: Path, prefix: str) -> dict[int, Path]:
    images = {}
    for path in output_dir.glob(f"{prefix}-*.jpg"):
        match = re.search(r"-(\d+)\.jpg$", path.name)
        if match:
            images[int(match.group(1))] = path
    return images


def load_font(size: int = 18):
    for candidate in ("arial.ttf", "calibri.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(candidate, size)
        except OSError:
            continue
    return ImageFont.load_default()


def annotate_image(image_path: Path, shapes: list[dict], slide_w: int, slide_h: int, output_path: Path) -> None:
    image = Image.open(image_path).convert("RGB")
    draw = ImageDraw.Draw(image)
    font = load_font(18)
    scale_x = image.width / slide_w
    scale_y = image.height / slide_h

    for index, shape in enumerate(shapes, 1):
        geom = shape["geometry"]
        x0 = int(geom["x"] * scale_x)
        y0 = int(geom["y"] * scale_y)
        x1 = int((geom["x"] + geom["w"]) * scale_x)
        y1 = int((geom["y"] + geom["h"]) * scale_y)
        color = (48, 105, 85) if shape["type"] == "text" else (166, 93, 48)
        draw.rectangle((x0, y0, x1, y1), outline=color, width=3)
        label = f"{index}. {shape['name']}"
        bbox = draw.textbbox((x0, y0), label, font=font)
        pad = 4
        draw.rectangle((bbox[0] - pad, bbox[1] - pad, bbox[2] + pad, bbox[3] + pad), fill=color)
        draw.text((x0, y0), label, fill=(255, 255, 245), font=font)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path, quality=92)


def extract_layout(unpacked_dir: Path, output_dir: Path, image_prefix: str) -> dict:
    slide_w, slide_h = read_slide_size(unpacked_dir)
    slide_images = find_slide_images(output_dir, image_prefix)
    slides = []
    annotated_dir = output_dir / "annotated_slides"

    for slide_path in sorted((unpacked_dir / "ppt" / "slides").glob("slide*.xml"), key=slide_sort_key):
        slide_num = slide_sort_key(slide_path)
        root = etree.parse(str(slide_path)).getroot()
        shapes = []

        sp_tree = root.find(f".//{{{NS_P}}}spTree")
        identity_transform = {
            "x": 0,
            "y": 0,
            "w": slide_w,
            "h": slide_h,
            "ch_x": 0,
            "ch_y": 0,
            "ch_w": slide_w,
            "ch_h": slide_h,
        }
        shape_iter = iter_slide_shapes(sp_tree, identity_transform, slide_w, slide_h) if sp_tree is not None else []
        for name, shape_type, text, geom in shape_iter:
            shapes.append(
                {
                    "name": name,
                    "type": shape_type,
                    "text": text,
                    "geometry": geom,
                    "geometry_in": {
                        "x": round(geom["x"] / EMU_PER_INCH, 3),
                        "y": round(geom["y"] / EMU_PER_INCH, 3),
                        "w": round(geom["w"] / EMU_PER_INCH, 3),
                        "h": round(geom["h"] / EMU_PER_INCH, 3),
                    },
                    "region": region_for_box(geom, slide_w, slide_h),
                    "likely_role": likely_role(name, shape_type, text, geom, slide_w, slide_h),
                }
            )

        annotated_path = None
        if slide_num in slide_images:
            annotated_path = annotated_dir / f"slide-{slide_num:03d}_annotated.jpg"
            annotate_image(slide_images[slide_num], shapes, slide_w, slide_h, annotated_path)

        slides.append(
            {
                "slide": slide_num,
                "image": str(slide_images.get(slide_num, "")),
                "annotated_image": str(annotated_path or ""),
                "shapes": shapes,
            }
        )

    return {
        "slide_size_emu": {"w": slide_w, "h": slide_h},
        "slide_size_in": {"w": round(slide_w / EMU_PER_INCH, 3), "h": round(slide_h / EMU_PER_INCH, 3)},
        "slides": slides,
    }


def main() -> None:
    configure_utf8_output()
    if len(sys.argv) != 4:
        print("Usage: python extract_layout_metadata.py <unpacked_dir> <output_dir> <slide_image_prefix>")
        sys.exit(1)

    unpacked_dir = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    image_prefix = sys.argv[3]
    metadata = extract_layout(unpacked_dir, output_dir, image_prefix)
    out_path = output_dir / "placeholder_layout.json"
    out_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[done] Layout metadata: {out_path}")
    print(f"[done] Annotated slides: {output_dir / 'annotated_slides'}")


if __name__ == "__main__":
    main()
