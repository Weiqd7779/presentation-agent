"""
Apply a data-driven presentation content mapping.

Expected mapping JSON:
{
  "slide_order": null,
  "content_map": {
    "1": {
      "TextBox 7": "Title",
      "TextBox 8": ["First paragraph", "Second paragraph"]
    }
  }
}
"""

import json
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import (
    configure_utf8_output,
    load_slide,
    reorder_slides,
    save_slide,
    set_textbox_text,
    replace_picture_image,
    set_table_data,
)


def _load_mapping(mapping_path: Path) -> tuple[list[int] | None, dict[int, dict[str, object]]]:
    data = json.loads(mapping_path.read_text(encoding="utf-8"))
    slide_order = data.get("slide_order")
    content_map = data.get("content_map")
    image_map = data.get("image_map", {})
    table_map = data.get("table_map", {})

    if slide_order is not None:
        if not isinstance(slide_order, list) or not all(isinstance(v, int) for v in slide_order):
            raise ValueError("'slide_order' must be null or a list of integer slide numbers")

    if not isinstance(content_map, dict):
        raise ValueError("'content_map' must be an object keyed by final slide number")

    normalized: dict[int, dict[str, object]] = {}
    for slide_key, boxes in content_map.items():
        try:
            slide_num = int(slide_key)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"content_map key must be a slide number, got {slide_key!r}") from exc
        if slide_num < 1:
            raise ValueError(f"slide number must be >= 1, got {slide_num}")
        if not isinstance(boxes, dict):
            raise ValueError(f"content_map[{slide_key!r}] must be an object of shape-name to text")
        normalized[slide_num] = boxes

    if not isinstance(image_map, dict):
        raise ValueError("'image_map' must be an object keyed by final slide number")

    normalized_images: dict[int, dict[str, object]] = {}
    for slide_key, pictures in image_map.items():
        try:
            slide_num = int(slide_key)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"image_map key must be a slide number, got {slide_key!r}") from exc
        if not isinstance(pictures, dict):
            raise ValueError(f"image_map[{slide_key!r}] must be an object of picture-name to image options")
        normalized_images[slide_num] = pictures

    if not isinstance(table_map, dict):
        raise ValueError("'table_map' must be an object keyed by final slide number")

    normalized_tables: dict[int, dict[str, object]] = {}
    for slide_key, tables in table_map.items():
        try:
            slide_num = int(slide_key)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"table_map key must be a slide number, got {slide_key!r}") from exc
        if not isinstance(tables, dict):
            raise ValueError(f"table_map[{slide_key!r}] must be an object of table-name to options")
        normalized_tables[slide_num] = tables

    return slide_order, normalized, normalized_images, normalized_tables


def apply_mapping(unpacked_dir: Path, mapping_path: Path) -> None:
    configure_utf8_output()
    unpacked_dir = Path(unpacked_dir)
    mapping_path = Path(mapping_path)

    slide_order, content_map, image_map, table_map = _load_mapping(mapping_path)

    if slide_order:
        print(f"==> Reordering slides: {slide_order}")
        reorder_slides(unpacked_dir, slide_order)
    else:
        print("==> Keeping original slide order.")

    print(f"==> Applying mapping: {mapping_path}")
    for slide_num in sorted(content_map):
        print(f"  Processing Slide {slide_num}...")
        tree, root = load_slide(unpacked_dir, slide_num)
        for box_name, value in content_map[slide_num].items():
            options = {}
            text_value = value
            if isinstance(value, dict):
                text_value = value.get("text", "")
                options = {k: v for k, v in value.items() if k != "text"}

            if isinstance(text_value, list):
                text_items = [str(item) for item in text_value]
                first = text_items[0] if text_items else ""
                set_textbox_text(root, str(box_name), first, extra_paras=text_items[1:], **options)
            else:
                set_textbox_text(root, str(box_name), str(text_value), **options)
        save_slide(tree, unpacked_dir, slide_num)

    for slide_num in sorted(table_map):
        print(f"  Processing Slide {slide_num} tables...")
        tree, root = load_slide(unpacked_dir, slide_num)
        for table_name, value in table_map[slide_num].items():
            options = value if isinstance(value, dict) else {"rows": value}
            set_table_data(root, str(table_name), options)
        save_slide(tree, unpacked_dir, slide_num)

    for slide_num in sorted(image_map):
        print(f"  Processing Slide {slide_num} images...")
        for picture_name, value in image_map[slide_num].items():
            if isinstance(value, dict):
                image_path = value.get("path")
                fill = value.get("fill", "contain")
                background = value.get("background", "white")
            else:
                image_path = value
                fill = "contain"
                background = "white"
            replace_picture_image(
                unpacked_dir,
                slide_num,
                str(picture_name),
                image_path,
                fill=str(fill),
                background=str(background),
            )


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python apply_mapping_json.py <unpacked_dir> <mapping.json>")
        sys.exit(1)

    apply_mapping(Path(sys.argv[1]), Path(sys.argv[2]))
