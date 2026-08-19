import argparse
import json
import re
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import configure_utf8_output


SLIDE_RE = re.compile(r"^## Slide (\d+)\s*$")
TEXTBOX_RE = re.compile(r'^\s*-\s+Name:\s+"(.+)"\s*$')
PICTURE_RE = re.compile(r'^\s*-\s+Picture Name:\s+"(.+)"\s*$')
DECORATIVE_RE = re.compile(r"(footer|date|slide\s*number|page\s*number|copyright)", re.IGNORECASE)
LEFTOVER_RE = re.compile(r"(\{\{.+?\}\}|\[.+?\]|lorem ipsum|placeholder)", re.IGNORECASE)


def _load_json(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"mapping.json is not valid JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("mapping.json must contain a JSON object")
    return data


def _parse_placeholders(path: Path) -> tuple[dict[int, set[str]], dict[int, set[str]], dict[int, set[str]]]:
    if not path.exists():
        raise FileNotFoundError(f"Missing placeholders.txt: {path}")

    text_shapes: dict[int, set[str]] = {}
    picture_shapes: dict[int, set[str]] = {}
    decorative_shapes: dict[int, set[str]] = {}
    current_slide: int | None = None
    last_shape: str | None = None

    for line in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        slide_match = SLIDE_RE.match(line)
        if slide_match:
            current_slide = int(slide_match.group(1))
            text_shapes.setdefault(current_slide, set())
            picture_shapes.setdefault(current_slide, set())
            decorative_shapes.setdefault(current_slide, set())
            last_shape = None
            continue

        if current_slide is None:
            continue

        textbox_match = TEXTBOX_RE.match(line)
        if textbox_match:
            last_shape = textbox_match.group(1)
            text_shapes[current_slide].add(last_shape)
            if DECORATIVE_RE.search(last_shape):
                decorative_shapes[current_slide].add(last_shape)
            continue

        picture_match = PICTURE_RE.match(line)
        if picture_match:
            last_shape = picture_match.group(1)
            picture_shapes[current_slide].add(last_shape)
            if DECORATIVE_RE.search(last_shape):
                decorative_shapes[current_slide].add(last_shape)
            continue

        if last_shape and "Current Text:" in line and DECORATIVE_RE.search(line):
            decorative_shapes[current_slide].add(last_shape)

    return text_shapes, picture_shapes, decorative_shapes


def _validate_slide_order(data: dict, original_slide_count: int) -> tuple[list[int], list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    slide_order = data.get("slide_order")

    if slide_order is None:
        final_to_original = list(range(1, original_slide_count + 1))
    elif isinstance(slide_order, list) and all(isinstance(value, int) for value in slide_order):
        final_to_original = slide_order
        if not slide_order:
            errors.append("slide_order cannot be an empty list")
        for value in slide_order:
            if value < 1 or value > original_slide_count:
                errors.append(f"slide_order references original slide {value}, but valid range is 1..{original_slide_count}")
        duplicates = sorted({value for value in slide_order if slide_order.count(value) > 1})
        if duplicates:
            warnings.append(f"slide_order duplicates original slide(s): {', '.join(str(v) for v in duplicates)}")
    else:
        errors.append("slide_order must be null or a list of integer original slide numbers")
        final_to_original = []

    return final_to_original, errors, warnings


def _as_slide_map(data: dict, key: str, errors: list[str]) -> dict[int, dict[str, object]]:
    raw = data.get(key, {})
    if key == "content_map" and raw is None:
        errors.append("content_map is required and must be an object")
        return {}
    if not isinstance(raw, dict):
        errors.append(f"{key} must be an object keyed by final slide number")
        return {}

    result: dict[int, dict[str, object]] = {}
    for slide_key, shapes in raw.items():
        try:
            slide_num = int(slide_key)
        except (TypeError, ValueError):
            errors.append(f"{key} key {slide_key!r} is not a slide number")
            continue
        if slide_num < 1:
            errors.append(f"{key} key {slide_key!r} must be >= 1")
            continue
        if not isinstance(shapes, dict):
            errors.append(f"{key}[{slide_key!r}] must be an object of shape-name to value")
            continue
        result[slide_num] = shapes
    return result


def preflight_mapping(output_dir: Path, mapping_path: Path) -> dict:
    output_dir = Path(output_dir).resolve()
    mapping_path = Path(mapping_path).resolve()
    data = _load_json(mapping_path)

    text_shapes, picture_shapes, decorative_shapes = _parse_placeholders(output_dir / "placeholders.txt")
    original_slide_count = max(set(text_shapes) | set(picture_shapes), default=0)
    if original_slide_count == 0:
        raise ValueError("placeholders.txt did not contain any slide placeholders")

    final_to_original, errors, warnings = _validate_slide_order(data, original_slide_count)
    content_map = _as_slide_map(data, "content_map", errors)
    image_map = _as_slide_map(data, "image_map", errors)
    table_map = _as_slide_map(data, "table_map", errors)

    final_slide_count = len(final_to_original)
    touched_shapes = 0
    duplicate_final_slides: list[int] = []

    for map_name, slide_map, shape_source in (
        ("content_map", content_map, text_shapes),
        ("image_map", image_map, picture_shapes),
    ):
        for final_slide, shapes in slide_map.items():
            if final_slide > final_slide_count:
                errors.append(f"{map_name} references final slide {final_slide}, but final slide count is {final_slide_count}")
                continue
            original_slide = final_to_original[final_slide - 1]
            allowed = shape_source.get(original_slide, set())
            decorative = decorative_shapes.get(original_slide, set())
            for shape_name, value in shapes.items():
                touched_shapes += 1
                if shape_name not in allowed:
                    errors.append(
                        f"{map_name}[{final_slide}] shape {shape_name!r} is not present on original slide {original_slide}"
                    )
                if shape_name in decorative:
                    warnings.append(f"{map_name}[{final_slide}] writes decorative placeholder {shape_name!r}")
                if map_name == "content_map" and _value_contains_leftover(value):
                    warnings.append(f"{map_name}[{final_slide}] value for {shape_name!r} appears to contain placeholder text")

    for final_slide, shapes in table_map.items():
        if final_slide > final_slide_count:
            errors.append(f"table_map references final slide {final_slide}, but final slide count is {final_slide_count}")
            continue
        touched_shapes += len(shapes)
        warnings.append(
            f"table_map[{final_slide}] shape names are range-checked only; table extraction is not present in placeholders.txt"
        )

    if final_to_original:
        for idx, original in enumerate(final_to_original, start=1):
            if final_to_original.count(original) > 1 and idx not in content_map and idx not in image_map and idx not in table_map:
                duplicate_final_slides.append(idx)
        if duplicate_final_slides:
            warnings.append(
                "Duplicated final slide(s) have no custom mapping: "
                + ", ".join(str(slide) for slide in duplicate_final_slides)
            )

    report = {
        "pass": not errors,
        "errors": errors,
        "warnings": warnings,
        "metrics": {
            "original_slide_count": original_slide_count,
            "final_slide_count": final_slide_count,
            "mapped_slide_count": len(set(content_map) | set(image_map) | set(table_map)),
            "touched_shape_count": touched_shapes,
            "warning_count": len(warnings),
            "error_count": len(errors),
        },
    }
    return report


def _value_contains_leftover(value: object) -> bool:
    if isinstance(value, dict):
        return _value_contains_leftover(value.get("text", ""))
    if isinstance(value, list):
        return any(_value_contains_leftover(item) for item in value)
    return bool(LEFTOVER_RE.search(str(value)))


def main() -> None:
    configure_utf8_output()
    parser = argparse.ArgumentParser(description="Preflight a presentation-agent mapping.json")
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("mapping_json", type=Path)
    parser.add_argument("--report", type=Path, default=None)
    args = parser.parse_args()

    try:
        report = preflight_mapping(args.output_dir, args.mapping_json)
    except Exception as exc:
        report = {"pass": False, "errors": [str(exc)], "warnings": [], "metrics": {"error_count": 1, "warning_count": 0}}

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    for warning in report.get("warnings", []):
        print(f"[WARN] {warning}")
    for error in report.get("errors", []):
        print(f"[ERROR] {error}")

    metrics = report.get("metrics", {})
    print(
        "[preflight] "
        f"pass={report.get('pass')} "
        f"errors={metrics.get('error_count', 0)} "
        f"warnings={metrics.get('warning_count', 0)} "
        f"mapped_shapes={metrics.get('touched_shape_count', 0)}"
    )
    sys.exit(0 if report.get("pass") else 2)


if __name__ == "__main__":
    main()
