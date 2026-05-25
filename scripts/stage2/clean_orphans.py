"""Remove orphaned files from an unpacked PPTX directory.

All structural XML manipulation uses defusedxml.minidom to preserve
namespace prefixes.  lxml is NOT used — it rewrites ns0:/ns1: prefixes
which PowerPoint interprets as corruption.

Removes:
- Orphaned slides (not in <p:sldIdLst>) and their .rels
- Orphaned .rels files for deleted resources
- Unreferenced media, embeddings, charts, diagrams, drawings, ink files
- Unreferenced theme files
- Unreferenced notes slides
- Content-Type overrides for deleted files

Usage:
    python clean_orphans.py <unpacked_dir>
"""

import re
import sys
from pathlib import Path

# Provide access to pptx_utils
SKILL_DIR = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(SKILL_DIR))

try:
    from pptx_utils import configure_utf8_output
except ImportError:
    def configure_utf8_output(): pass

import defusedxml.minidom


def _serialize_xml_bytes(dom) -> bytes:
    """Serialize minidom DOM to bytes, ensuring standalone='yes' in XML declaration."""
    raw = dom.toxml(encoding="UTF-8")
    return raw.replace(
        b'<?xml version="1.0" encoding="UTF-8"?>',
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        1,
    )


# ─────────────────────────────────────────────────────────────────────────
# 1. Discover live slides from sldIdLst
# ─────────────────────────────────────────────────────────────────────────

def _get_live_slide_files(unpacked_dir: Path) -> set[str]:
    """Return set of slide filenames that are referenced in <p:sldIdLst>."""
    pres_path = unpacked_dir / "ppt" / "presentation.xml"
    pres_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"

    if not pres_path.exists() or not pres_rels_path.exists():
        return set()

    # Map rId → slide filename from presentation.xml.rels
    rels_dom = defusedxml.minidom.parse(str(pres_rels_path))
    rid_to_slide: dict[str, str] = {}
    for rel in rels_dom.getElementsByTagName("Relationship"):
        rid = rel.getAttribute("Id")
        target = rel.getAttribute("Target")
        rel_type = rel.getAttribute("Type")
        if "slide" in rel_type and target.startswith("slides/"):
            rid_to_slide[rid] = target.replace("slides/", "")

    # Read referenced rIds from presentation.xml → sldIdLst
    pres_content = pres_path.read_text(encoding="utf-8")
    referenced_rids = set(re.findall(r'<p:sldId[^>]*r:id="([^"]+)"', pres_content))

    return {rid_to_slide[rid] for rid in referenced_rids if rid in rid_to_slide}


# ─────────────────────────────────────────────────────────────────────────
# 2. Remove orphaned slides
# ─────────────────────────────────────────────────────────────────────────

def _remove_orphaned_slides(unpacked_dir: Path) -> list[str]:
    """Delete slide XMLs and their .rels if not in <p:sldIdLst>.
    Also prune the matching <Relationship> from presentation.xml.rels.
    """
    slides_dir = unpacked_dir / "ppt" / "slides"
    slides_rels_dir = slides_dir / "_rels"
    pres_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"

    if not slides_dir.exists():
        return []

    live = _get_live_slide_files(unpacked_dir)
    removed: list[str] = []

    for slide_file in slides_dir.glob("slide*.xml"):
        if slide_file.name not in live:
            rel = slide_file.relative_to(unpacked_dir)
            slide_file.unlink()
            removed.append(str(rel))

            rels_file = slides_rels_dir / f"{slide_file.name}.rels"
            if rels_file.exists():
                rels_file.unlink()
                removed.append(str(rels_file.relative_to(unpacked_dir)))

    # Prune rels entries for removed slides
    if removed and pres_rels_path.exists():
        rels_dom = defusedxml.minidom.parse(str(pres_rels_path))
        changed = False
        for rel in list(rels_dom.getElementsByTagName("Relationship")):
            target = rel.getAttribute("Target")
            if target.startswith("slides/"):
                slide_name = target.replace("slides/", "")
                if slide_name not in live:
                    if rel.parentNode:
                        rel.parentNode.removeChild(rel)
                        changed = True
        if changed:
            with open(pres_rels_path, "wb") as f:
                f.write(_serialize_xml_bytes(rels_dom))

    return removed


# ─────────────────────────────────────────────────────────────────────────
# 3. Collect all file references from .rels
# ─────────────────────────────────────────────────────────────────────────

def _get_all_referenced_files(unpacked_dir: Path) -> set[Path]:
    """Walk every .rels and collect the resolved Paths of all targets."""
    referenced: set[Path] = set()
    for rels_file in unpacked_dir.rglob("*.rels"):
        try:
            dom = defusedxml.minidom.parse(str(rels_file))
            for rel in dom.getElementsByTagName("Relationship"):
                target = rel.getAttribute("Target")
                if not target or target.startswith("http"):
                    continue
                target_path = (rels_file.parent.parent / target).resolve()
                try:
                    referenced.add(target_path.relative_to(unpacked_dir.resolve()))
                except ValueError:
                    pass
        except Exception:
            pass
    return referenced


def _get_slide_referenced_files(unpacked_dir: Path) -> set[Path]:
    """Collect files referenced by slide and notesSlide relationship graphs."""
    referenced: set[Path] = set()
    rel_dirs = [
        unpacked_dir / "ppt" / "slides" / "_rels",
        unpacked_dir / "ppt" / "notesSlides" / "_rels",
    ]
    for rel_dir in rel_dirs:
        if not rel_dir.exists():
            continue
        for rels_file in rel_dir.glob("*.rels"):
            try:
                dom = defusedxml.minidom.parse(str(rels_file))
                for rel in dom.getElementsByTagName("Relationship"):
                    target = rel.getAttribute("Target")
                    if not target:
                        continue
                    target_path = (rels_file.parent.parent / target).resolve()
                    try:
                        referenced.add(target_path.relative_to(unpacked_dir.resolve()))
                    except ValueError:
                        pass
            except Exception:
                pass
    return referenced


# ─────────────────────────────────────────────────────────────────────────
# 4. Remove orphaned resource .rels
# ─────────────────────────────────────────────────────────────────────────

def _remove_orphaned_rels(unpacked_dir: Path) -> list[str]:
    """Remove .rels files whose parent resource is either missing or unreferenced."""
    resource_dirs = ["charts", "diagrams", "drawings"]
    removed: list[str] = []
    graph_refs = _get_all_referenced_files(unpacked_dir)
    slide_refs = _get_slide_referenced_files(unpacked_dir)
    primary_refs = graph_refs | slide_refs

    for dir_name in resource_dirs:
        rels_dir = unpacked_dir / "ppt" / dir_name / "_rels"
        if not rels_dir.exists():
            continue
        for rels_file in rels_dir.glob("*.rels"):
            resource_file = rels_dir.parent / rels_file.name.replace(".rels", "")
            try:
                resource_rel = resource_file.resolve().relative_to(unpacked_dir.resolve())
            except ValueError:
                continue
            if not resource_file.exists() or resource_rel not in primary_refs:
                rels_file.unlink()
                removed.append(str(rels_file.relative_to(unpacked_dir)))
    return removed


# ─────────────────────────────────────────────────────────────────────────
# 5. Remove orphaned media / embeddings / etc.
# ─────────────────────────────────────────────────────────────────────────

def _remove_orphaned_files(unpacked_dir: Path, referenced: set[Path]) -> list[str]:
    """Delete unreferenced files from resource directories."""
    resource_dirs = ["media", "embeddings", "charts", "diagrams", "tags", "drawings", "ink"]
    removed: list[str] = []

    for dir_name in resource_dirs:
        dir_path = unpacked_dir / "ppt" / dir_name
        if not dir_path.exists():
            continue
        for file_path in dir_path.glob("*"):
            if not file_path.is_file():
                continue
            rel = file_path.relative_to(unpacked_dir)
            if rel not in referenced:
                file_path.unlink()
                removed.append(str(rel))

    # Unreferenced themes
    theme_dir = unpacked_dir / "ppt" / "theme"
    if theme_dir.exists():
        for fp in theme_dir.glob("theme*.xml"):
            rel = fp.relative_to(unpacked_dir)
            if rel not in referenced:
                fp.unlink()
                removed.append(str(rel))
                theme_rels = theme_dir / "_rels" / f"{fp.name}.rels"
                if theme_rels.exists():
                    theme_rels.unlink()
                    removed.append(str(theme_rels.relative_to(unpacked_dir)))

    # Unreferenced notes slides
    notes_dir = unpacked_dir / "ppt" / "notesSlides"
    if notes_dir.exists():
        for fp in notes_dir.glob("*.xml"):
            if not fp.is_file():
                continue
            rel = fp.relative_to(unpacked_dir)
            if rel not in referenced:
                fp.unlink()
                removed.append(str(rel))

        notes_rels_dir = notes_dir / "_rels"
        if notes_rels_dir.exists():
            for fp in notes_rels_dir.glob("*.rels"):
                notes_file = notes_dir / fp.name.replace(".rels", "")
                if not notes_file.exists():
                    fp.unlink()
                    removed.append(str(fp.relative_to(unpacked_dir)))

    return removed


# ─────────────────────────────────────────────────────────────────────────
# 6. Update [Content_Types].xml
# ─────────────────────────────────────────────────────────────────────────

def _update_content_types(unpacked_dir: Path, removed_files: list[str]) -> list[str]:
    """Remove stale <Override> entries by checking actual on-disk existence."""
    ct_path = unpacked_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return []

    dom = defusedxml.minidom.parse(str(ct_path))
    changed = False
    removed_overrides: list[str] = []

    removed_set = {str(Path(p).as_posix()).lstrip("/") for p in removed_files}

    for override in list(dom.getElementsByTagName("Override")):
        part_name = override.getAttribute("PartName").lstrip("/")
        should_remove = part_name in removed_set or not (unpacked_dir / part_name).exists()
        if should_remove:
            if override.parentNode:
                override.parentNode.removeChild(override)
                changed = True
                removed_overrides.append(part_name)

    if changed:
        with open(ct_path, "wb") as f:
            f.write(_serialize_xml_bytes(dom))

    return removed_overrides


def _validate_integrity(unpacked_dir: Path) -> list[str]:
    """Fail-fast validation after each cleanup phase."""
    errors: list[str] = []

    for rels_file in unpacked_dir.rglob("*.rels"):
        try:
            dom = defusedxml.minidom.parse(str(rels_file))
        except Exception as exc:
            errors.append(f"Malformed rels {rels_file.relative_to(unpacked_dir)}: {exc}")
            continue
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if not target or target.startswith("http"):
                continue
            target_path = (rels_file.parent.parent / target).resolve()
            if not target_path.exists():
                errors.append(f"Broken rel target {rels_file.relative_to(unpacked_dir)} -> {target}")

    ct_path = unpacked_dir / "[Content_Types].xml"
    if ct_path.exists():
        try:
            dom = defusedxml.minidom.parse(str(ct_path))
            for override in dom.getElementsByTagName("Override"):
                part_name = override.getAttribute("PartName").lstrip("/")
                if part_name and not (unpacked_dir / part_name).exists():
                    errors.append(f"Orphan Content_Types override: {part_name}")
        except Exception as exc:
            errors.append(f"Malformed [Content_Types].xml: {exc}")

    return errors


# ─────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────

def clean_orphans(unpacked_dir: Path) -> list[str]:
    """Run the full orphan-cleanup pipeline. Returns list of removed paths."""
    all_removed: list[str] = []

    # Phase 1: slide-level
    slide_removed = _remove_orphaned_slides(unpacked_dir)
    all_removed.extend(slide_removed)
    ct_removed = _update_content_types(unpacked_dir, all_removed)
    all_removed.extend(ct_removed)
    errs = _validate_integrity(unpacked_dir)
    if errs:
        raise RuntimeError("Integrity check failed after slide cleanup:\n" + "\n".join(errs))

    # Phase 2: iterative resource cleanup (may cascade)
    while True:
        rels_removed = _remove_orphaned_rels(unpacked_dir)
        referenced = _get_all_referenced_files(unpacked_dir)
        files_removed = _remove_orphaned_files(unpacked_dir, referenced)
        ct_removed = _update_content_types(unpacked_dir, rels_removed + files_removed)

        batch = rels_removed + files_removed + ct_removed
        if not batch:
            break
        all_removed.extend(batch)

        errs = _validate_integrity(unpacked_dir)
        if errs:
            raise RuntimeError("Integrity check failed during iterative cleanup:\n" + "\n".join(errs))

    # Phase 3: sync Content_Types
    if all_removed:
        ct_removed = _update_content_types(unpacked_dir, all_removed)
        all_removed.extend(ct_removed)

    errs = _validate_integrity(unpacked_dir)
    if errs:
        raise RuntimeError("Integrity check failed after orphan cleanup:\n" + "\n".join(errs))

    return all_removed


if __name__ == "__main__":
    configure_utf8_output()
    if len(sys.argv) != 2:
        print("Usage: python clean_orphans.py <unpacked_dir>", file=sys.stderr)
        sys.exit(1)

    unpacked_dir = Path(sys.argv[1])
    if not unpacked_dir.exists():
        print(f"[ERROR] {unpacked_dir} not found", file=sys.stderr)
        sys.exit(1)

    removed = clean_orphans(unpacked_dir)
    if removed:
        print(f"Removed {len(removed)} orphaned files:")
        for f in removed:
            print(f"  {f}")
    else:
        print("No orphaned files found.")
