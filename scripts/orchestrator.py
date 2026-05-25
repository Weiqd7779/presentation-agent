import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

os.environ["MUPDF_MESSAGES"] = "0"

BASE_DIR = Path(__file__).resolve().parent
STAGE1_DIR = BASE_DIR / "stage1"
STAGE2_DIR = BASE_DIR / "stage2"

sys.path.append(str(STAGE1_DIR))
sys.path.append(str(STAGE2_DIR))
sys.path.append(str(BASE_DIR.parent))

try:
    from pptx_utils import configure_utf8_output
except ImportError:
    def configure_utf8_output() -> None:
        pass

configure_utf8_output()


def _is_known_noise(line: str) -> bool:
    upper = line.upper()
    return "MUPDF" in upper and "ERROR" in upper and "NO COMMON ANCESTOR" in upper


def run_command(cmd: list[str], stdout_file: Path | None = None, **kwargs) -> subprocess.Popen:
    """Run a subprocess with UTF-8 enabled and suppress known non-fatal MuPDF noise."""
    print(f"==> Running: {' '.join(str(c) for c in cmd)}")

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        **kwargs,
    )
    stdout, stderr = process.communicate()

    suppressed = 0
    if stdout:
        lines = stdout.splitlines()
        visible = []
        for line in lines:
            if _is_known_noise(line):
                suppressed += 1
            elif suppressed and not line.strip():
                continue
            else:
                visible.append(line)
        if stdout_file:
            stdout_file.write_text("\n".join(visible) + ("\n" if visible else ""), encoding="utf-8")
        elif visible:
            print("\n".join(visible))

    if stderr:
        visible_err = []
        for line in stderr.splitlines():
            if _is_known_noise(line):
                suppressed += 1
            elif suppressed and not line.strip():
                continue
            else:
                visible_err.append(line)
        if visible_err:
            print("\n".join(visible_err), file=sys.stderr)

    if suppressed:
        print(f"  [info] Suppressed {suppressed} known MuPDF structure warning(s).")

    if process.returncode != 0:
        raise subprocess.CalledProcessError(process.returncode, cmd, output=stdout, stderr=stderr)

    return process


def get_file_hash(filepath: Path) -> str:
    hasher = hashlib.md5()
    hasher.update(filepath.read_bytes())
    return hasher.hexdigest()


def _find_soffice() -> str:
    soffice_cmd = "soffice"
    if shutil.which(soffice_cmd):
        return soffice_cmd
    windows_path = Path("C:/Program Files/LibreOffice/program/soffice.exe")
    if windows_path.exists():
        return str(windows_path)
    raise RuntimeError("LibreOffice soffice was not found in PATH or the default Windows install path.")


def _extract_direct_pptx_text(pptx_file: Path, output_file: Path) -> None:
    try:
        run_command([sys.executable, str(STAGE2_DIR / "extract_pptx_text_direct.py"), str(pptx_file), str(output_file)])
    except subprocess.CalledProcessError as exc:
        output_file.write_text(
            f"[WARN] direct PPTX text extraction failed.\n{exc.stderr or exc.output or exc}\n",
            encoding="utf-8",
        )
        print("  [WARN] direct PPTX text extraction failed; continuing.")


def stage1_analyze(template_file: Path, output_dir: Path) -> None:
    template_file = template_file.resolve()
    output_dir = output_dir.resolve()

    if not template_file.exists():
        print(f"[error] Template not found: {template_file}")
        sys.exit(1)

    template_name = template_file.stem
    output_dir.mkdir(parents=True, exist_ok=True)

    spec_file = output_dir / f"{template_name}_spec.md"
    hash_file = output_dir / f".{template_name}_hash"
    current_hash = get_file_hash(template_file)

    if spec_file.exists() and hash_file.exists():
        if hash_file.read_text(encoding="utf-8").strip() == current_hash:
            print("CACHE_HIT: template analysis already exists for this file hash.")
            return

    print(f"CACHE_MISS: starting template analysis (hash: {current_hash})")

    unpacked_dir = output_dir / "unpacked"
    extract_txt = output_dir / "placeholders.txt"

    print("\n[Step 1/5] Extracting exact PPTX text boxes.")
    if unpacked_dir.exists():
        shutil.rmtree(unpacked_dir)
    run_command([sys.executable, str(STAGE2_DIR / "unpack_pptx.py"), str(template_file), str(unpacked_dir)])
    run_command(
        [sys.executable, str(STAGE2_DIR / "extract_placeholders_v2.py"), str(unpacked_dir)],
        stdout_file=extract_txt,
    )
    (output_dir / ".source_pptx").write_text(str(template_file), encoding="utf-8")

    print("\n[Step 2/5] Rendering PDF and slide images.")
    pdf_file = output_dir / f"{template_name}.pdf"
    run_command([_find_soffice(), "--headless", "--convert-to", "pdf", str(template_file), "--outdir", str(output_dir)])
    run_command([sys.executable, str(STAGE1_DIR / "pdf_to_images.py"), str(pdf_file), str(output_dir), f"{template_name}_slide"])
    run_command([sys.executable, str(STAGE1_DIR / "make_grids.py"), str(output_dir), f"{template_name}_slide"])

    print("\n[Step 2b/5] Extracting layout geometry and annotated placeholders.")
    run_command(
        [
            sys.executable,
            str(STAGE1_DIR / "extract_layout_metadata.py"),
            str(unpacked_dir),
            str(output_dir),
            f"{template_name}_slide",
        ]
    )

    print("\n[Step 3/5] Extracting template text with markitdown.")
    try:
        run_command(["uv", "run", "markitdown", str(template_file)], stdout_file=output_dir / f"{template_name}_text.md")
    except subprocess.CalledProcessError as exc:
        (output_dir / f"{template_name}_text.md").write_text(
            f"[WARN] markitdown failed; use placeholders.txt and rendered images instead.\n{exc.stderr or exc.output or exc}\n",
            encoding="utf-8",
        )
        print("  [WARN] markitdown failed; continuing with placeholders and rendered images.")
    _extract_direct_pptx_text(template_file, output_dir / f"{template_name}_text_direct.md")

    print("\n[Step 4/5] Running visual and XML analysis.")
    run_command(
        ["node", str(STAGE1_DIR / "analyze_visual.js"), str(template_file), str(output_dir), str(extract_txt)],
        stdout_file=spec_file,
    )

    print("\n[Step 5/5] Writing cache marker.")
    hash_file.write_text(current_hash, encoding="utf-8")
    print(f"[done] Analysis file: {spec_file}")
    print(f"[done] Placeholder list: {extract_txt}")


def _reset_unpacked_from_source(output_dir: Path) -> Path:
    unpacked_dir = output_dir / "unpacked"
    source_ptr = output_dir / ".source_pptx"
    if not source_ptr.exists():
        if not unpacked_dir.exists():
            raise FileNotFoundError(f"Missing .source_pptx and unpacked directory in {output_dir}")
        print("[info] .source_pptx not found; using existing unpacked directory.")
        return unpacked_dir

    source_pptx = Path(source_ptr.read_text(encoding="utf-8-sig").strip())
    if not source_pptx.exists():
        raise FileNotFoundError(f"Source PPTX from .source_pptx does not exist: {source_pptx}")

    print(f"\n--- 0. Reset workspace from source template ({source_pptx.name}) ---")
    if unpacked_dir.exists():
        shutil.rmtree(unpacked_dir)
    run_command([sys.executable, str(STAGE2_DIR / "unpack_pptx.py"), str(source_pptx), str(unpacked_dir)])
    return unpacked_dir


def _pack_and_verify(output_dir: Path) -> None:
    unpacked_dir = output_dir / "unpacked"
    output_pptx = output_dir / "output_draft.pptx"

    print("\n--- 2. Pack draft PPTX ---")
    run_command([sys.executable, str(STAGE2_DIR / "pack_pptx.py"), str(unpacked_dir), str(output_pptx)])

    print("\n--- 3. Run structural XML audit ---")
    run_command([sys.executable, str(STAGE2_DIR / "audit_injection.py"), str(unpacked_dir)])

    print("\n--- 4. Extract draft text for review ---")
    try:
        run_command(["uv", "run", "markitdown", str(output_pptx)], stdout_file=output_dir / "output_draft_text.md")
    except subprocess.CalledProcessError as exc:
        (output_dir / "output_draft_text.md").write_text(
            f"[WARN] markitdown failed; verify this deck from rendered QA images.\n{exc.stderr or exc.output or exc}\n",
            encoding="utf-8",
        )
        print("  [WARN] markitdown failed; continuing because PPTX pack and structural audit passed.")
    _extract_direct_pptx_text(output_pptx, output_dir / "output_draft_text_direct.md")
    print(f"\n[done] Draft deck: {output_pptx}")


def stage2_build(output_dir: Path, inject_script: Path) -> None:
    output_dir = output_dir.resolve()
    inject_script = inject_script.resolve()
    if not inject_script.exists():
        raise FileNotFoundError(f"Injection script not found: {inject_script}")

    unpacked_dir = _reset_unpacked_from_source(output_dir)
    print("\n--- 1. Run Python injection script ---")
    run_command([sys.executable, str(inject_script), str(unpacked_dir)])
    _pack_and_verify(output_dir)


def stage2_build_json(output_dir: Path, mapping_json: Path) -> None:
    output_dir = output_dir.resolve()
    mapping_json = mapping_json.resolve()
    if not mapping_json.exists():
        raise FileNotFoundError(f"Mapping JSON not found: {mapping_json}")

    unpacked_dir = _reset_unpacked_from_source(output_dir)
    print("\n--- 1. Apply mapping.json ---")
    run_command([sys.executable, str(STAGE2_DIR / "apply_mapping_json.py"), str(unpacked_dir), str(mapping_json)])
    _pack_and_verify(output_dir)


def stage3_audit(pptx_file: Path, output_dir: Path) -> None:
    pptx_file = pptx_file.resolve()
    output_dir = output_dir.resolve()

    if not pptx_file.exists():
        print(f"[error] PPTX not found: {pptx_file}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    qa_dir = output_dir / "qa"

    print("\n--- 1. Run Vision QA ---")
    run_command([sys.executable, str(BASE_DIR / "vision_qa.py"), str(pptx_file), str(qa_dir)])

    report_path = qa_dir / "vision_qa_report.json"
    if report_path.exists():
        print(f"\n[done] QA report: {report_path}")
    else:
        print("\n[warn] QA report was not created.")


def _qa_passed(output_dir: Path) -> bool:
    report_path = output_dir / "qa" / "vision_qa_report.json"
    if not report_path.exists():
        return False
    try:
        data = json.loads(report_path.read_text(encoding="utf-8"))
    except Exception:
        return False
    if "blocking_issue_count" in data:
        try:
            return int(data.get("blocking_issue_count", 1)) == 0
        except (TypeError, ValueError):
            return False
    return data.get("pass") is True


def stage_finalize(output_dir: Path, force: bool = False) -> None:
    output_dir = output_dir.resolve()
    unpacked_dir = output_dir / "unpacked"

    if not unpacked_dir.exists():
        print(f"[error] unpacked directory not found: {unpacked_dir}")
        sys.exit(1)

    if not force and not _qa_passed(output_dir):
        print("[blocked] QA has not passed. Run audit and fix issues before finalizing.")
        print("          Use --force only when the user explicitly wants to package a failing draft.")
        sys.exit(2)

    print("\n--- 1. Clean orphaned resources ---")
    run_command([sys.executable, str(STAGE2_DIR / "clean_orphans.py"), str(unpacked_dir)])

    output_pptx = output_dir / "output_final.pptx"
    print(f"\n--- 2. Pack final PPTX: {output_pptx.name} ---")
    run_command([sys.executable, str(STAGE2_DIR / "pack_pptx.py"), str(unpacked_dir), str(output_pptx)])
    print(f"\n[done] Final deck: {output_pptx}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Presentation Agent Orchestrator")
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_analyze = subparsers.add_parser("analyze", help="Stage 1: analyze template")
    p_analyze.add_argument("template", type=Path, help="PPTX template path")
    p_analyze.add_argument("output", type=Path, help="Output directory")

    p_build = subparsers.add_parser("build", help="Stage 2: build with Python injection script")
    p_build.add_argument("output", type=Path, help="Output directory")
    p_build.add_argument("inject_script", type=Path, help="Python injection script")

    p_build_json = subparsers.add_parser("build-json", help="Stage 2: build with mapping.json")
    p_build_json.add_argument("output", type=Path, help="Output directory")
    p_build_json.add_argument("mapping_json", type=Path, help="Mapping JSON file")

    p_audit = subparsers.add_parser("audit", help="Stage 3: visual QA")
    p_audit.add_argument("pptx", type=Path, help="Draft PPTX path")
    p_audit.add_argument("output", type=Path, help="Output directory")

    p_finalize = subparsers.add_parser("finalize", help="Finalize deck after passing QA")
    p_finalize.add_argument("output", type=Path, help="Output directory")
    p_finalize.add_argument("--force", action="store_true", help="Finalize even when QA has not passed")

    args = parser.parse_args()

    if args.command == "analyze":
        stage1_analyze(args.template, args.output)
    elif args.command == "build":
        stage2_build(args.output, args.inject_script)
    elif args.command == "build-json":
        stage2_build_json(args.output, args.mapping_json)
    elif args.command == "audit":
        stage3_audit(args.pptx, args.output)
    elif args.command == "finalize":
        stage_finalize(args.output, force=args.force)


if __name__ == "__main__":
    main()
