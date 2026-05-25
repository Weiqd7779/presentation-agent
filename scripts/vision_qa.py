import argparse
import json
import subprocess
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import configure_utf8_output, convert_pptx_to_pdf, pdf_to_images

configure_utf8_output()

VISION_CLIENT_JS = SKILL_DIR / "vision_llm_client.js"
BLOCKING_CONFIDENCE_THRESHOLD = 0.70


QA_PROMPT = """
You are auditing rendered PowerPoint slides for delivery-blocking visual defects.
Return JSON only. Do not include Markdown fences or commentary.

Only report concrete, visible issues that affect readability, usability, or correctness.
Do not report subjective polish suggestions, mild spacing preferences, or style opinions.
Be strict about delivery quality: pass only if the rendered slide is usable without manual cleanup.

Use this exact schema:
{
  "pass": true,
  "overall_summary": "Short summary.",
  "issues": [
    {
      "slide": 1,
      "type": "TEXT_OVERFLOW",
      "shape": "Visible shape name if known, otherwise a short location such as top-right title",
      "description": "What is visibly wrong.",
      "severity": "high",
      "confidence": 0.85,
      "blocking": true,
      "evidence": "Specific visual evidence from the slide.",
      "repair_hint": "Smallest practical fix.",
      "verification": "How to verify the original problem is fixed."
    }
  ]
}

Allowed issue types:
- TEXT_OVERFLOW: text is clipped, unreadably dense, or forced too small.
- MAPPING_COLLISION: inserted content appears in the wrong placeholder or overlaps unrelated content.
- LAYOUT_DISPLACEMENT: an element is visibly outside its intended container or off slide.
- FONT_INCONSISTENCY: a clear unintended font-size/style mismatch that harms readability.
- BROKEN_RENDER: blank slide, missing image, badly cropped image, irrelevant leftover template image, or obviously failed render.

Severity rules:
- high: blocks delivery; the audience would clearly notice or lose information.
- medium: should be fixed if easy, but does not block delivery.
- low: minor issue; do not set blocking=true.

Confidence rules:
- Use confidence from 0.0 to 1.0.
- Only set blocking=true when severity is high and confidence >= 0.70.
- Mark high/blocking when a brand or subject image is visibly cropped in a way that looks accidental, when a timeline/date label wraps or collides, or when same-role text boxes use visibly inconsistent font sizes.
- If no blocking issue exists, set pass=true only when the slide can be presented as-is.
"""


def _extract_json(raw: str) -> dict:
    text = raw.strip()
    if "---RESULT_START---" in text:
        text = text.split("---RESULT_START---", 1)[1].split("---RESULT_END---", 1)[0].strip()
    text = text.replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        if start < 0:
            raise
        depth = 0
        in_string = False
        escape = False
        for index in range(start, len(text)):
            char = text[index]
            if in_string:
                if escape:
                    escape = False
                elif char == "\\":
                    escape = True
                elif char == '"':
                    in_string = False
            elif char == '"':
                in_string = True
            elif char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    return json.loads(text[start : index + 1])
        raise


def _is_blocking_issue(issue: dict) -> bool:
    severity = str(issue.get("severity", "")).lower()
    try:
        confidence = float(issue.get("confidence", 0))
    except (TypeError, ValueError):
        confidence = 0.0
    return bool(issue.get("blocking")) and severity == "high" and confidence >= BLOCKING_CONFIDENCE_THRESHOLD


def _normalize_report(report: dict) -> dict:
    issues = report.get("issues", [])
    if not isinstance(issues, list):
        issues = []

    normalized_issues = []
    for issue in issues:
        if not isinstance(issue, dict):
            continue
        has_triage_fields = any(key in issue for key in ("severity", "confidence", "blocking"))
        normalized = {
            "slide": issue.get("slide"),
            "type": issue.get("type", "UNKNOWN"),
            "shape": issue.get("shape", ""),
            "description": issue.get("description", ""),
            "severity": str(issue.get("severity", "high" if not has_triage_fields else "medium")).lower(),
            "confidence": issue.get("confidence", 1.0 if not has_triage_fields else 0.5),
            "blocking": bool(issue.get("blocking", not has_triage_fields)),
            "evidence": issue.get("evidence", ""),
            "repair_hint": issue.get("repair_hint", ""),
            "verification": issue.get("verification", ""),
        }
        normalized["blocking"] = _is_blocking_issue(normalized)
        normalized_issues.append(normalized)

    blocking_count = sum(1 for issue in normalized_issues if issue["blocking"])
    report["issues"] = normalized_issues
    report["blocking_issue_count"] = blocking_count
    report["pass"] = blocking_count == 0
    report.setdefault("overall_summary", "")
    return report


def run_vision_audit(pptx_path: Path, output_dir: Path) -> None:
    pptx_path = Path(pptx_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"  [1/3] Converting {pptx_path.name} to PDF.")
    pdf_path = convert_pptx_to_pdf(pptx_path, output_dir, SKILL_DIR / "scripts")

    print("  [2/3] Rendering PDF pages to images.")
    image_paths = pdf_to_images(pdf_path, output_dir / "slides", dpi=200)

    print("  [3/3] Running Vision LLM audit.")
    temp_prompt_path = output_dir / "vision_prompt.txt"
    temp_prompt_path.write_text(QA_PROMPT, encoding="utf-8")

    image_args = [str(p) for p in image_paths]
    wrapper = f"""
const {{ callVisionLLM }} = require('{VISION_CLIENT_JS.as_posix()}');
const fs = require('fs');
const prompt = fs.readFileSync('{temp_prompt_path.as_posix()}', 'utf8');
const images = {json.dumps(image_args)};

(async () => {{
  try {{
    const result = await callVisionLLM(prompt, images);
    console.log('---RESULT_START---');
    console.log(result);
    console.log('---RESULT_END---');
  }} catch (e) {{
    console.error(e);
    process.exit(1);
  }}
}})();
"""

    wrapper_path = output_dir / "vision_cli_wrapper.js"
    wrapper_path.write_text(wrapper, encoding="utf-8")

    try:
        result = subprocess.run(
            ["node", str(wrapper_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=True,
        )
        report = _normalize_report(_extract_json(result.stdout))
    except subprocess.CalledProcessError as exc:
        details = (exc.stderr or exc.stdout or str(exc)).strip()
        print(f"  [ERROR] Vision LLM audit failed: {details}")
        report = {
            "pass": False,
            "overall_summary": "Vision LLM audit failed.",
            "error": details,
            "issues": [],
            "blocking_issue_count": 1,
        }
    except Exception as exc:
        print(f"  [ERROR] Vision LLM audit failed: {exc}")
        report = {
            "pass": False,
            "overall_summary": "Vision LLM audit failed.",
            "error": str(exc),
            "issues": [],
            "blocking_issue_count": 1,
        }

    report_path = output_dir / "vision_qa_report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    if report.get("pass"):
        print("\n[PASS] Visual QA passed.")
    else:
        issues = report.get("issues", [])
        blocking_count = report.get("blocking_issue_count", 0)
        print(f"\n[FAIL] Visual QA found {blocking_count} blocking issue(s), {len(issues)} total issue(s).")
        for issue in issues:
            print(f"   - [Slide {issue.get('slide')}] {issue.get('type')}: {issue.get('description')}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    run_vision_audit(args.pptx, args.output)
