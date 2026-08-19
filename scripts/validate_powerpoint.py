import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_DIR))

from pptx_utils import configure_utf8_output


def validate_powerpoint(pptx_path: Path, report_path: Path | None = None) -> dict:
    pptx_path = Path(pptx_path).resolve()
    if not pptx_path.exists():
        report = {
            "pass": False,
            "pptx": str(pptx_path),
            "slide_count": None,
            "error": f"PPTX not found: {pptx_path}",
        }
        _write_report(report, report_path)
        return report

    script = f"""
$ErrorActionPreference = 'Stop'
$path = [System.IO.Path]::GetFullPath('{_ps_quote(str(pptx_path))}')
$ppt = $null
$pres = $null
try {{
  $ppt = New-Object -ComObject PowerPoint.Application
  $pres = $ppt.Presentations.Open($path, $true, $false, $false)
  $count = $pres.Slides.Count
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  Write-Output (@{{ pass = $true; pptx = $path; slide_count = $count; error = $null }} | ConvertTo-Json -Compress)
}} catch {{
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  Write-Output (@{{ pass = $false; pptx = $path; slide_count = $null; error = $_.Exception.Message }} | ConvertTo-Json -Compress)
  exit 2
}} finally {{
  if ($pres -ne $null) {{ $pres.Close() | Out-Null }}
  if ($ppt -ne $null) {{ $ppt.Quit() | Out-Null }}
}}
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except FileNotFoundError as exc:
        report = {
            "pass": False,
            "pptx": str(pptx_path),
            "slide_count": None,
            "error": f"PowerShell unavailable: {exc}",
        }
        _write_report(report, report_path)
        return report

    raw = (result.stdout or "").strip().splitlines()
    payload = raw[-1] if raw else ""
    try:
        report = json.loads(payload)
    except json.JSONDecodeError:
        report = {
            "pass": False,
            "pptx": str(pptx_path),
            "slide_count": None,
            "error": (result.stderr or result.stdout or "PowerPoint validation produced no JSON output").strip(),
        }

    _write_report(report, report_path)
    return report


def _ps_quote(value: str) -> str:
    return value.replace("'", "''")


def _write_report(report: dict, report_path: Path | None) -> None:
    if report_path:
        report_path = Path(report_path)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")


def main() -> None:
    configure_utf8_output()
    parser = argparse.ArgumentParser(description="Validate a PPTX by opening it through the PowerPoint COM DOM")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--report", type=Path, default=None)
    args = parser.parse_args()

    if os.environ.get("PRESENTATION_AGENT_SKIP_POWERPOINT_VALIDATION") == "1":
        report = {
            "pass": True,
            "pptx": str(args.pptx.resolve()),
            "slide_count": None,
            "skipped": True,
            "error": None,
        }
        _write_report(report, args.report)
        print("[validate] skipped because PRESENTATION_AGENT_SKIP_POWERPOINT_VALIDATION=1")
        return

    report = validate_powerpoint(args.pptx, args.report)
    if report.get("pass"):
        print(f"[validate] PowerPoint opened file; slide_count={report.get('slide_count')}")
        sys.exit(0)

    print(f"[validate] PowerPoint validation failed: {report.get('error')}")
    sys.exit(2)


if __name__ == "__main__":
    main()
