#!/usr/bin/env python3
"""
SUPP Report Generator — Claude Skills API

Generates a professional Word document report from SUPP survey data
using Claude's code execution sandbox and docx skill.

Usage:
    python generate_report.py "School Name" survey.csv MTF2024.csv
    python generate_report.py "School Name" survey.csv MTF2024.csv --date "Fall 2025"
    python generate_report.py "School Name" survey.csv MTF2024.csv --output-dir reports/
"""

import argparse
import os
import sys
import time
from pathlib import Path

import anthropic

# Load .env for ANTHROPIC_API_KEY
try:
    from dotenv import load_dotenv

    # Check local, then skills/, then repo root
    for env_path in [".env", "../../.env", "../../../.env"]:
        full = os.path.join(os.path.dirname(os.path.abspath(__file__)), env_path)
        if os.path.exists(full):
            load_dotenv(full)
            break
except ImportError:
    pass  # dotenv not installed, rely on shell env

from report_prompt import SYSTEM_PROMPT, build_prompt

BETA_HEADERS = [
    "code-execution-2025-08-25",
    "files-api-2025-04-14",
    "skills-2025-10-02",
]


# ---------------------------------------------------------------------------
# File helpers (inlined from skills/file_utils.py for Python 3.9 compat)
# ---------------------------------------------------------------------------


def extract_file_ids(response):
    """Extract all file IDs from a Claude API response."""
    file_ids = []
    for block in response.content:
        if block.type == "bash_code_execution_tool_result":
            try:
                if hasattr(block, "content") and hasattr(block.content, "content"):
                    for item in block.content.content:
                        if hasattr(item, "file_id"):
                            file_ids.append(item.file_id)
            except Exception:
                continue
    # Deduplicate preserving order
    seen = set()
    return [fid for fid in file_ids if not (fid in seen or seen.add(fid))]


def download_file(client, file_id, output_path):
    """Download a single file from Files API."""
    Path(os.path.dirname(output_path) or ".").mkdir(parents=True, exist_ok=True)
    file_content = client.beta.files.download(file_id=file_id)
    with open(output_path, "wb") as f:
        f.write(file_content.read())
    size = os.path.getsize(output_path)
    return size


def download_all_files(client, response, output_dir="outputs", prefix=""):
    """Extract and download all files from a response."""
    file_ids = extract_file_ids(response)
    results = []
    for i, file_id in enumerate(file_ids, 1):
        try:
            meta = client.beta.files.retrieve_metadata(file_id=file_id)
            filename = meta.filename
        except Exception:
            filename = f"file_{i}.bin"

        if prefix:
            filename = f"{prefix}{filename}"

        output_path = os.path.join(output_dir, filename)
        try:
            size = download_file(client, file_id, output_path)
            results.append({"path": output_path, "size": size, "ok": True})
        except Exception as e:
            results.append({"path": output_path, "size": 0, "ok": False, "error": str(e)})

    return results


# ---------------------------------------------------------------------------
# Upload / Generate / Download
# ---------------------------------------------------------------------------


def upload_file(client, filepath, filename=None, mime_type=None):
    """Upload a file to the Files API and return the file object."""
    if filename is None:
        filename = os.path.basename(filepath)
    if mime_type is None:
        ext = os.path.splitext(filepath)[1].lower()
        mime_map = {
            ".csv": "text/csv",
            ".py": "text/x-python",
            ".txt": "text/plain",
            ".json": "application/json",
        }
        mime_type = mime_map.get(ext, "application/octet-stream")

    with open(filepath, "rb") as f:
        uploaded = client.beta.files.upload(file=(filename, f, mime_type))

    print(f"  Uploaded {filename} ({uploaded.size_bytes:,} bytes) -> {uploaded.id}")
    return uploaded


def generate_report(
    school_name,
    survey_csv,
    mtf_csv,
    output_dir="outputs",
    survey_date="2025-2026",
    grades="8th, 9th, 10th, 11th, 12th",
    model="claude-opus-4-6",
):
    """Run the full report generation pipeline."""
    client = anthropic.Anthropic()
    processor_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "supp_processor.py")

    # Validate inputs
    for path, label in [
        (survey_csv, "Survey CSV"),
        (mtf_csv, "MTF CSV"),
        (processor_path, "supp_processor.py"),
    ]:
        if not os.path.exists(path):
            print(f"Error: {label} not found: {path}")
            sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    # 1. Upload files
    print("\n1. Uploading files to Claude...")
    survey_file = upload_file(client, survey_csv, "survey_data.csv")
    mtf_file = upload_file(client, mtf_csv, "MTF2024.csv")
    processor_file = upload_file(client, processor_path, "supp_processor.py")

    # 2. Build prompt
    print("\n2. Building analysis prompt...")
    prompt = build_prompt(school_name, survey_date, grades)
    print(f"   Prompt: {len(prompt):,} chars")

    # 3. Call API
    print(f"\n3. Generating report with {model}...")
    print("   (This typically takes 2-5 minutes)")
    start = time.time()

    response = client.beta.messages.create(
        model=model,
        max_tokens=16384,
        system=SYSTEM_PROMPT,
        container={
            "skills": [
                {"type": "anthropic", "skill_id": "docx", "version": "latest"},
            ]
        },
        tools=[{"type": "code_execution_20250825", "name": "code_execution"}],
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "container_upload", "file_id": survey_file.id},
                    {"type": "container_upload", "file_id": mtf_file.id},
                    {"type": "container_upload", "file_id": processor_file.id},
                ],
            }
        ],
        betas=BETA_HEADERS,
    )

    elapsed = time.time() - start
    print(f"   Done in {elapsed:.1f}s (stop_reason: {response.stop_reason})")

    # 4. Show Claude's text output
    print("\n4. Claude's analysis notes:")
    for block in response.content:
        if block.type == "text":
            text = block.text.strip()
            if text:
                print(f"   {text[:800]}{'...' if len(text) > 800 else ''}")

    # 5. Download generated files
    print(f"\n5. Downloading report to {output_dir}/")
    file_ids = extract_file_ids(response)
    print(f"   Found {len(file_ids)} file(s)")

    if file_ids:
        safe_name = school_name.replace(" ", "_")
        results = download_all_files(
            client, response, output_dir=output_dir, prefix=f"{safe_name}_SUPP_"
        )

        # Print summary
        for r in results:
            if r["ok"]:
                kb = r["size"] / 1024
                print(f"   Downloaded: {r['path']} ({kb:.1f} KB)")
            else:
                print(f"   FAILED: {r['path']} — {r.get('error', 'unknown')}")

        # Return path to docx
        for r in results:
            if r["ok"] and r["path"].endswith(".docx"):
                return r["path"]
    else:
        print("   No files generated. Check output above.")
        print("\n   Response block types:")
        for i, block in enumerate(response.content):
            btype = block.type
            extra = ""
            if btype == "bash_code_execution_tool_result":
                if hasattr(block, "content") and hasattr(block.content, "return_code"):
                    extra = f" (return_code={block.content.return_code})"
            print(f"   [{i}] {btype}{extra}")

    return None


def main():
    parser = argparse.ArgumentParser(
        description="Generate a SUPP survey report using Claude Skills API"
    )
    parser.add_argument("school_name", help="Name of the school")
    parser.add_argument("survey_csv", help="Path to the Typeform survey CSV export")
    parser.add_argument("mtf_csv", help="Path to the MTF benchmark CSV")
    parser.add_argument(
        "--output-dir", default="outputs", help="Output directory (default: outputs/)"
    )
    parser.add_argument(
        "--date", default="2025-2026", help="Survey date(s) (default: 2025-2026)"
    )
    parser.add_argument(
        "--grades",
        default="8th, 9th, 10th, 11th, 12th",
        help="Grade levels (default: 8th, 9th, 10th, 11th, 12th)",
    )
    parser.add_argument(
        "--model",
        default="claude-opus-4-6",
        help="Claude model to use (default: claude-opus-4-6)",
    )

    args = parser.parse_args()

    print("=" * 60)
    print("SUPP Report Generator")
    print("=" * 60)
    print(f"School:  {args.school_name}")
    print(f"Survey:  {args.survey_csv}")
    print(f"MTF:     {args.mtf_csv}")
    print(f"Date:    {args.date}")
    print(f"Grades:  {args.grades}")
    print(f"Model:   {args.model}")

    result = generate_report(
        school_name=args.school_name,
        survey_csv=args.survey_csv,
        mtf_csv=args.mtf_csv,
        output_dir=args.output_dir,
        survey_date=args.date,
        grades=args.grades,
        model=args.model,
    )

    if result:
        print(f"\nReport saved to: {result}")
        print("Open in Microsoft Word to review.")
    else:
        print("\nReport generation failed. See output above.")
        sys.exit(1)


if __name__ == "__main__":
    main()
