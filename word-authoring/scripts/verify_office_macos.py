#!/usr/bin/env python3
"""Normal-open and PDF-export evidence for an exact DOCX/PPTX on macOS.

Does not click recovery dialogs, modify source files, or claim visual review.
Requires installed Microsoft Office and OS automation authorization.
"""
import argparse
import hashlib
import json
from pathlib import Path
import platform
import subprocess
from datetime import datetime, timezone

from check_office_package import check

WORD = '''on run argv
  set inputPath to item 1 of argv
  set pdfPath to item 2 of argv
  tell application "Microsoft Word"
    with timeout of 45 seconds
      repeat with existingDoc in documents
        if (posix full name of existingDoc) is inputPath then error "Target already open; preserve it and verify manually or close it first."
      end repeat
      open file name inputPath read only true add to recent files false repair false
      set taskDoc to active document
      if (posix full name of taskDoc) is not inputPath then error "Opened document identity mismatch."
      set itemCount to count paragraphs of taskDoc
      set appVersion to version
      try
        save as taskDoc file name pdfPath file format format PDF
      on error errorMessage number errorNumber
        close taskDoc saving no
        error errorMessage number errorNumber
      end try
      close taskDoc saving no
      return "NORMAL_OPEN_OK|" & itemCount & "|" & appVersion
    end timeout
  end tell
end run
'''



def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify(source, pdf, report, grant_export_dir=False):
    source = source.resolve(strict=True)
    pdf, report = pdf.absolute(), report.absolute()
    if source.suffix.lower() not in (".docx", ".pptx"):
        raise ValueError("Input must end in .docx or .pptx")
    if platform.system() != "Darwin":
        raise ValueError("This helper requires macOS; use native validation on the target platform.")
    if pdf.suffix.lower() != ".pdf" or report.suffix.lower() != ".json":
        raise ValueError("Use .pdf and .json output paths")
    if source == pdf or pdf == report:
        raise ValueError("Input and output paths must be distinct")
    for output in (pdf, report):
        if output.exists() or output.is_symlink():
            raise FileExistsError(f"Refusing to overwrite {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
    structure = check(source)
    if structure["status"] != "passed":
        raise ValueError("Structural preflight failed: " + json.dumps(structure, ensure_ascii=False))
    before = digest(source)
    result = {"source": str(source), "source_sha256": before,
              "checked_at_utc": datetime.now(timezone.utc).isoformat(),
              "native_open": "failed", "visual_review": "not_performed",
              "content_review": "not_performed"}
    try:
        if source.suffix.lower() == ".pptx":
            from powerpoint_macos import verify_ppt
            verify_ppt(source, pdf, result, grant_export_dir=grant_export_dir)
            result["source_unchanged"] = before == digest(source)
            if not result["source_unchanged"]:
                raise RuntimeError("Source bytes changed during native validation")
            with report.open("x", encoding="utf-8") as stream:
                json.dump(result, stream, ensure_ascii=False, indent=2)
            return result
        script = WORD
        run = subprocess.run(["osascript", "-", str(source), str(pdf)], input=script,
                             text=True, capture_output=True, timeout=55)
        if run.returncode != 0 or not run.stdout.strip().startswith("NORMAL_OPEN_OK|"):
            raise RuntimeError(run.stderr.strip() or run.stdout.strip() or "No normal-open result")
        if before != digest(source):
            raise RuntimeError("Source bytes changed during native validation")
        if not pdf.is_file() or not pdf.read_bytes().startswith(b"%PDF-"):
            raise RuntimeError("Native PDF export is missing or invalid")
        _, count, version = run.stdout.strip().split("|", 2)
        result.update(native_open="passed", app_version=version, object_count=int(count),
                      count_unit="paragraphs" if source.suffix.lower() == ".docx" else "slides",
                      pdf=str(pdf), pdf_sha256=digest(pdf), source_unchanged=True)
    except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as e:
        result["error"] = str(e)
        result["source_unchanged"] = source.is_file() and before == digest(source)
    with report.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2)
    return result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("file", type=Path)
    p.add_argument("--pdf", type=Path, required=True)
    p.add_argument("--report", type=Path, required=True)
    p.add_argument("--grant-export-dir", action="store_true", help="PPT only: handle a recognised Office access prompt solely for the helper-created temporary export folder")
    a = p.parse_args()
    try:
        result = verify(a.file, a.pdf, a.report, a.grant_export_dir)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["native_open"] == "passed" and "error" not in result else 1
    except (OSError, ValueError) as e:
        p.exit(1, str(e) + "\n")


if __name__ == "__main__":
    raise SystemExit(main())
