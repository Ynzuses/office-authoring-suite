#!/usr/bin/env python3
"""Best-effort, non-destructive native Excel open/calculation check."""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from pathlib import Path


def run_osascript(script: str, timeout: int):
    return subprocess.run(["osascript", "-e", script], text=True, capture_output=True, timeout=timeout)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--timeout", type=int, default=30)
    args = parser.parse_args(argv)
    result = {
        "status": "native_open_unverified",
        "path": str(args.workbook),
        "sha256": hashlib.sha256(args.workbook.read_bytes()).hexdigest() if args.workbook.is_file() else None,
        "opened_by_script": False,
    }
    if sys.platform != "darwin":
        result["reason"] = "macOS is required"
    elif not args.workbook.is_file():
        result["reason"] = "workbook does not exist"
    elif args.report and args.report.exists():
        result["reason"] = "refusing to overwrite an existing report"
    else:
        path = str(args.workbook.resolve())
        # AppleScript values are quoted as AppleScript string literals; the path is
        # still passed read-only and links are deliberately left at Excel defaults.
        literal = '"' + path.replace("\\", "\\\\").replace('"', '\\"') + '"'
        script = f'''on normalizePath(fileName)
    if fileName starts with "/" then return fileName
    return POSIX path of (fileName as alias)
end normalizePath
tell application "Microsoft Excel"
    activate
    set targetPath to {literal}
    set foundWorkbook to missing value
    repeat with candidate in every workbook
        try
            if my normalizePath(full name of candidate as text) is targetPath then
                set foundWorkbook to candidate
                exit repeat
            end if
        end try
    end repeat
    set openedHere to false
    if foundWorkbook is missing value then
        set foundWorkbook to open workbook workbook file name targetPath read only true add to mru false
        set openedHere to true
    end if
    set workbookPath to my normalizePath(full name of foundWorkbook as text)
    set sheetCount to count of worksheets of foundWorkbook
    calculate full rebuild
    if openedHere then close foundWorkbook saving no
    return workbookPath & "|" & (sheetCount as text) & "|" & (openedHere as text)
end tell'''
        try:
            proc = run_osascript(script, args.timeout)
            result["stdout"] = proc.stdout.strip()
            result["stderr"] = proc.stderr.strip()
            result["returncode"] = proc.returncode
            if proc.returncode == 0 and proc.stdout.strip():
                fields = proc.stdout.strip().split("|")
                result.update({"status": "passed", "excel_full_name": fields[0], "sheet_count": int(fields[1]), "opened_by_script": fields[2] == "true"})
            else:
                result["reason"] = "AppleScript did not confirm a normal open and calculation"
        except (OSError, subprocess.TimeoutExpired) as exc:
            result["reason"] = str(exc)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
