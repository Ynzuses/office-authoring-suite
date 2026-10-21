#!/usr/bin/env python3
"""Link the owned Office authoring skills into an explicit local discovery directory.

Dry-run by default. Refuses conflicts instead of overwriting another install.
No network access, global configuration edits, or third-party installation.
"""
import argparse
import hashlib
import json
from pathlib import Path

NAMES = ("word-authoring", "powerpoint-authoring", "excel-authoring", "research-document-delivery")


def deploy(destination, apply=False):
    source_root = Path(__file__).resolve().parent.parent
    destination = destination.expanduser().absolute()
    plan = []
    for name in NAMES:
        source, target = source_root / name, destination / name
        if not (source / "SKILL.md").is_file():
            raise ValueError(f"Incomplete source: {source}")
        if target.exists() or target.is_symlink():
            if not target.is_symlink() or target.resolve() != source.resolve():
                raise FileExistsError(f"Conflicting installation; inspect before changing: {target}")
        files = {str(p.relative_to(source)): hashlib.sha256(p.read_bytes()).hexdigest()
                 for p in sorted(source.rglob("*")) if p.is_file() and "__pycache__" not in p.parts}
        plan.append({"name": name, "source": str(source), "target": str(target), "sha256": files})
    if apply:
        destination.mkdir(parents=True, exist_ok=True)
        for item in plan:
            target = Path(item["target"])
            if not target.is_symlink():
                target.symlink_to(item["source"], target_is_directory=True)
    return {"applied": apply, "skills": plan}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skills-root", type=Path, required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    try:
        print(json.dumps(deploy(args.skills_root, args.apply), ensure_ascii=False, indent=2))
    except (OSError, ValueError) as exc:
        parser.exit(1, str(exc) + "\n")
