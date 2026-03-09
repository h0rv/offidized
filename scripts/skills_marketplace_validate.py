#!/usr/bin/env python3
"""Validate skills marketplace metadata and package layout."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILLS_DIR = ROOT / "skills"
REGISTRY_PATH = SKILLS_DIR / "registry.json"


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def validate() -> list[str]:
    errors: list[str] = []

    if not REGISTRY_PATH.exists():
        return [f"missing registry: {REGISTRY_PATH}"]

    registry = load_json(REGISTRY_PATH)
    skills = registry.get("skills")
    if not isinstance(skills, list):
        return ["registry.skills must be an array"]

    for item in skills:
        if not isinstance(item, dict):
            errors.append("registry entry must be an object")
            continue

        skill_id = item.get("id")
        manifest_rel = item.get("manifest")
        version = item.get("version")

        if not isinstance(skill_id, str) or not skill_id:
            errors.append("skill id missing/invalid")
            continue
        if not isinstance(manifest_rel, str) or not manifest_rel:
            errors.append(f"{skill_id}: manifest missing/invalid")
            continue

        manifest_path = SKILLS_DIR / manifest_rel
        if not manifest_path.exists():
            errors.append(f"{skill_id}: manifest not found at {manifest_path}")
            continue

        manifest = load_json(manifest_path)
        if manifest.get("id") != skill_id:
            errors.append(f"{skill_id}: manifest id mismatch ({manifest.get('id')!r})")
        if version is not None and manifest.get("version") != version:
            errors.append(
                f"{skill_id}: version mismatch registry={version} manifest={manifest.get('version')}"
            )

        targets = manifest.get("targets")
        if not isinstance(targets, dict):
            errors.append(f"{skill_id}: manifest.targets must be object")
            continue

        for target_name in ("codex", "claude-code"):
            target = targets.get(target_name)
            if not isinstance(target, dict):
                errors.append(f"{skill_id}: missing target {target_name}")
                continue

            rel_path = target.get("path")
            entry = target.get("entry")
            if not isinstance(rel_path, str) or not rel_path:
                errors.append(f"{skill_id}:{target_name}: invalid path")
                continue
            if not isinstance(entry, str) or not entry:
                errors.append(f"{skill_id}:{target_name}: invalid entry")
                continue

            target_dir = manifest_path.parent / rel_path
            if not target_dir.exists():
                errors.append(f"{skill_id}:{target_name}: missing dir {target_dir}")
                continue

            entry_path = target_dir / entry
            if not entry_path.exists():
                errors.append(f"{skill_id}:{target_name}: missing entry {entry_path}")

            if target_name == "codex":
                if not (target_dir / "SKILL.md").exists():
                    errors.append(f"{skill_id}:codex missing SKILL.md")

    return errors


def main() -> int:
    errors = validate()
    if errors:
        print("skills marketplace validation failed:", file=sys.stderr)
        for err in errors:
            print(f"- {err}", file=sys.stderr)
        return 1
    print("skills marketplace validation passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
