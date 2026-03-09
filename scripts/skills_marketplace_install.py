#!/usr/bin/env python3
"""Install a marketplace skill payload for codex or claude-code."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILLS_DIR = ROOT / "skills"
REGISTRY_PATH = SKILLS_DIR / "registry.json"

DEFAULT_DEST = {
    "codex": Path.home() / ".codex" / "skills",
    "claude-code": Path.home() / ".claude" / "skills",
}


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def resolve_manifest(skill_id: str) -> tuple[Path, dict]:
    registry = load_json(REGISTRY_PATH)
    for item in registry.get("skills", []):
        if item.get("id") == skill_id:
            rel = item.get("manifest")
            if not isinstance(rel, str):
                raise ValueError(f"skill {skill_id} has invalid manifest reference")
            manifest_path = SKILLS_DIR / rel
            if not manifest_path.exists():
                raise FileNotFoundError(f"manifest not found: {manifest_path}")
            return manifest_path, load_json(manifest_path)
    raise KeyError(f"unknown skill id: {skill_id}")


def install(skill_id: str, target: str, dest: Path | None) -> Path:
    manifest_path, manifest = resolve_manifest(skill_id)
    targets = manifest.get("targets")
    if not isinstance(targets, dict) or target not in targets:
        raise ValueError(f"target {target} not supported by {skill_id}")

    target_cfg = targets[target]
    if not isinstance(target_cfg, dict):
        raise ValueError(f"invalid target config for {target}")

    rel_path = target_cfg.get("path")
    if not isinstance(rel_path, str) or not rel_path:
        raise ValueError(f"invalid path for target {target}")

    source_dir = manifest_path.parent / rel_path
    if not source_dir.exists():
        raise FileNotFoundError(f"missing source payload: {source_dir}")

    base_dest = dest if dest is not None else DEFAULT_DEST[target]
    install_dir = base_dest / skill_id
    install_dir.parent.mkdir(parents=True, exist_ok=True)

    if install_dir.exists():
        shutil.rmtree(install_dir)

    shutil.copytree(source_dir, install_dir)
    return install_dir


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Install a skill from skills/registry.json"
    )
    parser.add_argument("--skill", required=True, help="Skill id")
    parser.add_argument(
        "--target",
        required=True,
        choices=["codex", "claude-code"],
        help="Target runtime",
    )
    parser.add_argument(
        "--dest",
        help="Override base destination directory (defaults to ~/.codex/skills or ~/.claude/skills)",
    )
    args = parser.parse_args()

    try:
        dest_override = Path(args.dest).expanduser() if args.dest else None
        installed_to = install(args.skill, args.target, dest_override)
        print(f"installed {args.skill} for {args.target} at {installed_to}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"install failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
