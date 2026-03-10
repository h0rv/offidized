#!/usr/bin/env bash
set -euo pipefail

usage() {
	cat <<'EOF'
Install an offidized marketplace skill directly from GitHub without cloning.

Usage:
  install_skill.sh --skill <skill-id> --target <codex|claude-code> [--ref <git-ref>] [--repo <owner/repo>] [--dest <dir>]

Examples:
  install_skill.sh --skill offidized-cli --target codex
  install_skill.sh --skill offidized-cli --target claude-code --ref v0.1.0
EOF
}

skill=""
target=""
ref="main"
repo="h0rv/offidized"
dest=""

while [[ $# -gt 0 ]]; do
	case "$1" in
	--skill)
		skill="${2:-}"
		shift 2
		;;
	--target)
		target="${2:-}"
		shift 2
		;;
	--ref)
		ref="${2:-}"
		shift 2
		;;
	--repo)
		repo="${2:-}"
		shift 2
		;;
	--dest)
		dest="${2:-}"
		shift 2
		;;
	-h | --help)
		usage
		exit 0
		;;
	*)
		echo "unknown argument: $1" >&2
		usage >&2
		exit 1
		;;
	esac
done

if [[ -z "$skill" || -z "$target" ]]; then
	usage >&2
	exit 1
fi

if [[ "$target" != "codex" && "$target" != "claude-code" ]]; then
	echo "target must be one of: codex, claude-code" >&2
	exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
	echo "python3 is required" >&2
	exit 1
fi

tmp_dir="$(mktemp -d)"
cleanup() {
	rm -rf "$tmp_dir"
}
trap cleanup EXIT

archive_path="${tmp_dir}/repo.tar.gz"

download_archive() {
	local url="$1"
	if curl -fsSL "$url" -o "$archive_path"; then
		return 0
	fi
	return 1
}

archive_url=""
for candidate in \
	"https://github.com/${repo}/archive/refs/heads/${ref}.tar.gz" \
	"https://github.com/${repo}/archive/refs/tags/${ref}.tar.gz" \
	"https://github.com/${repo}/archive/${ref}.tar.gz"; do
	echo "Trying ${candidate}" >&2
	if download_archive "$candidate"; then
		archive_url="$candidate"
		break
	fi
done

if [[ -z "$archive_url" ]]; then
	echo "failed to download repository archive for ref ${ref}" >&2
	exit 1
fi

echo "Downloaded ${archive_url}" >&2
tar -xzf "$archive_path" -C "$tmp_dir"

repo_root="$(find "$tmp_dir" -mindepth 1 -maxdepth 1 -type d | head -n 1)"
if [[ -z "$repo_root" ]]; then
	echo "failed to unpack repository archive" >&2
	exit 1
fi

python3 - "$repo_root" "$skill" "$target" "$dest" <<'PY'
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

repo_root = Path(sys.argv[1])
skill_id = sys.argv[2]
target = sys.argv[3]
dest_arg = sys.argv[4]

registry_path = repo_root / "skills" / "registry.json"
registry = json.loads(registry_path.read_text())

entry = next((item for item in registry.get("skills", []) if item.get("id") == skill_id), None)
if entry is None:
    raise SystemExit(f"unknown skill id: {skill_id}")

manifest_rel = entry.get("manifest")
if not isinstance(manifest_rel, str) or not manifest_rel:
    raise SystemExit(f"invalid manifest for skill: {skill_id}")

manifest_path = repo_root / "skills" / manifest_rel
manifest = json.loads(manifest_path.read_text())
targets = manifest.get("targets")
if not isinstance(targets, dict) or target not in targets:
    raise SystemExit(f"target {target} not supported by {skill_id}")

target_cfg = targets[target]
rel_path = target_cfg.get("path")
if not isinstance(rel_path, str) or not rel_path:
    raise SystemExit(f"invalid target path for {skill_id}:{target}")

source_dir = manifest_path.parent / rel_path
if not source_dir.exists():
    raise SystemExit(f"missing payload: {source_dir}")

if dest_arg:
    base_dest = Path(dest_arg).expanduser()
else:
    home = Path.home()
    base_dest = home / (".codex" if target == "codex" else ".claude") / "skills"

install_dir = base_dest / skill_id
install_dir.parent.mkdir(parents=True, exist_ok=True)
if install_dir.exists():
    shutil.rmtree(install_dir)

shutil.copytree(source_dir, install_dir)
print(f"installed {skill_id} for {target} at {install_dir}")
PY
