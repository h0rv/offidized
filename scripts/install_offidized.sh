#!/usr/bin/env bash
set -euo pipefail

usage() {
	cat <<'EOF'
Install the offidized agent skill and local tooling in one command.

Usage:
  install_offidized.sh --target <codex|claude-code> [--skill <skill-id>] [--ref <git-ref>] [--version <tag|latest>] [--repo <owner/repo>] [--bin-dir <dir>] [--skip-python]

Examples:
  install_offidized.sh --target codex
  install_offidized.sh --target claude-code --ref v0.1.0
EOF
}

target=""
skill="offidized-cli"
ref="main"
version="latest"
repo="h0rv/offidized"
bin_dir="${HOME}/.local/bin"
skip_python="false"

while [[ $# -gt 0 ]]; do
	case "$1" in
	--target)
		target="${2:-}"
		shift 2
		;;
	--skill)
		skill="${2:-}"
		shift 2
		;;
	--ref)
		ref="${2:-}"
		shift 2
		;;
	--version)
		version="${2:-}"
		shift 2
		;;
	--repo)
		repo="${2:-}"
		shift 2
		;;
	--bin-dir)
		bin_dir="${2:-}"
		shift 2
		;;
	--skip-python)
		skip_python="true"
		shift
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

if [[ "$target" != "codex" && "$target" != "claude-code" ]]; then
	echo "target must be one of: codex, claude-code" >&2
	exit 1
fi

if ! command -v curl >/dev/null 2>&1; then
	echo "curl is required" >&2
	exit 1
fi

tmp_dir="$(mktemp -d)"
cleanup() {
	rm -rf "$tmp_dir"
}
trap cleanup EXIT

run_remote_script() {
	local script_name="$1"
	shift
	local raw_url="https://raw.githubusercontent.com/${repo}/${ref}/scripts/${script_name}"
	local local_path="${tmp_dir}/${script_name}"
	echo "Downloading ${raw_url}" >&2
	curl -fsSL "$raw_url" -o "$local_path"
	chmod +x "$local_path"
	"$local_path" "$@"
}

run_remote_script install_skill.sh --skill "$skill" --target "$target" --ref "$ref" --repo "$repo"
run_remote_script install_ofx.sh --version "$version" --repo "$repo" --bin-dir "$bin_dir"

if [[ "$skip_python" == "false" ]]; then
	if command -v python3 >/dev/null 2>&1 && python3 -m pip --version >/dev/null 2>&1; then
		python3 -m pip install --user --upgrade offidized
	else
		echo "warning: python3 + pip not available; skipping Python package install" >&2
	fi
fi

echo "offidized setup complete"
