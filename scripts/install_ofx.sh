#!/usr/bin/env bash
set -euo pipefail

usage() {
	cat <<'EOF'
Install the ofx CLI from GitHub release binaries.

Usage:
  install_ofx.sh [--version <tag|latest>] [--repo <owner/repo>] [--bin-dir <dir>]

Examples:
  install_ofx.sh
  install_ofx.sh --version v0.1.0
  install_ofx.sh --bin-dir "$HOME/.local/bin"
EOF
}

version="latest"
repo="h0rv/offidized"
bin_dir="${HOME}/.local/bin"

while [[ $# -gt 0 ]]; do
	case "$1" in
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

if ! command -v curl >/dev/null 2>&1; then
	echo "curl is required" >&2
	exit 1
fi

if ! command -v tar >/dev/null 2>&1; then
	echo "tar is required" >&2
	exit 1
fi

os="$(uname -s)"
arch="$(uname -m)"

case "$os" in
Darwin) target_os="apple-darwin" ;;
Linux) target_os="unknown-linux-gnu" ;;
*)
	echo "unsupported operating system: $os" >&2
	exit 1
	;;
esac

case "$arch" in
x86_64 | amd64) target_arch="x86_64" ;;
arm64 | aarch64) target_arch="aarch64" ;;
*)
	echo "unsupported architecture: $arch" >&2
	exit 1
	;;
esac

target="${target_arch}-${target_os}"
archive_name="ofx-${target}.tar.gz"
checksums_name="ofx-checksums.txt"

release_base="https://github.com/${repo}/releases"
if [[ "$version" == "latest" ]]; then
	asset_url="${release_base}/latest/download/${archive_name}"
	checksums_url="${release_base}/latest/download/${checksums_name}"
else
	asset_url="${release_base}/download/${version}/${archive_name}"
	checksums_url="${release_base}/download/${version}/${checksums_name}"
fi

tmp_dir="$(mktemp -d)"
cleanup() {
	rm -rf "$tmp_dir"
}
trap cleanup EXIT

archive_path="${tmp_dir}/${archive_name}"
checksums_path="${tmp_dir}/${checksums_name}"
extract_dir="${tmp_dir}/extract"

echo "Downloading ${asset_url}" >&2
curl -fsSL "$asset_url" -o "$archive_path"

if curl -fsSL "$checksums_url" -o "$checksums_path"; then
	if command -v shasum >/dev/null 2>&1; then
		expected="$(grep "  ${archive_name}\$" "$checksums_path" | awk '{print $1}')"
		if [[ -n "${expected}" ]]; then
			actual="$(shasum -a 256 "$archive_path" | awk '{print $1}')"
			if [[ "$actual" != "$expected" ]]; then
				echo "checksum mismatch for ${archive_name}" >&2
				exit 1
			fi
		fi
	fi
else
	echo "warning: could not download ${checksums_url}; skipping checksum verification" >&2
fi

mkdir -p "$extract_dir"
tar -xzf "$archive_path" -C "$extract_dir"

binary_path="$(find "$extract_dir" -type f -name ofx | head -n 1)"
if [[ -z "$binary_path" ]]; then
	echo "downloaded archive did not contain an ofx binary" >&2
	exit 1
fi

mkdir -p "$bin_dir"
install_path="${bin_dir}/ofx"
cp "$binary_path" "$install_path"
chmod +x "$install_path"

echo "installed ofx to ${install_path}"
case ":${PATH}:" in
*":${bin_dir}:"*) ;;
*)
	echo "warning: ${bin_dir} is not on PATH" >&2
	;;
esac
