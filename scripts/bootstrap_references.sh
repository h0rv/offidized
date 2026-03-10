#!/usr/bin/env bash
set -euo pipefail
IFS=$'\n\t'

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd -- "${script_dir}/.." && pwd)"

REPO_NAMES=(
	"ClosedXML"
	"OfficeIMO"
	"ShapeCrawler"
	"npoi"
	"Open-XML-SDK"
)

REPO_URLS=(
	"https://github.com/ClosedXML/ClosedXML.git"
	"https://github.com/EvotecIT/OfficeIMO.git"
	"https://github.com/ShapeCrawler/ShapeCrawler.git"
	"https://github.com/nissl-lab/npoi.git"
	"https://github.com/dotnet/Open-XML-SDK.git"
)

die() {
	printf 'error: %s\n' "$*" >&2
	exit 1
}

log() {
	local name="$1"
	shift
	printf '[%s] %s\n' "$name" "$*"
}

detect_default_jobs() {
	local detected=""

	if command -v getconf >/dev/null 2>&1; then
		detected="$(getconf _NPROCESSORS_ONLN 2>/dev/null || true)"
	fi

	if [[ -z "${detected}" ]] && command -v sysctl >/dev/null 2>&1; then
		detected="$(sysctl -n hw.ncpu 2>/dev/null || true)"
	fi

	if ! [[ "${detected}" =~ ^[1-9][0-9]*$ ]]; then
		detected=4
	fi

	if ((detected > ${#REPO_NAMES[@]})); then
		detected=${#REPO_NAMES[@]}
	fi

	if ((detected < 1)); then
		detected=1
	fi

	printf '%s\n' "${detected}"
}

references_dir="${repo_root}/references"
jobs="$(detect_default_jobs)"
shallow_clone=1
force_non_git=0

usage() {
	cat <<EOF
Bootstrap reference repositories into ${repo_root}/references.

Usage:
  $(basename "$0") [options]

Options:
  -j, --jobs N           Number of parallel clone/update jobs (default: ${jobs})
  --references-dir PATH  Target references directory (default: ${repo_root}/references)
  --full-clone           Disable shallow clone for newly cloned repositories
  --force-non-git        If a target path exists and is not a git repo, remove it and clone
  -h, --help             Show this help text

Behavior:
  - Missing repo directory: clone it.
  - Existing git repo: fetch + fast-forward pull (no merge commits).
  - Existing non-git path: skip by default (non-destructive).

Examples:
  $(basename "$0")
  $(basename "$0") --jobs 3
  $(basename "$0") --force-non-git
EOF
}

while (($# > 0)); do
	case "$1" in
	-j | --jobs)
		[[ $# -ge 2 ]] || die "missing value for $1"
		jobs="$2"
		shift 2
		;;
	--jobs=*)
		jobs="${1#*=}"
		shift
		;;
	--references-dir)
		[[ $# -ge 2 ]] || die "missing value for $1"
		references_dir="$2"
		shift 2
		;;
	--references-dir=*)
		references_dir="${1#*=}"
		shift
		;;
	--full-clone)
		shallow_clone=0
		shift
		;;
	--force-non-git)
		force_non_git=1
		shift
		;;
	-h | --help)
		usage
		exit 0
		;;
	*)
		die "unknown option: $1 (run with --help)"
		;;
	esac
done

[[ "${jobs}" =~ ^[1-9][0-9]*$ ]] || die "--jobs must be a positive integer"

if ! command -v git >/dev/null 2>&1; then
	die "git is required but was not found in PATH"
fi

mkdir -p "${references_dir}"

tmp_root="$(mktemp -d "${TMPDIR:-/tmp}/bootstrap_references.XXXXXX")"
status_dir="${tmp_root}/status"
mkdir -p "${status_dir}"
trap 'rm -rf "${tmp_root}"' EXIT

run_repo() {
	local name="$1"
	local url="$2"
	local repo_dir="$3"
	local status_file="$4"
	local status="FAILED"
	local branch=""
	local clone_args=()

	trap 'printf "%s\n" "${status}" > "${status_file}"' EXIT

	if [[ -d "${repo_dir}" ]]; then
		if git -C "${repo_dir}" rev-parse --git-dir >/dev/null 2>&1; then
			log "${name}" "updating existing clone in ${repo_dir}"
			git -C "${repo_dir}" fetch --prune origin
			branch="$(git -C "${repo_dir}" symbolic-ref --quiet --short HEAD 2>/dev/null || true)"
			if [[ -n "${branch}" ]]; then
				git -C "${repo_dir}" pull --ff-only --no-rebase origin "${branch}"
				log "${name}" "updated (fast-forward) on branch ${branch}"
			else
				log "${name}" "detached HEAD detected; fetched origin only"
			fi
			status="UPDATED"
			return 0
		fi

		if ((force_non_git == 0)); then
			log "${name}" "skipped: ${repo_dir} exists but is not a git repository"
			status="SKIPPED"
			return 0
		fi

		log "${name}" "removing non-git directory due to --force-non-git: ${repo_dir}"
		rm -rf "${repo_dir}"
	elif [[ -e "${repo_dir}" ]]; then
		if ((force_non_git == 0)); then
			log "${name}" "skipped: ${repo_dir} exists but is not a directory/git repository"
			status="SKIPPED"
			return 0
		fi

		log "${name}" "removing non-directory path due to --force-non-git: ${repo_dir}"
		rm -rf "${repo_dir}"
	fi

	log "${name}" "cloning ${url}"
	clone_args=(clone)
	if ((shallow_clone == 1)); then
		clone_args+=(--depth 1)
	fi
	clone_args+=("${url}" "${repo_dir}")
	git "${clone_args[@]}"
	status="CLONED"
	log "${name}" "clone complete"
}

wait_for_slot() {
	while :; do
		local running_jobs
		running_jobs="$(jobs -pr | wc -l | tr -d '[:space:]')"
		if ((running_jobs < jobs)); then
			break
		fi
		sleep 0.1
	done
}

printf 'Bootstrapping reference repositories in %s\n' "${references_dir}"
printf 'Parallel jobs: %s\n' "${jobs}"
if ((shallow_clone == 1)); then
	printf 'Clone mode: shallow (--depth 1 for new clones)\n'
else
	printf 'Clone mode: full\n'
fi
printf '\n'

declare -a pids=()
declare -a pid_names=()

for i in "${!REPO_NAMES[@]}"; do
	name="${REPO_NAMES[$i]}"
	url="${REPO_URLS[$i]}"
	repo_dir="${references_dir}/${name}"
	status_file="${status_dir}/${name}.status"

	wait_for_slot

	(
		run_repo "${name}" "${url}" "${repo_dir}" "${status_file}"
	) &
	pids+=("$!")
	pid_names+=("${name}")
done

overall_rc=0
for i in "${!pids[@]}"; do
	if ! wait "${pids[$i]}"; then
		overall_rc=1
	fi
done

cloned=0
updated=0
skipped=0
failed=0

printf '\nStatus Summary\n'
for name in "${REPO_NAMES[@]}"; do
	status_path="${status_dir}/${name}.status"
	status="FAILED"
	if [[ -f "${status_path}" ]]; then
		status="$(<"${status_path}")"
	fi

	case "${status}" in
	CLONED)
		((cloned += 1))
		;;
	UPDATED)
		((updated += 1))
		;;
	SKIPPED)
		((skipped += 1))
		;;
	*)
		((failed += 1))
		overall_rc=1
		;;
	esac

	printf '  %-14s %s\n' "${name}" "${status}"
done

printf '\nTotals: cloned=%d updated=%d skipped=%d failed=%d\n' \
	"${cloned}" "${updated}" "${skipped}" "${failed}"

if ((skipped > 0)); then
	printf 'Note: skipped entries were left untouched (use --force-non-git to replace them).\n'
fi

exit "${overall_rc}"
