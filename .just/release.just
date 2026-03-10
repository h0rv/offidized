root := parent_directory(source_directory())

# Publishable Rust crates in dependency-safe order.
cargo_publish_order := "offidized-opc offidized-codegen offidized-formula offidized-schema offidized-docx offidized-pptx offidized-xlsx offidized offidized-docview offidized-pptview offidized-xlview offidized-ffi offidized-ir offidized-cli offidized-wasm"

default:
    @just --list

# Validate local tooling, manifest versions, and required schema data.
requirements: _check-tools versions _check-openxml-data
    @echo "Release prerequisites look sane."

# Validate registry credentials for crates.io, PyPI, and npm.
auth:
    #!/usr/bin/env bash
    set -euo pipefail

    missing=()

    if [[ -z "${CARGO_REGISTRY_TOKEN:-}" ]] && [[ ! -f "${HOME}/.cargo/credentials.toml" ]] && [[ ! -f "${HOME}/.cargo/credentials" ]]; then
        missing+=("crates.io: set CARGO_REGISTRY_TOKEN or run cargo login")
    fi

    if [[ -z "${MATURIN_PYPI_TOKEN:-}" ]] && [[ -z "${TWINE_PASSWORD:-}" ]]; then
        missing+=("PyPI: set MATURIN_PYPI_TOKEN (preferred) or TWINE_PASSWORD/TWINE_USERNAME")
    fi

    if [[ -z "${NPM_TOKEN:-}" ]] && [[ -z "${NODE_AUTH_TOKEN:-}" ]] && [[ ! -f "${HOME}/.npmrc" ]] && [[ ! -f "{{root}}/.npmrc" ]]; then
        missing+=("npm: set NPM_TOKEN/NODE_AUTH_TOKEN or configure ~/.npmrc")
    fi

    if (( ${#missing[@]} > 0 )); then
        printf 'Missing release auth:\n' >&2
        printf '  - %s\n' "${missing[@]}" >&2
        exit 1
    fi

    echo "Release auth looks configured."

# Validate the full release flow without publishing.
dry-run: requirements crates-dry-run python-dry-run npm-dry-run
    @echo "Release dry-run completed."

# Publish everything: crates.io first, then PyPI, then npm.
all: requirements auth crates python npm
    @echo "Released crates.io, PyPI, and npm artifacts."

# Publish all Rust crates to crates.io.
crates:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    IFS=' ' read -r -a crates <<< "{{cargo_publish_order}}"
    total="${#crates[@]}"
    settle_seconds="${CRATES_IO_SETTLE_SECONDS:-30}"

    for index in "${!crates[@]}"; do
        crate="${crates[$index]}"
        human_index=$((index + 1))
        echo "[${human_index}/${total}] cargo publish -p ${crate}"
        cargo publish --locked -p "${crate}"
        if (( human_index < total )); then
            echo "Waiting ${settle_seconds}s for crates.io index propagation..."
            sleep "${settle_seconds}"
        fi
    done

# Validate the crates.io publish flow with cargo dry-runs.
crates-dry-run:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    IFS=' ' read -r -a crates <<< "{{cargo_publish_order}}"
    total="${#crates[@]}"

    for index in "${!crates[@]}"; do
        crate="${crates[$index]}"
        human_index=$((index + 1))
        echo "[${human_index}/${total}] cargo publish --dry-run -p ${crate}"
        cargo publish --dry-run --locked -p "${crate}"
    done

# Publish the Python bindings package to PyPI.
python:
    cd {{root}}/crates/offidized-py && uvx maturin publish --release

# Validate the Python release flow without uploading.
python-dry-run:
    cd {{root}}/crates/offidized-py && uvx maturin publish --dry-run

# Publish all generated wasm/npm packages.
npm:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    echo "[1/4] npm publish offidized-wasm"
    cd crates/offidized-wasm
    wasm-pack build --target web --release --out-dir pkg
    npm publish ./pkg --access public
    cd "{{root}}"

    echo "[2/4] npm publish offidized-docview (generated pkg)"
    cd crates/offidized-docview
    wasm-pack build --target web --release --out-dir pkg
    npm publish ./pkg --access public
    cd "{{root}}"

    echo "[3/4] npm publish offidized-pptview (generated pkg)"
    cd crates/offidized-pptview
    wasm-pack build --target web --release --out-dir pkg
    npm publish ./pkg --access public
    cd "{{root}}"

    echo "[4/4] npm publish offidized-xlview"
    cd crates/offidized-xlview
    wasm-pack build --target web --release --out-dir pkg
    node scripts/patch-pkg.mjs
    node build.mjs
    npm publish ./pkg --access public

# Validate all generated wasm/npm packages with npm publish dry-runs.
npm-dry-run:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    echo "[1/4] npm publish --dry-run offidized-wasm"
    cd crates/offidized-wasm
    wasm-pack build --target web --release --out-dir pkg
    npm publish --dry-run ./pkg --access public
    cd "{{root}}"

    echo "[2/4] npm publish --dry-run offidized-docview (generated pkg)"
    cd crates/offidized-docview
    wasm-pack build --target web --release --out-dir pkg
    npm publish --dry-run ./pkg --access public
    cd "{{root}}"

    echo "[3/4] npm publish --dry-run offidized-pptview (generated pkg)"
    cd crates/offidized-pptview
    wasm-pack build --target web --release --out-dir pkg
    npm publish --dry-run ./pkg --access public
    cd "{{root}}"

    echo "[4/4] npm publish --dry-run offidized-xlview"
    cd crates/offidized-xlview
    wasm-pack build --target web --release --out-dir pkg
    node scripts/patch-pkg.mjs
    node build.mjs
    npm publish --dry-run ./pkg --access public

[private]
_check-tools:
    #!/usr/bin/env bash
    set -euo pipefail

    missing=()
    for cmd in cargo npm node python3 uv wasm-pack; do
        if ! command -v "${cmd}" >/dev/null 2>&1; then
            missing+=("${cmd}")
        fi
    done

    if (( ${#missing[@]} > 0 )); then
        printf 'Missing release tools:\n' >&2
        printf '  - %s\n' "${missing[@]}" >&2
        exit 1
    fi

[private]
_check-openxml-data:
    #!/usr/bin/env bash
    set -euo pipefail

    default_root="{{root}}/references/Open-XML-SDK/data"
    data_root="${OFFIDIZED_OPENXML_DATA_ROOT:-$default_root}"

    if [[ ! -d "${data_root}" ]]; then
        echo "Missing Open XML SDK schema data at ${data_root}" >&2
        echo "Set OFFIDIZED_OPENXML_DATA_ROOT or run ./scripts/bootstrap_references.sh" >&2
        exit 1
    fi

versions:
    #!/usr/bin/env bash
    set -euo pipefail

    node -e '
        const fs = require("fs");
        const path = require("path");
        const root = process.argv[1];

        const read = (relativePath) =>
            fs.readFileSync(path.join(root, relativePath), "utf8");
        const parseTomlVersion = (text, relativePath) => {
            const match = text.match(/version\s*=\s*"([^"]+)"/);
            if (!match) {
                throw new Error(`Could not find version in ${relativePath}`);
            }
            return match[1];
        };

        const workspaceCargo = read("Cargo.toml");
        const workspaceBlock = workspaceCargo.match(/\[workspace\.package\][\s\S]*?(?=\n\[|$)/);
        if (!workspaceBlock) {
            throw new Error("Missing [workspace.package] block in Cargo.toml");
        }

        const workspaceVersion = parseTomlVersion(workspaceBlock[0], "Cargo.toml");
        const versions = [
            ["crates/offidized-py/pyproject.toml", parseTomlVersion(read("crates/offidized-py/pyproject.toml"), "crates/offidized-py/pyproject.toml")],
            ["crates/offidized-docview/package.json", JSON.parse(read("crates/offidized-docview/package.json")).version],
            ["crates/offidized-pptview/package.json", JSON.parse(read("crates/offidized-pptview/package.json")).version],
            ["crates/offidized-xlview/package.json", JSON.parse(read("crates/offidized-xlview/package.json")).version],
        ];

        const mismatches = versions.filter(([, version]) => version !== workspaceVersion);
        if (mismatches.length > 0) {
            console.error(`Workspace version is ${workspaceVersion}, but these package manifests differ:`);
            for (const [relativePath, version] of mismatches) {
                console.error(`  - ${relativePath}: ${version}`);
            }
            process.exit(1);
        }

        console.log(`Version check passed: ${workspaceVersion}`);
    ' "{{root}}"
