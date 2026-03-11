root := parent_directory(source_directory())

# Publishable Rust crates in dependency-safe order.
cargo_publish_order := "offidized-opc offidized-codegen offidized-formula offidized-schema offidized-docx offidized-pptx offidized-xlsx offidized offidized-docview offidized-pptview offidized-xlview offidized-ffi offidized-ir offidized-cli offidized-wasm"

default:
    @just --list

# Validate local tooling, manifest versions, and required schema data.
requirements: _check-tools versions _check-openxml-data
    @echo "Release prerequisites look sane."

# Validate registry credentials for crates.io, PyPI, and npm.
auth: auth-crates auth-python auth-npm
    @echo "Release auth looks configured."

# Validate crates.io credentials.
auth-crates:
    #!/usr/bin/env bash
    set -euo pipefail

    if [[ -n "${CARGO_REGISTRY_TOKEN:-}" ]] || [[ -f "${HOME}/.cargo/credentials.toml" ]] || [[ -f "${HOME}/.cargo/credentials" ]]; then
        echo "crates.io auth looks configured."
        exit 0
    fi

    echo "Missing crates.io auth: set CARGO_REGISTRY_TOKEN or run cargo login" >&2
    exit 1

# Validate PyPI credentials.
auth-python:
    #!/usr/bin/env bash
    set -euo pipefail

    if [[ -n "${MATURIN_PYPI_TOKEN:-}" ]] || { [[ -n "${MATURIN_USERNAME:-}" ]] && [[ -n "${MATURIN_PASSWORD:-}" ]]; }; then
        echo "PyPI auth looks configured."
        exit 0
    fi

    echo "Missing PyPI auth: set MATURIN_PYPI_TOKEN or both MATURIN_USERNAME and MATURIN_PASSWORD" >&2
    exit 1

# Validate npm credentials.
auth-npm:
    #!/usr/bin/env bash
    set -euo pipefail

    if [[ -n "${NPM_TOKEN:-}" ]] || [[ -n "${NODE_AUTH_TOKEN:-}" ]] || [[ -f "${HOME}/.npmrc" ]] || [[ -f "{{root}}/.npmrc" ]]; then
        echo "npm auth looks configured."
        exit 0
    fi

    echo "Missing npm auth: set NPM_TOKEN/NODE_AUTH_TOKEN or configure ~/.npmrc" >&2
    exit 1

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

# Validate the crates.io publish flow before publication.
# Later crates cannot use `cargo publish --dry-run` until earlier internal
# versions exist on crates.io, so they fall back to `cargo package --list`.
crates-dry-run:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    publishable_now=(
        offidized-opc
        offidized-codegen
        offidized-formula
    )
    package_only=(
        offidized-schema
        offidized-docx
        offidized-pptx
        offidized-xlsx
        offidized
        offidized-docview
        offidized-pptview
        offidized-xlview
        offidized-ffi
        offidized-ir
        offidized-cli
        offidized-wasm
    )

    total=$(( ${#publishable_now[@]} + ${#package_only[@]} ))
    step=0

    for crate in "${publishable_now[@]}"; do
        step=$((step + 1))
        echo "[${step}/${total}] cargo publish --dry-run -p ${crate}"
        cargo publish --dry-run --locked -p "${crate}"
    done

    for crate in "${package_only[@]}"; do
        step=$((step + 1))
        echo "[${step}/${total}] cargo package --list -p ${crate}"
        cargo package --allow-dirty --list --locked -p "${crate}" >/dev/null
    done

# Publish the Python bindings package to PyPI.
python:
    cd {{root}}/crates/offidized-py && uvx maturin publish --release

# Validate the Python release flow without uploading.
python-dry-run:
    cd {{root}}/crates/offidized-py && uvx maturin build --release

# Publish the npm-ready wasm packages.
npm:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    echo "[1/2] npm publish offidized-wasm"
    cd crates/offidized-wasm
    wasm-pack build --target web --release --out-dir pkg
    npm publish ./pkg --access public
    cd "{{root}}"

    echo "[2/2] npm publish offidized-xlview"
    cd crates/offidized-xlview
    npm ci
    wasm-pack build --target web --release --out-dir pkg
    node scripts/patch-pkg.mjs
    node build.mjs
    npm publish ./pkg --access public

# Validate the npm-ready wasm packages with npm publish dry-runs.
npm-dry-run:
    #!/usr/bin/env bash
    set -euo pipefail
    cd "{{root}}"

    echo "[1/2] npm publish --dry-run offidized-wasm"
    cd crates/offidized-wasm
    wasm-pack build --target web --release --out-dir pkg
    npm publish --dry-run ./pkg --access public
    cd "{{root}}"

    echo "[2/2] npm publish --dry-run offidized-xlview"
    cd crates/offidized-xlview
    npm ci
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

    cargo metadata --format-version 1 --no-deps | node -e '
        const fs = require("fs");
        const path = require("path");
        const metadata = JSON.parse(fs.readFileSync(0, "utf8"));
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
        const workspaceRoot = path.resolve(root);
        const versions = [
            ["crates/offidized-py/pyproject.toml", parseTomlVersion(read("crates/offidized-py/pyproject.toml"), "crates/offidized-py/pyproject.toml")],
            ["crates/offidized-docview/package.json", JSON.parse(read("crates/offidized-docview/package.json")).version],
            ["crates/offidized-pptview/package.json", JSON.parse(read("crates/offidized-pptview/package.json")).version],
            ["crates/offidized-xlview/package.json", JSON.parse(read("crates/offidized-xlview/package.json")).version],
        ];

        const mismatches = versions.filter(([, version]) => version !== workspaceVersion);

        const workspaceMemberIds = new Set(metadata.workspace_members);
        const internalPrefix = "^" + workspaceVersion;

        for (const pkg of metadata.packages) {
            if (!workspaceMemberIds.has(pkg.id)) {
                continue;
            }
            if (Array.isArray(pkg.publish) && pkg.publish.length === 0) {
                continue;
            }
            if (pkg.version !== workspaceVersion) {
                mismatches.push([path.relative(workspaceRoot, pkg.manifest_path), pkg.version]);
            }

            for (const dep of pkg.dependencies) {
                if (!dep.path) {
                    continue;
                }
                if (!path.resolve(dep.path).startsWith(workspaceRoot + path.sep)) {
                    continue;
                }
                if (!dep.name.startsWith("offidized-")) {
                    continue;
                }
                if (dep.req !== internalPrefix) {
                    mismatches.push([
                        path.relative(workspaceRoot, pkg.manifest_path) + ` -> ${dep.name}`,
                        dep.req,
                    ]);
                }
            }
        }

        if (mismatches.length > 0) {
            console.error(`Workspace version is ${workspaceVersion}, but these package manifests differ:`);
            for (const [relativePath, version] of mismatches) {
                console.error(`  - ${relativePath}: ${version}`);
            }
            process.exit(1);
        }

        console.log(`Version check passed: ${workspaceVersion}`);
    ' "{{root}}"
