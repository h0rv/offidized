# What Is Needed For A Codex + Claude Code Skills Marketplace

## P0 Requirements

1. Packaging standard

- Stable skill id and semver version.
- Per-target payload mapping (`codex`, `claude-code`).
- Deterministic install paths.

2. Registry and discovery

- Machine-readable index with id, version, manifest path, tags.
- Lightweight metadata for search/filter.

3. Validation gate

- CI validator for registry/manifests/payload existence.
- Reject publish when required target files are missing.

4. Install UX

- One command install by skill id + target runtime.
- Safe overwrite and clear destination path output.

## P1 Requirements

1. Distribution

- Build immutable skill archives (`.tar.gz`/`.zip`).
- Serve registry + artifacts over HTTPS/CDN.

2. Trust and integrity

- Checksum file per artifact.
- Optional signature verification before install.

3. Compatibility matrix

- Runtime/version constraints (Codex client version, Claude Code version).
- Deprecation + migration notes.

4. Quality

- Golden tests for install and invocation templates.
- Lint for prompt length and forbidden patterns.

## P2 Requirements

1. Ranking and analytics

- Download counts, stars, health score.
- Version adoption metrics.

2. Publisher workflow

- Namespaces/owners, ownership transfer, release channels.
- Automated CI publish pipeline.

3. Security policy

- Static scanning for embedded secrets/unsafe commands.
- Review workflow for high-risk skills.

## Current State In This Repo

- `skills/registry.json` provides index/discovery.
- `skills/catalog/offidized-mcp/manifest.json` provides package metadata.
- `scripts/skills_marketplace_validate.py` provides validation gate.
- `scripts/skills_marketplace_install.py` provides runtime install command.

This is enough for internal/private marketplace beta.
