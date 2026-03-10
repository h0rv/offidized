# Skills Marketplace (Codex + Claude Code)

This directory is the source of truth for distributable agent skills.

## Layout

- `catalog/<skill-id>/manifest.json`: canonical metadata and compatibility targets.
- `catalog/<skill-id>/codex/`: Codex skill payload (`SKILL.md`, `agents/openai.yaml`, resources).
- `catalog/<skill-id>/claude-code/`: Claude Code payload (slash-command style markdown + metadata).
- `registry.json`: publish index for all skills.
- `schema/skill-manifest.schema.json`: manifest schema.

## Marketplace Requirements

1. Stable skill ID + semantic version.
2. Compatibility metadata per target (`codex`, `claude-code`).
3. Install payload per target in deterministic paths.
4. Validation tooling for manifests + required files.
5. Install tooling for local agent environments.
6. Optional signed package distribution for public marketplace.

## Current Commands

- `just skills-validate`
- `just skills-install-codex <skill>`
- `just skills-install-claude <skill>`

## Direct Install

- Skill only for Codex: `curl -fsSL https://raw.githubusercontent.com/h0rv/offidized/main/scripts/install_skill.sh | bash -s -- --skill offidized-cli --target codex`
- Skill only for Claude Code: `curl -fsSL https://raw.githubusercontent.com/h0rv/offidized/main/scripts/install_skill.sh | bash -s -- --skill offidized-cli --target claude-code`
- CLI only: `curl -fsSL https://raw.githubusercontent.com/h0rv/offidized/main/scripts/install_ofx.sh | bash -s --`
- Full setup for Codex: `curl -fsSL https://raw.githubusercontent.com/h0rv/offidized/main/scripts/install_offidized.sh | bash -s -- --target codex --skill offidized-cli`
- Full setup for Claude Code: `curl -fsSL https://raw.githubusercontent.com/h0rv/offidized/main/scripts/install_offidized.sh | bash -s -- --target claude-code --skill offidized-cli`
- Append `--ref vX.Y.Z` for skill/setup scripts or `--version vX.Y.Z` for the CLI script to pin a release.
