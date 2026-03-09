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
