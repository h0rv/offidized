# Validate skills registry/manifests/payload layout.
validate:
    python3 scripts/skills_marketplace_validate.py

# Install a skill payload for Codex.
install-codex skill:
    python3 scripts/skills_marketplace_install.py --skill "{{skill}}" --target codex

# Install a skill payload for Claude Code.
install-claude skill:
    python3 scripts/skills_marketplace_install.py --skill "{{skill}}" --target claude-code
