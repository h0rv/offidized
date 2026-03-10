# offidized-cli

Use the `ofx` CLI first for Office file work. Switch to Python only when the job needs real scripting.

## Procedure

1. Prefer `cargo run -p offidized-cli -- <subcommand>` inside this repo.
2. Inspect first with `info`, `read`, or `part --list`.
3. For unified edits, run `nodes`, then `capabilities`, then `edit --lint --strict`.
4. For IR workflows, run `derive`, edit the IR text, then `apply`.
5. Use `-o <path>` by default instead of `-i`.
6. Move to `crates/offidized-py` only for loops, generated content, or repeated transformations.

## Guardrails

- Treat IDs from `nodes` as source of truth.
- Keep `--lint --strict` on unless the user explicitly asks to relax validation.
- Treat `apply --dry-run` as a light check, not a real diff preview.
- Read `references/cli-workflows.md` for command patterns.
- Read `references/python-bindings.md` before using the Python bindings.
