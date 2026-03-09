# offidized-codegen

Code generator that transforms XSD/JSON schema definitions into Rust structs for OOXML elements.

Part of [offidized](../../README.md).

## Overview

Reads schema data from `references/Open-XML-SDK/data/` and generates the typed structs used by `offidized-schema`. Runs at build time (via `build.rs`), not as a proc macro, to keep CI fast.
