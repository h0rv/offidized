# offidized-schema

Generated Rust types for all OOXML schema elements (SpreadsheetML, WordprocessingML, PresentationML).

Part of [offidized](../../README.md).

## Overview

Types are generated at build time by `offidized-codegen` from the same JSON schema data that .NET Open XML SDK uses. This provides typed access to every element in the OOXML spec while preserving unknown elements via `RawXmlNode` for roundtrip fidelity.

Never hand-write schema types — they are always generated.
