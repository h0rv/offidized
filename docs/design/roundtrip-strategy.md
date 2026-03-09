# Roundtrip Strategy

## The Problem

OOXML files (.xlsx, .docx, .pptx) are ZIP packages containing hundreds of XML parts, relationships, and embedded resources. Any library that parses and re-serializes these files risks losing elements it doesn't understand — custom XML, vendor extensions, newer spec features, embedded media.

## The Approach

The key insight from Microsoft's Open XML SDK: any XML element your code doesn't explicitly model gets preserved as a raw XML fragment on read and written back verbatim on save.

In offidized, this is implemented via `RawXmlNode`:

1. **Parsing**: When the parser encounters an XML element it doesn't have a typed struct for, it captures the entire element (including children and attributes) as a `RawXmlNode`.
2. **Writing**: During serialization, after writing all known typed fields, any captured `RawXmlNode` children are written back in their original position.
3. **Dirty tracking**: If a part hasn't been modified, its original ZIP bytes are copied directly — no parse/serialize roundtrip at all.

## What This Means

- You can implement 20% of the spec and still have 100% roundtrip fidelity
- Users can always drop to the `offidized-schema` typed XML layer for niche features
- New features are additive — they never break existing roundtrip behavior

## Clone-and-Strip Save

When saving a modified file:

1. Clone the source package (ZIP archive)
2. Remove parts that are "owned" (will be rewritten)
3. Rebuild owned parts from the in-memory model
4. Write the result — unmodified parts survive as original bytes

## Testing

Roundtrip fidelity is tested by:

- Opening real-world files from Excel/Word/PowerPoint
- Making minimal modifications
- Saving and comparing: everything untouched must survive unchanged
- The `differential_corpus.sh` script runs Rust vs C# roundtrip comparison across all test files
