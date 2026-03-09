# Reference Implementations

offidized's API design and correctness is informed by studying mature C# OOXML libraries. Clone these into `references/` for development:

```bash
cd references/
git clone --depth 1 https://github.com/ClosedXML/ClosedXML.git
git clone --depth 1 https://github.com/EvotecIT/OfficeIMO.git
git clone --depth 1 https://github.com/ShapeCrawler/ShapeCrawler.git
git clone --depth 1 https://github.com/nissl-lab/npoi.git
git clone --depth 1 https://github.com/dotnet/Open-XML-SDK.git
```

Or use the bootstrap script: `./scripts/bootstrap_references.sh`

## What to Study Where

| Need to understand...                 | Look at...                                          |
| ------------------------------------- | --------------------------------------------------- |
| **Excel**                             |                                                     |
| xlsx API design, cell/range model     | ClosedXML/ClosedXML/Excel/                          |
| xlsx IO (reading/writing XML parts)   | ClosedXML/ClosedXML/Excel/IO/                       |
| xlsx style system                     | ClosedXML/ClosedXML/Excel/Style/                    |
| xlsx formula engine                   | ClosedXML/ClosedXML/Excel/CalcEngine/               |
| xlsx pivot tables                     | ClosedXML/ClosedXML/Excel/PivotTables/              |
| **Word**                              |                                                     |
| docx API design                       | OfficeIMO/OfficeIMO.Word/                           |
| docx tables, lists, sections          | OfficeIMO/OfficeIMO.Word/                           |
| docx headers/footers, TOC             | OfficeIMO/OfficeIMO.Word/                           |
| **PowerPoint**                        |                                                     |
| pptx API design, shape model          | ShapeCrawler/src/ShapeCrawler/                      |
| pptx slides, layouts, masters         | ShapeCrawler/src/ShapeCrawler/                      |
| pptx tables, charts, text frames      | ShapeCrawler/src/ShapeCrawler/                      |
| **Multi-format (fallback reference)** |                                                     |
| xlsx via POI model                    | npoi/main/NPOI.OOXML/XSSF/                          |
| docx via POI model                    | npoi/main/NPOI.OOXML/XWPF/                          |
| pptx via POI model                    | npoi/main/NPOI.OOXML/XSLF/                          |
| **Low-level OOXML foundation**        |                                                     |
| JSON schema data (codegen input)      | Open-XML-SDK/data/                                  |
| Schema codegen approach               | Open-XML-SDK/gen/                                   |
| OPC packaging layer                   | Open-XML-SDK/src/DocumentFormat.OpenXml/Packaging/  |
| Validation system                     | Open-XML-SDK/src/DocumentFormat.OpenXml/Validation/ |

## Why These Libraries

| Library                                                          | Format | Why it's here                                                                                            |
| ---------------------------------------------------------------- | ------ | -------------------------------------------------------------------------------------------------------- |
| [**ClosedXML**](https://github.com/ClosedXML/ClosedXML)          | xlsx   | Best high-level Excel API in any language. Intuitive, well-tested, actively maintained.                  |
| [**OfficeIMO**](https://github.com/EvotecIT/OfficeIMO)           | docx   | Clean, modern C# Word library. Full feature coverage including TOC, watermarks, charts, tracked changes. |
| [**ShapeCrawler**](https://github.com/ShapeCrawler/ShapeCrawler) | pptx   | Modern C# PowerPoint library. Clean shape model, good text/table/chart support.                          |
| [**NPOI**](https://github.com/nissl-lab/npoi)                    | all    | .NET port of Apache POI. Massive coverage including legacy formats. Good fallback for edge cases.        |
| [**Open XML SDK**](https://github.com/dotnet/Open-XML-SDK)       | all    | Microsoft's foundation layer. We use its JSON schema data as codegen input.                              |

## Schema Data

The Open XML SDK's `data/` directory contains the structured OOXML spec as JSON — this is the input to `offidized-codegen`.
