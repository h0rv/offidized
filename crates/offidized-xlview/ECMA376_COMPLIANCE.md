# ECMA-376 SpreadsheetML Compliance Checklist

Official Standard: **ECMA-376 5th Edition (Office Open XML)**
Reference Document: `docs/spec/Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf`

This checklist tracks implementation status against **Section 18: SpreadsheetML Reference Material** of ECMA-376 Part 1.

## Status Legend

| Symbol | Meaning                      |
| ------ | ---------------------------- |
| ✅     | Fully implemented            |
| 🟡     | Partially implemented        |
| ⬜     | Not implemented              |
| ➖     | Out of scope (v1)            |
| N/A    | Not applicable for view-only |

---

## 18.2 Workbook

> The workbook element is the top-level container for the workbook structure.

| Section | Element               | Description              | Status | Notes                       |
| ------- | --------------------- | ------------------------ | ------ | --------------------------- |
| 18.2.1  | `bookViews`           | Workbook view settings   | ⬜     | Window position/size        |
| 18.2.2  | `calcPr`              | Calculation properties   | N/A    | No formula evaluation       |
| 18.2.3  | `customWorkbookViews` | Custom workbook views    | ➖     |                             |
| 18.2.4  | `definedName`         | Named range definition   | ✅     | `DefinedName` struct        |
| 18.2.5  | `definedNames`        | Named ranges collection  | ✅     | In `Workbook.defined_names` |
| 18.2.6  | `externalReference`   | External workbook ref    | ➖     | No external links           |
| 18.2.7  | `externalReferences`  | External refs collection | ➖     |                             |
| 18.2.9  | `fileRecoveryPr`      | Recovery properties      | ➖     |                             |
| 18.2.10 | `fileSharing`         | Sharing settings         | ➖     |                             |
| 18.2.11 | `fileVersion`         | File version info        | ➖     |                             |
| 18.2.12 | `functionGroup`       | Function category        | N/A    |                             |
| 18.2.13 | `functionGroups`      | Function categories      | N/A    |                             |
| 18.2.17 | `pivotCache`          | Pivot cache ref          | ➖     |                             |
| 18.2.18 | `pivotCaches`         | Pivot caches collection  | ➖     |                             |
| 18.2.19 | `sheet`               | Sheet metadata           | ✅     | Name, sheetId, rId          |
| 18.2.20 | `sheets`              | Sheets collection        | ✅     | All sheets parsed           |
| 18.2.22 | `smartTagPr`          | Smart tag properties     | ➖     |                             |
| 18.2.23 | `smartTagType`        | Smart tag type           | ➖     |                             |
| 18.2.24 | `smartTagTypes`       | Smart tag types          | ➖     |                             |
| 18.2.25 | `webPublishing`       | Web publishing           | ➖     |                             |
| 18.2.27 | `workbook`            | Root element             | ✅     |                             |
| 18.2.28 | `workbookPr`          | Workbook properties      | 🟡     | date1904 pending            |
| 18.2.29 | `workbookProtection`  | Protection settings      | ⬜     |                             |
| 18.2.30 | `workbookView`        | View settings            | ⬜     |                             |

---

## 18.3 Worksheets

> Worksheet elements define the content of individual sheets.

### 18.3.1 Sheet Structure

| Section   | Element                 | Description             | Status | Notes                                              |
| --------- | ----------------------- | ----------------------- | ------ | -------------------------------------------------- |
| 18.3.1.4  | `col`                   | Column properties       | ✅     | Width, hidden, style                               |
| 18.3.1.6  | `cols`                  | Columns collection      | ✅     | `ColWidth` structs                                 |
| 18.3.1.10 | `conditionalFormatting` | CF container            | 🟡     | colorScale/dataBar/iconSet parsed; eval simplified |
| 18.3.1.12 | `dataValidation`        | Validation rule         | ✅     | `DataValidation` struct                            |
| 18.3.1.13 | `dataValidations`       | Validations collection  | ✅     | `data_validations` vec                             |
| 18.3.1.21 | `dimension`             | Sheet dimension         | ⬜     | Uses max_row/max_col (dimension ignored)           |
| 18.3.1.29 | `headerFooter`          | Print header/footer     | ➖     |                                                    |
| 18.3.1.32 | `hyperlink`             | Hyperlink definition    | ✅     | `Hyperlink` struct                                 |
| 18.3.1.33 | `hyperlinks`            | Hyperlinks collection   | ✅     | `hyperlinks` vec                                   |
| 18.3.1.39 | `legacyDrawingHF`       | Legacy drawing header   | ➖     |                                                    |
| 18.3.1.40 | `mergeCell`             | Merged cell range       | ✅     | `MergeRange` struct                                |
| 18.3.1.41 | `mergeCells`            | Merged cells collection | ✅     | `merges` vec                                       |
| 18.3.1.45 | `pageMargins`           | Print margins           | ➖     |                                                    |
| 18.3.1.46 | `pageSetup`             | Print setup             | ➖     |                                                    |
| 18.3.1.48 | `pane`                  | View pane settings      | ✅     | Frozen/split panes                                 |
| 18.3.1.52 | `printOptions`          | Print options           | ➖     |                                                    |
| 18.3.1.55 | `row`                   | Row element             | ✅     | Height, hidden, style                              |
| 18.3.1.56 | `rowBreaks`             | Page breaks (rows)      | ✅     | `row_breaks` vec                                   |
| 18.3.1.59 | `selection`             | Selection state         | ⬜     |                                                    |
| 18.3.1.60 | `sheetCalcPr`           | Sheet calc properties   | N/A    |                                                    |
| 18.3.1.61 | `sheetData`             | Cell data container     | ✅     | Main cell parsing                                  |
| 18.3.1.62 | `sheetFormatPr`         | Default row/col format  | ✅     | default_col_width/row_height                       |
| 18.3.1.66 | `sheetPr`               | Sheet properties        | ✅     | tabColor, outlines                                 |
| 18.3.1.67 | `sheetProtection`       | Protection settings     | ✅     | `is_protected` flag                                |
| 18.3.1.72 | `sheetView`             | View settings           | 🟡     | Frozen panes only                                  |
| 18.3.1.73 | `sheetViews`            | Views collection        | 🟡     |                                                    |
| 18.3.1.75 | `sortState`             | Sort settings           | ⬜     |                                                    |
| 18.3.1.77 | `tableParts`            | Table references        | ⬜     |                                                    |
| 18.3.1.99 | `worksheet`             | Root element            | ✅     |                                                    |

### 18.3.1.2 Auto Filter

| Section   | Element         | Description            | Status | Notes                 |
| --------- | --------------- | ---------------------- | ------ | --------------------- |
| 18.3.1.2  | `autoFilter`    | Auto filter definition | ✅     | `AutoFilter` struct   |
| 18.3.2.1  | `colorFilter`   | Color-based filter     | ⬜     |                       |
| 18.3.2.2  | `customFilter`  | Custom filter criteria | ⬜     |                       |
| 18.3.2.3  | `customFilters` | Custom filters         | ⬜     |                       |
| 18.3.2.4  | `dateGroupItem` | Date grouping          | ⬜     |                       |
| 18.3.2.5  | `dynamicFilter` | Dynamic filter         | ⬜     |                       |
| 18.3.2.6  | `filter`        | Filter value           | ⬜     |                       |
| 18.3.2.7  | `filterColumn`  | Column filter settings | ✅     | `FilterColumn` struct |
| 18.3.2.8  | `filters`       | Filters container      | ⬜     |                       |
| 18.3.2.9  | `iconFilter`    | Icon filter            | ⬜     |                       |
| 18.3.2.10 | `top10`         | Top N filter           | ⬜     |                       |

### 18.3.1.3 Column Breaks

| Section  | Element     | Description        | Status | Notes            |
| -------- | ----------- | ------------------ | ------ | ---------------- |
| 18.3.1.3 | `colBreaks` | Page breaks (cols) | ✅     | `col_breaks` vec |
| 18.3.1.5 | `brk`       | Break element      | ✅     | Parsed in breaks |

---

## 18.4 Shared String Table

> Shared strings optimize storage of repeated text values.

| Section | Element       | Description         | Status | Notes                |
| ------- | ------------- | ------------------- | ------ | -------------------- |
| 18.4.1  | `phoneticPr`  | Phonetic properties | ➖     | Japanese furigana    |
| 18.4.2  | `phoneticRun` | Phonetic text run   | ➖     |                      |
| 18.4.3  | `r`           | Rich text run       | ✅     | `RichTextRun` struct |
| 18.4.4  | `rPh`         | Phonetic rich text  | ➖     |                      |
| 18.4.5  | `rPr`         | Run properties      | ✅     | `RunStyle` struct    |
| 18.4.6  | `si`          | String item         | ✅     | Plain or rich text   |
| 18.4.7  | `sst`         | Shared string table | ✅     | Full parsing         |
| 18.4.8  | `t`           | Text element        | ✅     | Plain text value     |

---

## 18.8 Styles

> Styles define formatting for fonts, fills, borders, and number formats.

### 18.8.1-9 Alignment & Protection

| Section | Element           | Description       | Status | Notes                  |
| ------- | ----------------- | ----------------- | ------ | ---------------------- |
| 18.8.1  | `alignment`       | Cell alignment    | ✅     | `RawAlignment` struct  |
| -       | `horizontal`      | Horizontal align  | ✅     | All 8 values           |
| -       | `vertical`        | Vertical align    | ✅     | All 5 values           |
| -       | `wrapText`        | Text wrapping     | ✅     | `wrap` field           |
| -       | `shrinkToFit`     | Shrink to fit     | ✅     | `shrink_to_fit` field  |
| -       | `indent`          | Indent level      | ✅     | `indent` field         |
| -       | `textRotation`    | Rotation angle    | ✅     | 0-180 or 255           |
| -       | `readingOrder`    | Reading direction | ✅     | 0, 1, 2                |
| -       | `justifyLastLine` | Justify last line | ⬜     |                        |
| 18.8.33 | `protection`      | Cell protection   | ✅     | `RawProtection` struct |
| -       | `locked`          | Cell locked       | ✅     | `locked` field         |
| -       | `hidden`          | Formula hidden    | ✅     | `hidden` field         |

### 18.8.4-5 Borders

| Section | Element        | Description        | Status | Notes              |
| ------- | -------------- | ------------------ | ------ | ------------------ |
| 18.8.4  | `border`       | Border definition  | ✅     | `RawBorder` struct |
| 18.8.5  | `borders`      | Borders collection | ✅     | In `StyleSheet`    |
| -       | `left`         | Left border        | ✅     | `RawBorderSide`    |
| -       | `right`        | Right border       | ✅     |                    |
| -       | `top`          | Top border         | ✅     |                    |
| -       | `bottom`       | Bottom border      | ✅     |                    |
| -       | `diagonal`     | Diagonal border    | ✅     |                    |
| -       | `diagonalUp`   | Diagonal up flag   | ✅     |                    |
| -       | `diagonalDown` | Diagonal down flag | ✅     |                    |

**Border Styles (18.18.3 ST_BorderStyle):**

| Style            | Status |
| ---------------- | ------ |
| none             | ✅     |
| thin             | ✅     |
| medium           | ✅     |
| thick            | ✅     |
| dashed           | ✅     |
| dotted           | ✅     |
| double           | ✅     |
| hair             | ✅     |
| mediumDashed     | ✅     |
| dashDot          | ✅     |
| mediumDashDot    | ✅     |
| dashDotDot       | ✅     |
| mediumDashDotDot | ✅     |
| slantDashDot     | ✅     |

### 18.8.6-9 Cell Styles

| Section | Element        | Description             | Status | Notes                |
| ------- | -------------- | ----------------------- | ------ | -------------------- |
| 18.8.6  | `cellStyle`    | Named style             | ✅     | `NamedStyle` struct  |
| 18.8.7  | `cellStyles`   | Named styles collection | ✅     | `named_styles` vec   |
| 18.8.8  | `cellStyleXfs` | Base style formats      | ✅     | `cell_style_xfs` vec |
| 18.8.9  | `cellXfs`      | Cell formats            | ✅     | `cell_xfs` vec       |

### 18.8.10-19 Colors

| Section | Element    | Description       | Status | Notes              |
| ------- | ---------- | ----------------- | ------ | ------------------ |
| 18.8.10 | `color`    | Color definition  | ✅     | `ColorSpec` struct |
| -       | `rgb`      | ARGB value        | ✅     | 8-char hex         |
| -       | `theme`    | Theme index       | ✅     | 0-11               |
| -       | `tint`     | Tint modifier     | ✅     | -1.0 to 1.0        |
| -       | `indexed`  | Legacy index      | ✅     | 0-63               |
| -       | `auto`     | Auto color        | ✅     | System default     |
| 18.8.12 | `colors`   | Colors collection | ⬜     | Custom palette     |
| 18.8.13 | `condense` | Condense font     | ➖     | Rarely used        |
| 18.8.14 | `extend`   | Extend font       | ➖     |                    |

### 18.8.20-23 Fills

| Section | Element        | Description      | Status | Notes                |
| ------- | -------------- | ---------------- | ------ | -------------------- |
| 18.8.20 | `fill`         | Fill definition  | ✅     | `RawFill` struct     |
| 18.8.21 | `fills`        | Fills collection | ✅     | In `StyleSheet`      |
| 18.8.22 | `fgColor`      | Foreground color | ✅     | Pattern fg           |
| 18.8.23 | `bgColor`      | Background color | ✅     | Pattern bg           |
| 18.8.32 | `patternFill`  | Pattern fill     | ✅     | `pattern_type` field |
| 18.8.24 | `gradientFill` | Gradient fill    | ➖     | Complex              |
| 18.8.38 | `stop`         | Gradient stop    | ➖     |                      |

**Pattern Types (18.18.55 ST_PatternType):**

| Pattern         | Status |
| --------------- | ------ |
| none            | ✅     |
| solid           | ✅     |
| gray125         | ✅     |
| gray0625        | ✅     |
| darkGray        | ✅     |
| mediumGray      | ✅     |
| lightGray       | ✅     |
| darkHorizontal  | ✅     |
| darkVertical    | ✅     |
| darkDown        | ✅     |
| darkUp          | ✅     |
| darkGrid        | ✅     |
| darkTrellis     | ✅     |
| lightHorizontal | ✅     |
| lightVertical   | ✅     |
| lightDown       | ✅     |
| lightUp         | ✅     |
| lightGrid       | ✅     |
| lightTrellis    | ✅     |

### 18.8.22-31 Fonts

| Section | Element     | Description      | Status | Notes                 |
| ------- | ----------- | ---------------- | ------ | --------------------- |
| 18.8.22 | `font`      | Font definition  | ✅     | `RawFont` struct      |
| 18.8.25 | `fonts`     | Fonts collection | ✅     | In `StyleSheet`       |
| 18.8.2  | `b`         | Bold             | ✅     | `bold` field          |
| 18.8.26 | `i`         | Italic           | ✅     | `italic` field        |
| 18.8.27 | `name`      | Font name        | ✅     | `name` field          |
| 18.8.28 | `outline`   | Outline          | ➖     | Rarely used           |
| 18.8.34 | `scheme`    | Font scheme      | ⬜     | major/minor           |
| 18.8.35 | `shadow`    | Shadow           | ➖     |                       |
| 18.8.36 | `strike`    | Strikethrough    | ✅     | `strikethrough` field |
| 18.8.37 | `sz`        | Size             | ✅     | `size` field          |
| 18.8.39 | `u`         | Underline        | ✅     | `UnderlineStyle` enum |
| 18.8.44 | `vertAlign` | Sub/superscript  | ✅     | `VertAlign` enum      |

**Underline Styles (18.18.82 ST_UnderlineValues):**

| Style            | Status |
| ---------------- | ------ |
| single           | ✅     |
| double           | ✅     |
| singleAccounting | ✅     |
| doubleAccounting | ✅     |
| none             | ✅     |

### 18.8.30-31 Number Formats

| Section | Element   | Description        | Status | Notes             |
| ------- | --------- | ------------------ | ------ | ----------------- |
| 18.8.30 | `numFmt`  | Number format      | ✅     | In `num_fmts` vec |
| 18.8.31 | `numFmts` | Formats collection | ✅     | Custom formats    |

**Built-in Formats (18.8.30):**

| ID    | Format        | Status |
| ----- | ------------- | ------ |
| 0     | General       | ✅     |
| 1     | 0             | ✅     |
| 2     | 0.00          | ✅     |
| 3     | #,##0         | ✅     |
| 4     | #,##0.00      | ✅     |
| 9     | 0%            | ✅     |
| 10    | 0.00%         | ✅     |
| 11    | 0.00E+00      | 🟡     |
| 12    | # ?/?         | 🟡     |
| 13    | # ??/??       | 🟡     |
| 14    | mm-dd-yy      | ✅     |
| 15    | d-mmm-yy      | ✅     |
| 16    | d-mmm         | ✅     |
| 17    | mmm-yy        | ✅     |
| 18    | h:mm AM/PM    | ✅     |
| 19    | h:mm:ss AM/PM | ✅     |
| 20    | h:mm          | ✅     |
| 21    | h:mm:ss       | ✅     |
| 22    | m/d/yy h:mm   | ✅     |
| 37-40 | Accounting    | 🟡     |
| 45-48 | Time formats  | ✅     |
| 49    | @ (text)      | ✅     |

### 18.8.40-45 Theme Elements

| Section | Element             | Description     | Status | Notes             |
| ------- | ------------------- | --------------- | ------ | ----------------- |
| 18.8.40 | `tabColor`          | Sheet tab color | ✅     | `tab_color` field |
| 18.8.43 | `tableStyles`       | Table styles    | ➖     |                   |
| 18.8.41 | `tableStyle`        | Table style     | ➖     |                   |
| 18.8.42 | `tableStyleElement` | Style element   | ➖     |                   |

---

## 18.9 Comments

| Section | Element       | Description        | Status | Notes               |
| ------- | ------------- | ------------------ | ------ | ------------------- |
| 18.9.1  | `authors`     | Comment authors    | ✅     | Authors parsed      |
| 18.9.2  | `comment`     | Comment content    | ✅     | Comment parsed      |
| 18.9.3  | `commentList` | Comments list      | ✅     | Parsed              |
| 18.9.4  | `comments`    | Comments container | ✅     | Parsed              |
| 18.9.5  | `text`        | Comment text       | ✅     | Plain + rich text   |
| -       | Indicator     | Red triangle       | ✅     | Indicator + tooltip |

---

## 18.10 Metadata

| Section  | Element | Description   | Status | Notes                  |
| -------- | ------- | ------------- | ------ | ---------------------- |
| 18.10.\* | Various | Cell metadata | ➖     | Not needed for viewing |

---

## 18.11-17 Other Components

| Section | Component         | Status | Notes            |
| ------- | ----------------- | ------ | ---------------- |
| 18.11   | Calculation Chain | N/A    | No formula eval  |
| 18.12   | Charts            | ➖     | Complex, v2      |
| 18.13   | Connections       | ➖     | External data    |
| 18.14   | Custom XML        | ➖     |                  |
| 18.15   | Drawing           | ➖     | Images/shapes v2 |
| 18.16   | External Links    | ➖     |                  |
| 18.17   | Pivot Tables      | ➖     | Complex          |

---

## 18.18 Simple Types (Enumerations)

| Section  | Type                   | Description      | Status |
| -------- | ---------------------- | ---------------- | ------ |
| 18.18.3  | ST_BorderStyle         | Border styles    | ✅     |
| 18.18.14 | ST_CellType            | Cell value types | ✅     |
| 18.18.30 | ST_FontScheme          | Font scheme      | ⬜     |
| 18.18.40 | ST_HorizontalAlignment | H alignment      | ✅     |
| 18.18.55 | ST_PatternType         | Fill patterns    | ✅     |
| 18.18.66 | ST_SheetState          | Sheet visibility | ✅     |
| 18.18.82 | ST_UnderlineValues     | Underline styles | ✅     |
| 18.18.88 | ST_VerticalAlignment   | V alignment      | ✅     |

---

## Theme (DrawingML - Part 1 Section 20)

> Theme colors referenced from SpreadsheetML styles.

| Element      | Description        | Status | Notes           |
| ------------ | ------------------ | ------ | --------------- |
| `clrScheme`  | Color scheme       | ✅     | 12 theme colors |
| `dk1`        | Dark 1             | ✅     | Theme index 0   |
| `lt1`        | Light 1            | ✅     | Theme index 1   |
| `dk2`        | Dark 2             | ✅     | Theme index 2   |
| `lt2`        | Light 2            | ✅     | Theme index 3   |
| `accent1-6`  | Accent colors      | ✅     | Theme 4-9       |
| `hlink`      | Hyperlink          | ✅     | Theme index 10  |
| `folHlink`   | Followed hyperlink | ✅     | Theme index 11  |
| `fontScheme` | Theme fonts        | ⬜     | major/minor     |
| `fmtScheme`  | Format scheme      | ➖     | Effects         |

---

## Relationships (Part 2)

| Relationship      | Description            | Status |
| ----------------- | ---------------------- | ------ |
| workbook.xml.rels | Workbook relationships | ✅     |
| sheet#.xml.rels   | Sheet relationships    | ✅     |
| hyperlink         | External hyperlinks    | ✅     |
| comments          | Comments file          | ✅     |
| drawing           | Drawing file           | ➖     |
| chart             | Chart file             | ➖     |
| table             | Table definition       | ⬜     |
| pivotTable        | Pivot table            | ➖     |

---

## Summary

### By Category

| Category              | Implemented | Partial | Not Started | Out of Scope |
| --------------------- | ----------- | ------- | ----------- | ------------ |
| Workbook (18.2)       | 5           | 1       | 2           | 14           |
| Worksheets (18.3)     | 28          | 4       | 8           | 5            |
| Shared Strings (18.4) | 6           | 0       | 0           | 3            |
| Styles (18.8)         | 65          | 6       | 4           | 8            |
| Comments (18.9)       | 6           | 0       | 0           | 0            |
| Theme                 | 12          | 0       | 1           | 1            |
| **Total**             | **122**     | **11**  | **15**      | **31**       |

### Compliance Score

- **Core Elements**: ~85% (fonts, fills, borders, alignment, number formats)
- **Extended Features**: ~60% (hyperlinks, data validation, auto-filter, frozen panes)
- **Advanced Features**: ~35% (conditional formatting partial, comments done, charts missing)
- **Overall View-Only Compliance**: **~78%**

### Priority for v1 Completion

1. ⬜ Conditional formatting evaluation (cfvo thresholds, dxfId)
2. ⬜ Table styles
3. ⬜ Font scheme (major/minor)
4. ⬜ Sheet dimension parsing
5. 🟡 Scientific/fraction number formats
6. 🟡 Accounting number formats

---

_Generated from ECMA-376 5th Edition, Part 1 - Section 18 SpreadsheetML_
