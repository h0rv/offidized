# BUG-012: Locale-Specific Formats Missing

**Priority**: MEDIUM
**Status**: Open
**Component**: number format rendering

## Problem

Date, time, and currency formats are hardcoded in English. Excel respects workbook and system locale for formatting, but offidized-xlview always uses English month names, US date order, and US currency symbols.

## Current Behavior

All dates show English month names regardless of workbook locale.

## Expected Behavior

1. Detect workbook locale from `calcPr` or content types
2. Format dates/times according to locale
3. Use locale-appropriate currency symbols

## Where Data Should Come From

Locale hints can come from several places:

```xml
<!-- xl/workbook.xml -->
<workbook>
  <calcPr calcId="191029"/>
  <!-- No explicit locale usually means system locale -->
</workbook>

<!-- Number format codes often contain locale hints -->
<numFmt numFmtId="164" formatCode="[$-409]d-mmm-yy"/>
<!-- $-409 = US English locale -->

<numFmt numFmtId="165" formatCode="[$-407]d. mmm yy"/>
<!-- $-407 = German locale -->
```

The `[$-XXXX]` prefix in format codes indicates locale:

- `409` = US English
- `407` = German
- `40C` = French
- `411` = Japanese

## Impact

- Non-English users see English month names
- Date formats may be confusing (MM/DD vs DD/MM)
- Currency symbols may be wrong for non-US files

## Proposed Fix

1. Parse locale hints from format codes `[$-XXXX]`
2. Build a locale mapping for common languages
3. Use locale-appropriate month names, date separators, etc.

```rust
struct LocaleInfo {
    month_names: [&str; 12],
    month_abbrevs: [&str; 12],
    date_separator: &str,
    time_separator: &str,
    currency_symbol: &str,
}

fn get_locale(code: u32) -> LocaleInfo {
    match code {
        0x409 => LOCALE_EN_US,
        0x407 => LOCALE_DE_DE,
        0x40C => LOCALE_FR_FR,
        // ...
    }
}
```

## References

- ECMA-376 Part 1, Section 18.8.31 (numFmt)
- Windows LCID codes: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a
