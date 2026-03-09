# Known Issues

Open issues where rendering differs from Excel. All are edge cases -- common spreadsheets render correctly.

| ID  | Priority | Description                               |
| --- | -------- | ----------------------------------------- |
| 007 | MEDIUM   | Auto color always resolves to black       |
| 008 | MEDIUM   | Border color default wrong                |
| 009 | MEDIUM   | Gradient fills not parsed                 |
| 010 | MEDIUM   | baseColWidth not parsed                   |
| 011 | MEDIUM   | Color transforms not parsed               |
| 012 | MEDIUM   | Locale-specific number formats missing    |
| 013 | MEDIUM   | Missing built-in numfmt IDs               |
| 015 | MEDIUM   | Conditional formatting thresholds ignored |
| 016 | MEDIUM   | Shrink-to-fit not implemented             |
| 017 | MEDIUM   | Indent uses hardcoded px                  |
| 018 | MEDIUM   | Default row height hardcoded              |

Each issue file contains: problem description, current vs expected behavior, file locations, and ECMA-376 spec references.
