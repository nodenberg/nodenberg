# EXCELs.md

Implementation notes for Excel (XLSX) internals used in Nodenberg (`03_docker-version`).

## 1. XLSX structure used by this implementation

An XLSX file is a ZIP archive. The implementation mainly edits the following files directly:

- `xl/sharedStrings.xml`
  - String storage (cells reference strings by index)
- `xl/worksheets/sheetN.xml`
  - Row/cell layout and cell references (for example: `<c r="A1" ...><v>index</v></c>`)
- `xl/workbook.xml`
  - Sheet definitions and named ranges (for example: `_xlnm.Print_Area`)
- `xl/_rels/workbook.xml.rels`
  - Mapping between workbook entries and worksheet files

Related implementation files:
- `src/lib/placeholderReplacer.ts`
- `src/lib/excelGenerator.ts`
- `src/lib/sheetSelector.ts`

## 2. Cell string representation

### 2.1 Plain shared string

```xml
<si><t>{{company_name}}</t></si>
```

### 2.2 Rich text shared string

```xml
<si>
  <r><rPr>...</rPr><t>{{</t></r>
  <r><rPr>...</rPr><t>company_name</t></r>
  <r><rPr>...</rPr><t>}}</t></r>
</si>
```

Important:
- A visually single placeholder such as `{{company_name}}` may be split into multiple runs.
- Raw string replacement alone can fail in these cases.

## 3. Meaning of `si` / `r` / `rPr`

- `si`: one shared string entry (referenced by cell index)
- `r`: one run inside an `si` (substring unit)
- `rPr`: run-level text style (font, size, bold, color, etc.)

Cell-wide style is managed separately:
- Cell `s` attribute in `sheetN.xml` -> style definition in `styles.xml`

## 4. Placeholder replacement logic (current)

### 4.1 Standard placeholder `{{key}}`

Target:
- `si` entries in `sharedStrings.xml`

Method:
1. Concatenate all `t` nodes in an `si` to build a logical string.
2. Replace `{{key}}`.
3. Rebuild the `si` after replacement.
   - Collapse to a single run using the first run's `rPr` as the style base.

Implementation points:
- `replacePrimitivePlaceholdersWithFirstRunStyle(...)`
- `extractSiText(...)`
- `extractFirstRunProperties(...)`

Notes:
- String indices are generally kept as-is (no `<v>` index remapping for this path).
- Replaced text is normalized to the first run style.

### 4.2 Legacy array placeholder `{{#array.field}}`

Target:
- `sheet1.xml` (legacy compatibility)

Method:
1. Find the shared string index for the placeholder.
2. Insert rows when needed.
3. Add one shared string per value.
4. Replace cell `<v>oldIndex</v>` with `<v>newIndex</v>`.

### 4.3 Section placeholder `{{##section.table.cell}}`

Target:
- The sheet where the placeholder exists (resolved dynamically)

Method:
1. Detect contiguous row blocks per `section.table`.
2. Duplicate the block based on record count.
3. Add shared strings and update indices.
4. Record each duplicated record block as a keep-together unit for pagination.
5. Extend `Print_Area` (single range) and insert manual row breaks (`<rowBreaks>`) from the generated layout.

Constraints:
- Do not mix `{{#...}}` and `{{##...}}` in the same template.
- Do not define duplicate blocks for the same section.
- Every row that contains a `{{##section.table.*}}` placeholder becomes part of
  that record block and is duplicated per record. Keep table header rows free
  of section placeholders, otherwise the header row is repeated for every
  record. Static cells inside block rows are duplicated by design (use them
  for per-record labels/frames).

## 5. Page break (`Print_Area` + `rowBreaks`) update

Page layout is expressed the same way Excel does it for manual printing:
a **single** `Print_Area` range plus **manual row breaks** in the worksheet XML.

- Read `_xlnm.Print_Area` from `workbook.xml` and use the first range as the base.
- Extend the end row of that single range by the number of inserted rows
  (`startRow:endRow + insertedRows`). The range is never split into multiple
  comma-separated areas — comma-separated areas are independent print areas in
  OOXML and each one always starts a new page, which is not what manual page
  breaks look like.
- Read page settings from `sheetN.xml` such as paper size, orientation,
  margins, and fit-to-width settings. The printable height is
  `pageHeight - topMargin - bottomMargin`; header/footer margins are *inside*
  the top/bottom margins in OOXML, so they do not reduce the content area.
- Estimate row heights (after all placeholder replacement, so wrapped text is
  measured with real values) and walk the rows, emitting a manual break
  `<brk id="N" max="16383" man="1"/>` ("break after row N") whenever the next
  row would exceed the page capacity.
- Keep-together units (record blocks, image merge spans) are never split: if a
  unit would cross the boundary, the break is placed before the unit instead.
  No blank padding rows are inserted; the break itself starts the new page.
- If a section spans multiple pages, the next section starts on a new page.
- Manual breaks already present in the template are preserved and respected as
  forced page boundaries.
- When manual breaks are emitted, `fitToHeight` is forced to `0` because
  fit-to-height scaling and manual breaks are mutually exclusive in Excel.

The `<rowBreaks>` element is placed according to the `CT_Worksheet` element
order (after `pageSetup` / `headerFooter`):

```xml
<rowBreaks count="2" manualBreakCount="2">
  <brk id="48" max="16383" man="1"/>
  <brk id="96" max="16383" man="1"/>
</rowBreaks>
```

Implementation points:
- `applyPaginationToSheet(...)`
- `computeManualRowBreaks(...)`
- `upsertRowBreaks(...)`
- `parsePageLayoutInfo(...)`

### 5.1 Current limitations

The current page calculation still has limits. It does not fully reproduce:

- Auto-height expansion that Excel/LibreOffice decides at render time without explicit `wrapText`
- Height occupied by drawing objects (images/shapes)
- Font rendering differences between Excel and LibreOffice

Current behavior:
- Explicit row heights (`<row ht="...">`) are used.
- `wrapText="1"` styles are considered when estimating row height.
- Cells without wrapping use their row height as-is.

As a result, templates with long text, inserted images, or renderer-specific layout behavior may still produce page breaks that differ slightly from the final printed/PDF layout.

### 5.2 Print layout options

The print layout options must be applied to the generated XLSX before PDF
conversion. PDF generation must not own a separate margin, paper, or scaling
model. The implemented pipeline is:

1. Read the template XLSX.
2. Apply requested `printLayout` settings to worksheet XML.
3. Replace placeholders and expand section tables.
4. Recalculate `Print_Area` and `rowBreaks` from the updated worksheet print
   settings.
5. Return the XLSX as-is for Excel output, or convert that same XLSX to PDF.

This keeps "print from Excel" and "generate PDF" on the same source of truth.
Changing margins only at PDF conversion time is not acceptable because section
pagination has already been calculated against the XLSX print settings.

Request shape:

```json
{
  "options": {
    "printLayout": {
      "paperSize": "A4",
      "orientation": "portrait",
      "marginPreset": "narrow",
      "margins": {
        "top": 0,
        "bottom": 0
      },
      "fit": {
        "width": 1,
        "height": 0
      },
      "recalculatePagination": true
    }
  }
}
```

Margin behavior:

- If `printLayout` is omitted, preserve the template settings.
- `marginPreset` selects a complete preset such as `normal`, `narrow`, or
  `wide`.
- `margins` is a partial override object using inch values. Supported keys are
  `left`, `right`, `top`, `bottom`, `header`, and `footer`.
- If both `marginPreset` and `margins` are provided, resolve the preset first
  and then override only the provided margin keys. This supports requests such
  as "narrow, but top/bottom/header/footer are zero".
- `margins: {}` is invalid. At least one field must be supplied.
- Negative values are invalid. Unsupported keys are invalid.

Pagination behavior:

- `recalculatePagination` may be omitted.
- `recalculatePagination: true` is accepted and documents the default behavior.
- `recalculatePagination: false` is rejected. Applying print layout without
  recalculating `Print_Area` and `rowBreaks` can create an inconsistent XLSX/PDF
  layout, so disabling recalculation is not supported.

Suggested presets:

```text
normal: left=0.7,  right=0.7,  top=0.75, bottom=0.75, header=0.3, footer=0.3
narrow: left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3, footer=0.3
wide:   left=1.0,  right=1.0,  top=1.0,  bottom=1.0,  header=0.5, footer=0.5
```

OOXML fields to update:

```xml
<!-- xl/worksheets/sheetN.xml -->
<pageMargins left="0.25" right="0.25" top="0" bottom="0" header="0" footer="0"/>
<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>

<!-- xl/workbook.xml -->
<definedName name="_xlnm.Print_Area" localSheetId="0">'Sheet1'!$A$1:$D$80</definedName>
```

Section pagination must use the updated worksheet XML. The page height
calculation remains:

```text
printableHeightPoints = pageHeightPoints - topMarginInch * 72 - bottomMarginInch * 72
pageCapacityPoints = printableHeightPoints / effectiveScale
```

Header and footer margins are inside the top/bottom margins in OOXML, so they
do not reduce the body height again. They still must be written to the XML so
Excel, LibreOffice, and PDF conversion see the same worksheet settings.

Width behavior is a separate decision from margins:

- `fit.width=1, fit.height=0` lets Excel/LibreOffice scale the print area to
  the printable page width.
- Column widths are not automatically expanded by margin changes.
- If a future feature needs the table itself to fill the wider printable area,
  add an explicit strategy such as
  `widthStrategy: "expandColumnsToPrintableWidth"`. Do not hide that behavior
  inside margin presets.

## 6. Where `wrapText` is stored

Wrapping behavior is mainly determined by this combination:

1. `xl/styles.xml`
- `wrapText="1"` in `alignment` under `cellXfs` / `xf`
- Example: `<xf ...><alignment wrapText="1" .../></xf>`

2. `xl/worksheets/sheetN.xml`
- Each cell `<c ... s="N">` uses style index `N`
- Resolve `s` -> corresponding `xf` in `styles.xml` to determine wrapping

Notes:
- Effective style can also be influenced by defaults (`xfId`) and row/column settings.
- Final rendered row height may still be decided by Excel/LibreOffice at render time.
- The current implementation uses `wrapText` and page settings as an estimate, not a perfect renderer clone.

## 7. Template info APIs

### 7.1 `POST /template/placeholders`

- Input: `templateBase64`, optional `detailed`
- Behavior:
  - Extract `{{...}}` placeholders from `sharedStrings.xml`
  - If `detailed=true`, return `{ placeholder, key, count }`
- Implementation:
  - `PlaceholderReplacer.findPlaceholders()`
  - `PlaceholderReplacer.getPlaceholderInfo()`

### 7.2 `POST /template/info`

- Input: `templateBase64`
- Behavior:
  - Return basic template metadata such as sheet count
- Implementation:
  - `ExcelGenerator.getTemplateInfo()`

### 7.3 `POST /template/sheets`

- Input: `templateBase64`
- Behavior:
  - Return display order, `sheetId`, and sheet names from `workbook.xml`
- Implementation:
  - `ExcelGenerator.getTemplateSheets()`

## 8. Common pitfalls

1. Placeholders are detected but not replaced during generation
- Cause: run splitting (`<r><t>`) or XML fragmentation
- Fix: replace based on the logical `si` string

2. Unexpected text appears in PDF output
- Example: sheet name appears via header token `&A`
- Check `workbook.xml` sheet names and `sheetN.xml` header/footer first

3. Section block detection errors
- `section.table` rows are not contiguous
- Same section appears in multiple separate blocks

4. Page breaks are offset from expectation
- Check template `Print_Area`, `pageSetup`, `pageMargins`, and row heights first
- Inspect `wrapText` in `styles.xml` and cell style indices in `sheetN.xml`
- For section-table templates, confirm that one record block is contiguous and that the record itself is not taller than one page

## 9. Debug commands

```bash
# Inspect shared strings
unzip -p template.xlsx xl/sharedStrings.xml | less

# Inspect worksheet structure
unzip -p template.xlsx xl/worksheets/sheet1.xml | less

# Inspect sheet names and Print_Area
unzip -p template.xlsx xl/workbook.xml | less

# Inspect wrapping styles (wrapText)
unzip -p template.xlsx xl/styles.xml | rg -n "wrapText|<xf|<alignment"

# Inspect cell style indices (s attribute)
unzip -p template.xlsx xl/worksheets/sheet1.xml | rg -n "<c r=| s=\""
```
