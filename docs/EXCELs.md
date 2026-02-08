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
4. Expand `Print_Area` automatically when needed.

Constraints:
- Do not mix `{{#...}}` and `{{##...}}` in the same template.
- Do not define duplicate blocks for the same section.

## 5. Page break (`Print_Area`) update

- Read `_xlnm.Print_Area` from `workbook.xml`.
- Use first-page range height as the baseline.
- Compute required page count.
- Rebuild ranges in the form `A1:Q40,A41:Q80,...`.
- Current logic is row-count based, not physical height based.

Implementation points:
- `parseFirstPrintArea(...)`
- `buildPagedPrintAreas(...)`
- `updatePrintAreaForSheet(...)`

### 5.1 Current limitations

The current page calculation does not include:

- Actual row height (`<row ht="...">`)
- Auto-height expansion from text wrapping
- Height occupied by drawing objects (images/shapes)

As a result, templates with long text, inserted images, or heavy wrapping may produce page breaks that differ from the final printed/PDF layout.

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
- Final rendered row height may be decided by Excel/LibreOffice at render time.
- For better page-break accuracy in the future, `wrapText` evaluation plus height estimation is required.

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
- Calculation is row-count based and does not include wrap/image height
- Inspect `wrapText` in `styles.xml` and cell style indices in `sheetN.xml`

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
