# Table with Orientation Changes - Example

This example demonstrates how to insert tables with dynamic page orientation changes in DOCX documents using the `docx-update` library.

## File Location

`examples/example_table_orientation.go`

## What It Demonstrates

This example shows a complete workflow for creating professional reports with wide tables that require landscape orientation:

1. **Portrait Section** - Document introduction
2. **Landscape Section** - Wide employee data table (8 columns)
3. **Portrait Section** - Analysis and summary
4. **A4 Landscape Section** - Quarterly sales report (7 columns)
5. **Portrait Section** - Final conclusion

## Running the Example

```bash
# From the project root directory
go run examples/example_table_orientation.go
```

## Expected Output

```
Opening template: ./templates/docx_template.docx
Adding initial content in portrait...
Inserting section break (portrait → landscape)...
Inserting wide table in landscape section...
Inserting section break (landscape → portrait)...
Adding conclusion in portrait...
Inserting section break (portrait → A4 landscape)...
Returning to portrait orientation...
Saving document to: ./outputs/table_orientation_demo.docx

✓ Document created successfully!

Document Structure:
  1. Portrait section - Introduction
  2. Letter Landscape section - Employee table (8 columns)
  3. Portrait section - Analysis
  4. A4 Landscape section - Quarterly sales table (7 columns)
  5. Portrait section - Conclusion

Output: ./outputs/table_orientation_demo.docx
```

## Key Code Patterns

### 1. Insert Section Break to Switch to Landscape

```go
// Switch from portrait to landscape
err = updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterLandscape(),
})
```

### 2. Insert Wide Table in Landscape Section

```go
err = updater.InsertTable(docx.TableOptions{
    Position: docx.PositionEnd,
    Columns: []docx.ColumnDefinition{
        {Title: "Column1", Alignment: docx.CellAlignLeft},
        {Title: "Column2", Alignment: docx.CellAlignRight},
        // ... more columns
    },
    Rows: [][]string{
        {"Data1", "Data2", ...},
        {"Data3", "Data4", ...},
    },
    HeaderBold:        true,
    HeaderBackground:  "2E75B5",
    AlternateRowColor: "E7E6E6",
    BorderStyle:       docx.BorderSingle,
    TableAlignment:    docx.AlignCenter,
    RepeatHeader:      true,
})
```

### 3. Return to Portrait Orientation

```go
// Switch back from landscape to portrait
err = updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  docx.PageLayoutLetterPortrait(),
})
```

## Available Page Layouts

### US Letter (8.5" × 11")
- `PageLayoutLetterPortrait()` - Portrait with 1" margins
- `PageLayoutLetterLandscape()` - Landscape with 1" margins

### A4 (210mm × 297mm)
- `PageLayoutA4Portrait()` - Portrait with default margins
- `PageLayoutA4Landscape()` - Landscape with default margins

### A3 (297mm × 420mm)
- `PageLayoutA3Portrait()` - Portrait with default margins
- `PageLayoutA3Landscape()` - Landscape with default margins

### Legal (8.5" × 14")
- `PageLayoutLegalPortrait()` - Portrait with default margins

## Section Break Types

- `SectionBreakNextPage` - Start new section on next page (use for orientation changes)
- `SectionBreakContinuous` - Start new section on same page (different margins)
- `SectionBreakEvenPage` - Start new section on next even-numbered page
- `SectionBreakOddPage` - Start new section on next odd-numbered page

## Custom Page Layout

You can also create custom page layouts with specific dimensions and margins:

```go
customLayout := &docx.PageLayoutOptions{
    PageWidth:    docx.PageWidthLetter,  // Width in twips (1/1440 inch)
    PageHeight:   docx.PageHeightLetter,
    Orientation:  docx.OrientationLandscape,
    MarginTop:    docx.MarginNarrow,     // 0.5"
    MarginRight:  docx.MarginNarrow,
    MarginBottom: docx.MarginNarrow,
    MarginLeft:   docx.MarginNarrow,
    MarginHeader: docx.MarginHeaderFooterDefault,
    MarginFooter: docx.MarginHeaderFooterDefault,
    MarginGutter: 0,
}

updater.InsertSectionBreak(docx.BreakOptions{
    Position:    docx.PositionEnd,
    SectionType: docx.SectionBreakNextPage,
    PageLayout:  customLayout,
})
```

## Margin Constants

- `MarginDefault` - 1.0 inch (1440 twips)
- `MarginNarrow` - 0.5 inch (720 twips)
- `MarginWide` - 1.5 inches (2160 twips)
- `MarginHeaderFooterDefault` - 0.5 inch (720 twips)

## Use Cases

This pattern is ideal for:

- **Financial Reports** - Wide data tables with many columns
- **Employee Databases** - Personnel information with multiple fields
- **Project Status Reports** - Gantt charts or resource allocation tables
- **Inventory Management** - Product catalogs with detailed specifications
- **Performance Dashboards** - Quarterly metrics and KPI tables

## Testing the Output

To verify the orientation changes were applied correctly:

```bash
# Count landscape orientation markers (should be 2)
unzip -p outputs/table_orientation_demo.docx word/document.xml | grep -o 'w:orient="landscape"' | wc -l

# Verify the file is a valid DOCX
file outputs/table_orientation_demo.docx
# Output: outputs/table_orientation_demo.docx: Microsoft OOXML
```

## Related Examples

- `example_table.go` - Basic table insertion with styling
- `example_page_layout.go` - Page layout and section break examples
- `example_breaks.go` - Page and section break patterns

## API Reference

For complete API documentation, see:
- `docs/API_DOCUMENTATION.md` - Complete library reference
- `docs/API_REFERENCE.md` - Fiber v3 integration patterns
