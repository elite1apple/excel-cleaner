# Raw File Format Requirements

This document describes the expected format for raw Excel files that will be processed by the cleaning script.

## Header Row Position

The data headers should be located at **row 9** (0-indexed: row 8) in the Excel file. Rows 0-7 typically contain project metadata (project name, date, location, etc.).

## Required Column Headers

The raw Excel file must have the following columns (in any order, but these exact names):

| Column Name | Description | Example Values | Required |
|------------|-------------|----------------|----------|
| `Tag/Unit` | Room/tag number | 120, 121, 122 | Yes |
| `Q` | Quantity (usually 1) | 1 | No |
| `Product` | Product type | "Manual Shade" | No |
| `Roll` | Roll type | "Rev" (for reverse roll) | No |
| `Width` | Window width in inches | 32, 62.25, 42.5 | **Yes** |
| `Height` | Window height in inches | 86, 122, 83 | **Yes** |
| `Chain ` | Chain length value | 72, 48, 60, 62 | No (but needed for ControlLength) |
| `Fabric` | Fabric type | "Bed", "Liv" | **Yes** |
| `Control` | Control side | "L" (Left), "R" (Right) | **Yes** |
| `Deducts ` | Deduction code | "Dl" (Left), "Dr" (Right), "D" (Both) | No |

**Note**: Column names must match exactly, including trailing spaces (e.g., `Chain ` has a trailing space, `Deducts ` has a trailing space).

## Data Format Rules

### Tag Column
- Contains room/tag numbers (e.g., 120, 121, 223)
- **NaN values are allowed** - they will be forward-filled from the previous non-NaN Tag value
- Example: If Tag 120 is followed by NaN rows, those rows belong to room 120

### Width and Height
- Must be numeric values (integers or decimals)
- Can have trailing periods (e.g., "24.25.") - script will clean these
- Required for all data rows

### Fabric Column
- Values: "Bed" or "Liv" (case-insensitive)
- Used to determine color number mapping
- Required for all data rows

### Control Column
- Values: "L" (Left) or "R" (Right)
- Case-insensitive (will be converted to uppercase)
- Required for all data rows

### Roll Column
- If value is "Rev" (case-insensitive), ReverseRoll will be set to "Yes"
- Otherwise, ReverseRoll will be set to "No"
- Can be empty/NaN

### Chain Column
- Numeric value representing chain length (e.g., 72, 48, 60)
- Will be formatted as ControlLength: `{value}"LOOP` (e.g., `72"LOOP`)
- Can be empty/NaN (ControlLength will be None)

### Deducts Column
- Values: "Dl" (Left deduction), "Dr" (Right deduction), or "D" (Both sides)
- Can also contain explicit values: "Dl=1/2", "D=1", etc.
- If just code is used ("D"), value is taken from cell I7 or "D=X" note
- Can be empty/NaN (no deduction)

### Unnamed: 12 Column
- **Removed** - Deduction values are now handled within the Deducts column or via D value lookup.

## Notes Section

The script automatically removes rows that contain note keywords:
- "Total"
- "all Finshed" (or variations)
- "Punch Reverse"
- "Deducts"

These rows typically appear at the bottom of the file and should not contain valid data rows.

## Example Raw File Structure

```
Row 0-7: Project metadata (Project name, Date, Location, etc.)
Row 8:   Headers (Tag, Drawing Ref., Q, Product, Roll, Width, Height, Chain , Fabric, Mount, Control, Deducts , Unnamed: 12)
Row 9+:  Data rows
...
Last rows: Notes section (will be removed)
```

## Sample Data Row

| Tag/Unit | Q | Product | Roll | Width | Height | Chain | Fabric | Control | Deducts |
|-----|---|---------|------|-------|--------|-------|--------|---------|----------|
| 120 | 1 | Manual Shade | Rev | 32 | 86 | 72 | Bed | L | - |
| - | 1 | Manual Shade | Rev | 62.25 | 86 | 72 | Bed | R | - |
| - | 1 | Manual Shade | Rev | 42.25 | 122 | 72 | Liv | L | Dl |
| - | 1 | Manual Shade | Rev | 52 | 122 | 72 | Liv | R | Dr |

**Note**: The second row has Tag = "-" (NaN), which means it belongs to Tag 120 from the previous row.

## Validation Checklist

Before processing, ensure your raw file has:

- [ ] Headers at row 9 (0-indexed: row 8)
- [ ] Required columns: Tag/Unit, Width, Height, Fabric, Control
- [ ] All data rows have valid Width, Height, Fabric, and Control values
- [ ] Tag values are numeric (or NaN for continuation rows)
- [ ] Fabric values are "Bed" or "Liv"
- [ ] Control values are "L" or "R"
- [ ] Notes section is at the bottom (will be auto-removed)

## Common Issues

1. **Missing headers**: If headers are at a different row, use `header_row` parameter
2. **Column name mismatches**: Column names must match exactly (including spaces)
3. **Invalid Tag values**: Tag must be numeric or NaN (not text)
4. **Missing required columns**: Script will fail if Width, Height, Fabric, or Control are missing
5. **Deduction mismatch**: (REMOVED - Deductions can now differ from Control side)



