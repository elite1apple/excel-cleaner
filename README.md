# Excel Cleaner for Window Shade Automation

A web application and Python script that cleans raw Excel files containing window shade specifications into a standardized format for automation processing.

## 🌐 Web Application (NEW!)

**The easiest way to use this tool!** Upload your Excel files through a web browser and download cleaned results instantly.

### Quick Start - Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the web app
python app.py

# Visit http://localhost:5000
```

### Features

- **Drag & Drop Upload** - Easy file upload interface
- **Auto-Detection** - Automatically extracts color codes from Excel headers
- **Optional Overrides** - Manual inputs for edge cases or unrecognized fabrics
- **Instant Download** - Get your cleaned file in seconds
- **Premium Design** - Modern, responsive interface

### Deployment

Deploy to free hosting platforms like Render, Railway, or PythonAnywhere.

📚 **Full deployment guide:** [DEPLOYMENT.md](DEPLOYMENT.md)

**Key Benefit:** When you update `clean_excel.py` with new logic, just push to GitHub and your web app automatically updates!

---

## 🐍 Command Line Usage

This script can also be used directly from the command line or as a Python module.

## Features

- **Automatic Data Cleaning**: Removes header rows, notes, and invalid data
- **Unicode Fraction Support**: Handles fraction characters in width/height values (¾, ¼, ⅛, ½, etc.)
- **Yellow Highlighting**: Automatically highlights cells with width > 144 inches in yellow
- **Room Identification**: Combines Tag and Fabric columns into standardized Room format (e.g., "120-Bed")
- **Deduction Parsing**: Handles various deduction patterns:
  - `Dl` = Left side deduction
  - `Dr` = Right side deduction  
  - `D` = Both sides deduction
  - Supports fractional values (1/2, 1/4, etc.) and decimals
- **Numeric Tag Validation**: Automatically identifies rows with non-numeric tags
- **Skipped Sheet**: Moves rows with alphabetic/non-numeric tags to a "Skipped" sheet for review
- **Color Number Mapping**: Maps fabric types (Bed/Liv) to color numbers
- **Control Length Formatting**: Converts chain values to standardized format (e.g., `72"LOOP`) - set for all rows with Chain values
- **ReverseRoll Handling**: Sets "Yes" if Roll="Rev" is present, otherwise "No" (not None/NaN)
- **Special Instructions Generation**: Automatically generates instructions based on deductions and reverse roll settings

## Usage

### Basic Usage

```bash
python clean_excel.py <input_file> [output_file] [bed_color] [liv_color]
```

### Parameters

- `input_file` (required): Path to the raw Excel file to clean
- `output_file` (optional): Path for the cleaned output file. If not provided, auto-generates as `<input_file>-cleaned.xlsx`
- `bed_color` (optional): Color number for Bed fabric (e.g., "YUNOWH")
- `liv_color` (optional): Color number for Liv fabric (e.g., "PWS3WHIT")

### Examples

```bash
# Clean with color numbers
python clean_excel.py "Symington Mid Rise L1 to L3 .xlsx" "cleaned.xlsx" "YUNOWH" "PWS3WHIT"

# Auto-generate output filename
python clean_excel.py "raw_data.xlsx" "" "BED123" "LIV456"

# Use as Python module
from clean_excel import clean_excel_file
df = clean_excel_file("input.xlsx", "output.xlsx", "BED_COLOR", "LIV_COLOR")
```

## Input File Format

The script expects Excel files with the following structure:

- **Header Row**: Data headers are typically at row 9 (0-indexed: 8)
- **Columns**:
  - `Tag`: Room/tag number (may have NaN values that belong to previous tag)
  - `Fabric`: Fabric type ("Bed" or "Liv")
  - `Width`: Window width
  - `Height`: Window height
  - `Control`: Control side ("L" or "R")
  - `Deducts`: Deduction code ("Dl", "Dr", or "D")
  - `Unnamed: 12`: Numeric deduction value
  - `Roll`: Roll type ("Rev" for reverse roll)
  - `Chain`: Chain length value

## Output Format

The cleaned file contains these columns:

- `Color Number`: Color code based on fabric type
- `Width`: Window width (float)
- `Height`: Window height (integer)
- `ExtraFabricDeduction`: Deduction value (float, if applicable)
- `ReverseRoll`: "Yes" if Roll="Rev" is present in raw file, otherwise "No"
- `ControlSide`: "LEFT" or "RIGHT"
- `ControlLength`: Formatted chain length (e.g., `72"LOOP`) - set for all rows that have Chain values in raw file
- `Room`: Combined Tag-Fabric format (e.g., "120-Bed")
- `SpecialInstructions`: Generated instructions text

## Raw File Format

See `RAW_FILE_FORMAT.md` for detailed requirements on the expected format of raw Excel files.

**Key Requirements:**
- Headers must be at row 9 (0-indexed: row 8)
- Required columns: `Tag`, `Width`, `Height`, `Fabric`, `Control`
- Optional but important: `Roll`, `Chain `, `Deducts `, `Unnamed: 12`
- Column names must match exactly (including trailing spaces)

## Customization

### Adjusting Header Row

If your files have headers at a different row, modify the `header_row` parameter:

```python
df = clean_excel_file("input.xlsx", header_row=7)  # 0-indexed
```

### Changing Note Detection Keywords

Customize which rows are considered "notes" to exclude:

```python
df = clean_excel_file(
    "input.xlsx",
    notes_start_keywords=["Total", "Notes", "Summary"]
)
```

### Color Numbers

Color numbers can be provided in **three ways** (in order of priority):

1. **Command line arguments** (highest priority):
   ```bash
   python clean_excel.py input.xlsx output.xlsx BED123 LIV456
   ```

2. **In the Excel file header** (rows 0-7):
   The script automatically extracts color codes from the header section. Supported formats:
   - `Bed = COLOR123` or `Bedroom = COLOR123` or `Bed/Bedroom = COLOR123`
   - `Living = COLOR456` or `Liv = COLOR456` or `Living/Liv = COLOR456`
   - `Studio = COLOR789`
   - `Den = COLOR012`
   - `Kitchen = COLOR345`
   - `Bath = COLOR678` or `Bathroom = COLOR678`
   
   Example header rows:
   ```
   Row 1: Bed/Bedroom = BED123
   Row 2: Living/Liv = LIV456
   Row 3: Studio = STUDIO789
   Row 4: Den = DEN012
   Row 5: Kitchen = KITCHEN345
   Row 6: Bath = BATH678
   ```

3. **Manual update** (after cleaning):
   - Update color numbers directly in the cleaned Excel file
   - Or use `fabric_colors` parameter programmatically:
     ```python
     fabric_colors = {'Studio': 'STUDIO123', 'Kitchen': 'KITCHEN456', 'Den': 'DEN789'}
     ```

**Other Fabric Types:**
The script automatically handles additional fabric types (Kitchen, Studio, Den, etc.):
- Creates placeholder: `{FABRIC}_COLOR_NEEDED` (e.g., `STUDIO_COLOR_NEEDED`)
- You can provide custom colors via `fabric_colors` parameter or update manually

## Deduction Value Handling

The script handles deduction values in multiple formats:

### Format Options:

**In Deducts Column:**
- `Dl=1/2` or `Dl-1/2` → Left side, 1/2 inch (stores 0.5)
- `Dr=1/2` or `Dr-1/2` → Right side, 1/2 inch (stores 0.5)
- `D=1/2` or `D-1/2` → Both sides, each gets 1/2 (stores 1.0: 1/2 + 1/2)
- `D=1` → Both sides, total is 1 (stores 1.0 directly)
- `Dl`, `Dr`, `D` (no value) → Uses default from notes section

**In Notes Section:**
- Standard format: `Dl 1/2 Left, Dr = 1/2 Inch Right, D 1/2 Inch both sides`
- Or: `Dl=1/2, Dr=1/2, D=1/2`

### Important Logic:

**When `D=1/2` is specified (per-side value):**
- `Dl` = 1/2 (left side only, stores 0.5)
- `Dr` = 1/2 (right side only, stores 0.5)
- `D` = 1 (both sides: 1/2 + 1/2 = 1, stores 1.0)

**Examples:**
- `Dl=1/2` → Left side, 1/2 inch deduction (stores 0.5)
- `Dr=1/2` → Right side, 1/2 inch deduction (stores 0.5)
- `D=1/2` → Both sides, each gets 1/2, total = 1 (stores 1.0)
- `D=1` → Both sides, total is 1 (stores 1.0)
- `Dl=1/4, Dr=1/4` → Left and right, 1/4 inch each (stores 0.25 each)

## Troubleshooting

### Missing Color Numbers

If you see `BED_COLOR_NEEDED` or `LIV_COLOR_NEEDED` in the output, provide color numbers as parameters.

### Incorrect Row Counts

If the cleaned file has fewer rows than expected:
- Check if note rows are being incorrectly filtered
- Verify header_row is correct
- Check for rows with missing Width/Height values

### Deduction Values Not Parsing

If deductions aren't being applied:
- Verify the `Deducts` column contains "Dl", "Dr", or "D"
- Check that `Unnamed: 12` column has numeric values
- Ensure Control side matches deduction side (L matches Dl, R matches Dr)

## Requirements

- Python 3.7+
- pandas
- openpyxl

Install dependencies:
```bash
pip install pandas openpyxl
```

