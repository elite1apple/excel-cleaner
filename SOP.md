# Standard Operating Procedure (SOP): Excel Cleaner for Window Shade Automation

## Purpose
This tool converts raw Excel files containing window shade specifications into a standardized, cleaned format ready for automation processing.

## Getting Started: Generating a Template
If you are starting from scratch and need to input data, the easiest way to ensure your format is correct is to use the built-in template generator:
1. Open your terminal or command prompt.
2. Navigate to the folder: `cd c:\Users\mdash\Downloads\Excel-clean`
3. Run: `python create_standard_template.py`
4. This will create a `STANDARD_TEMPLATE.xlsx` file with all the correct headers, example data, and an instructions tab.

---

## Option 1: Using the Web Application (Recommended & Easiest)

### Step 1: Start the Web Application
1. Open your terminal or command prompt.
2. Navigate to the `Excel-clean` folder.
3. Run the application by typing:
   ```bash
   python app.py
   ```
4. Open your web browser and go to `http://localhost:5000`.

### Step 2: Upload Your Excel File
Ensure your raw Excel file meets the required format:
- **Headers**: Your column names (Tag/Unit, Width, Height, Fabric, Control, etc.) must be exactly on **Row 9**.
- **Data Rows**: Start on Row 10. Blank `Tag` cells will automatically inherit the last specified Tag.
- **Metadata (Optional)**: Rows 1-7 can contain Project Name, Date, and Color Codes (e.g., `Bed = YUNOWH`). Cell `I7` usually contains your general D value (e.g., `D = 1/2`).

### Step 3: Process & Download
1. On the web page, drag and drop your `.xlsx` file.
2. The tool will scan for non-numeric tags and ask how you want to handle them (skip, keep, extract).
3. If it cannot find color codes in the Excel file, you can type them directly into the web interface. 
4. Click to process the file and your cleaned output (e.g., `filename-cleaned.xlsx`) will download instantly.

---

## Option 2: Using the Command Line Script

If you prefer processing files via the command line quickly without a user interface:

### Step 1: Run the Script
1. Open your terminal or command prompt.
2. Navigate to the `Excel-clean` folder.
3. Run the cleaning script, providing the path to your raw file:
   ```bash
   python clean_excel.py "path/to/your/raw_file.xlsx"
   ```

### Step 2: Advanced Usage (Optional)
You can directly pass custom output names and color codes via the command line:
```bash
# Order: script.py [input_file] [output_file] [bed_color] [liv_color]
python clean_excel.py "raw_data.xlsx" "cleaned_data.xlsx" "YUNOWH" "PWS3WHIT"
```

### Step 3: Retrieve the Output
The script will output a completed `.xlsx` file in your directory containing the normalized `Width`, `Height`, `Color Number`, `Room` combinations, and `SpecialInstructions`, perfectly formatted for your automation software.

---

## Important Rules & Troubleshooting

- **Bottom Notes Are Ignored**: The tool will automatically skip rows at the bottom that contain "Total", "Notes", "Punch Reverse", etc.
- **Reverse Rolls**: Ensure the **Roll** column contains the exact text `Rev` if you want it flagged as a reverse roll.
- **Deduction Codes**: 
  - `Dl` = Left deduction
  - `Dr` = Right deduction
  - `D` = Both sides
- **Missing Colors**: If you see `BED_COLOR_NEEDED` in your output Excel file, it means you forgot to provide a bed color. You can just do a "Find and Replace" in the final Excel file, or provide the color in the web app UI beforehand.
