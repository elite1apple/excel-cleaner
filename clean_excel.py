"""
Excel Cleaning Script for Window Shade Automation
Cleans raw Excel files into standardized format for automation processing
"""

import pandas as pd
import re
from pathlib import Path
from typing import Dict, Optional, Tuple
import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter


def parse_fraction_string(value_str: str) -> Optional[float]:
    """
    Parse a string containing numbers with fractions, including Unicode fraction characters.
    
    Handles:
    - Regular decimals: "54", "51.5"
    - Mixed fractions with spaces: "54 3/4", "51 1/4"
    - Unicode fractions: "54¾", "51¼", "80⅛", "54½"
    - Regular fractions: "3/4", "1/4"
    
    Examples:
    - "54¾" → 54.75
    - "51 1/4" → 51.25
    - "80⅛" → 80.125
    - "51" → 51.0
    """
    if not value_str or pd.isna(value_str):
        return None
    
    # Unicode fraction mappings
    unicode_fractions = {
        '¼': 0.25,
        '½': 0.5,
        '¾': 0.75,
        '⅐': 1/7,
        '⅑': 1/9,
        '⅒': 0.1,
        '⅓': 1/3,
        '⅔': 2/3,
        '⅕': 0.2,
        '⅖': 0.4,
        '⅗': 0.6,
        '⅘': 0.8,
        '⅙': 1/6,
        '⅚': 5/6,
        '⅛': 0.125,
        '⅜': 0.375,
        '⅝': 0.625,
        '⅞': 0.875,
    }
    
    value_str = str(value_str).strip()
    
    # Replace Unicode fractions with their decimal equivalents
    for unicode_frac, decimal_val in unicode_fractions.items():
        if unicode_frac in value_str:
            # Extract whole number part if present
            parts = value_str.split(unicode_frac)
            if parts[0].strip():
                try:
                    whole = float(parts[0].strip())
                    return whole + decimal_val
                except (ValueError, TypeError):
                    pass
            else:
                return decimal_val
    
    # Handle regular fractions with spaces (e.g., "54 3/4" or "51 1/4")
    if ' ' in value_str and '/' in value_str:
        parts = value_str.split()
        if len(parts) == 2:
            try:
                whole = float(parts[0])
                frac_parts = parts[1].split('/')
                if len(frac_parts) == 2:
                    numerator = float(frac_parts[0])
                    denominator = float(frac_parts[1])
                    return whole + (numerator / denominator)
            except (ValueError, TypeError, ZeroDivisionError):
                pass
    
    # Handle standalone fractions (e.g., "3/4")
    if '/' in value_str and ' ' not in value_str:
        parts = value_str.split('/')
        if len(parts) == 2:
            try:
                numerator = float(parts[0])
                denominator = float(parts[1])
                return numerator / denominator
            except (ValueError, TypeError, ZeroDivisionError):
                pass
    
    # Try regular float parsing
    try:
        return float(value_str.rstrip('.'))
    except (ValueError, TypeError):
        return None


def extract_deduction_values_from_notes(df_raw: pd.DataFrame, header_row: int = 8, deduction_cell: str = None) -> Dict[str, float]:
    """
    Extract deduction values from notes section.
    
    First tries to read from a specific cell (e.g., "G8" or "I6"), then falls back to searching rows.
    
    Simple logic:
    - Look for "D = 1" or "D=1" in notes (can be fraction like "D=1/2" or decimal "D=1")
    - D value is the TOTAL value for both sides
    - Dl and Dr are automatically set to D/2 (half of D)
    
    Examples:
    - Notes: "D=1" → D=1.0, Dl=0.5, Dr=0.5
    - Notes: "D=1/2" → D=0.5, Dl=0.25, Dr=0.25
    
    Args:
        deduction_cell: Optional cell reference like "G8" or "I6" to read D value from
    
    Returns dict: {'D': value, 'Dl': D/2, 'Dr': D/2, 'DL': D/2, 'DR': D/2}
    """
    result = {}
    
    # First, try to read from specific cell if provided
    if deduction_cell:
        try:
            import openpyxl
            from openpyxl.utils import coordinate_from_string, column_index_from_string
            
            # Parse cell reference (e.g., "G8" -> row=8, col=G)
            col_letter, row_num = coordinate_from_string(deduction_cell.upper())
            col_idx = column_index_from_string(col_letter) - 1  # Convert to 0-based index
            row_idx = row_num - 1  # Convert to 0-based index
            
            # Check if row/col are within bounds
            if 0 <= row_idx < len(df_raw) and 0 <= col_idx < len(df_raw.columns):
                cell_value = df_raw.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_text = str(cell_value)
                    print(f"DEBUG: Reading from cell {deduction_cell}: {cell_text}")
                    
                    # Look for D=value pattern in the cell
                    d_pattern = re.search(r'D\s*[=:]\s*((?:\d+/\d+)|(?:\d*\.?\d+))', cell_text, re.IGNORECASE)
                    if d_pattern:
                        value_str = d_pattern.group(1)
                        print(f"DEBUG: Matched D pattern in cell {deduction_cell}, value_str='{value_str}'")
                        
                        # Parse value
                        if '/' in value_str:
                            value = int(value_str.split('/')[0]) / int(value_str.split('/')[1])
                        else:
                            value = float(value_str)
                        
                        result['D'] = value
                        half_value = value / 2
                        result['Dl'] = half_value
                        result['Dr'] = half_value
                        result['DL'] = half_value
                        result['DR'] = half_value
                        
                        print(f"DEBUG: Extracted from cell {deduction_cell} - D={result['D']}, Dl={result['Dl']}, Dr={result['Dr']}")
                        return result
        except Exception as e:
            print(f"DEBUG: Could not read from cell {deduction_cell}: {e}")
    
    # Fall back to searching rows (check header section first, then bottom)
    print(f"DEBUG extract_deduction_values_from_notes: Searching for D value in rows...")
    
    # Search header section first (rows near header_row)
    for idx in range(max(0, header_row - 5), min(header_row + 5, len(df_raw))):
        row_text = ' '.join([str(x) for x in df_raw.iloc[idx].tolist() if pd.notna(x)])
        
        # Look for D=value pattern
        d_pattern = re.search(r'D\s*[=:]\s*((?:\d+/\d+)|(?:\d*\.?\d+))', row_text, re.IGNORECASE)
        if d_pattern:
            value_str = d_pattern.group(1)
            print(f"DEBUG: Found D pattern in header row {idx}: {row_text[:200]}")
            print(f"DEBUG: Matched D pattern, value_str='{value_str}'")
            
            # Parse value
            if '/' in value_str:
                parts = value_str.split('/')
                value = int(parts[0]) / int(parts[1])
                print(f"DEBUG: Parsed as fraction: {value}")
            else:
                value = float(value_str)
                print(f"DEBUG: Parsed as decimal: {value}")
            
            result['D'] = value
            half_value = value / 2
            result['Dl'] = half_value
            result['Dr'] = half_value
            result['DL'] = half_value
            result['DR'] = half_value
            
            print(f"DEBUG: Final result - D={result['D']}, Dl={result['Dl']}, Dr={result['Dr']}")
            return result
    
    # If not found in header, search bottom rows (skip data rows)
    print(f"DEBUG: Not found in header, searching bottom rows...")
    for idx in range(len(df_raw) - 1, max(len(df_raw) - 20, header_row), -1):
        row_text = ' '.join([str(x) for x in df_raw.iloc[idx].tolist() if pd.notna(x)])
        
        # Skip if this looks like a data row
        first_col = df_raw.iloc[idx, 0] if len(df_raw.columns) > 0 else None
        if pd.notna(first_col):
            try:
                float(str(first_col).strip())
                continue  # Skip data rows
            except (ValueError, TypeError):
                pass
        
        # Look for D=value pattern
        d_pattern = re.search(r'D\s*[=:]\s*((?:\d+/\d+)|(?:\d*\.?\d+))', row_text, re.IGNORECASE)
        if d_pattern:
            value_str = d_pattern.group(1)
            print(f"DEBUG: Found D pattern in notes row {idx}: {row_text[:200]}")
            
            if '/' in value_str:
                value = int(value_str.split('/')[0]) / int(value_str.split('/')[1])
            else:
                value = float(value_str)
            
            result['D'] = value
            half_value = value / 2
            result['Dl'] = half_value
            result['Dr'] = half_value
            result['DL'] = half_value
            result['DR'] = half_value
            
            print(f"DEBUG: Final result - D={result['D']}, Dl={result['Dl']}, Dr={result['Dr']}")
            return result
    
    return result


def parse_deduction_value(deduct_code: str, numeric_value: float, notes_text: str = "", default_values: Dict[str, float] = None) -> Optional[float]:
    """
    Parse deduction value from Deducts code.
    
    Simple logic:
    - If Deducts = "D": use D value from notes (total value)
    - If Deducts = "Dl": use Dl value from notes (which is D/2)
    - If Deducts = "Dr": use Dr value from notes (which is D/2)
    
    If code has explicit value (e.g., "D=1" or "D-1/2"), parse and return that.
    Otherwise, use default_values from notes.
    """
    if pd.isna(deduct_code):
        return None
    
    deduct_code_str = str(deduct_code).strip()
    deduct_code_upper = deduct_code_str.upper()
    
    # Extract code part (D, Dl, or Dr) - handle formats like "D-1/2" or just "D"
    code_match = re.match(r'^(D[LR]?)\s*[-=]\s*((?:\d+/\d+)|(?:\d+\.?\d*))', deduct_code_upper)
    
    if code_match:
        # Code has explicit value (e.g., "D=1" or "D-1/2")
        code_part = code_match.group(1).upper()
        value_str = code_match.group(2)
        
        # Parse value
        if '/' in value_str:
            # Fraction (e.g., "1/2")
            parts = value_str.split('/')
            value = int(parts[0]) / int(parts[1])
        else:
            # Decimal (e.g., "1")
            value = float(value_str)
        
        # For D codes:
        # - "D=1" means total is 1.0, return as-is
        # - "D-1/2" means per-side is 1/2, total is 1.0, return 1.0
        if code_part == 'D':
            if '=' in deduct_code_str:
                return value  # D=1 → total is 1.0
            else:
                return value * 2  # D-1/2 → per-side is 1/2, total is 1.0
        else:
            # Dl or Dr: value is per-side, return as-is
            return value
    
    # No explicit value in code, check numeric_value column
    if not pd.isna(numeric_value):
        if deduct_code_upper == 'D':
            return float(numeric_value) * 2  # Total = 2 * per-side
        else:
            return float(numeric_value)  # Per-side value
    
    # No explicit value, use default_values from notes
    if default_values:
        if deduct_code_upper == 'D':
            # Return D value (total)
            d_value = default_values.get('D', 1.0)
            if 'D' not in default_values:
                print(f"DEBUG parse_deduction_value: 'D' not in default_values, using fallback 1.0")
            else:
                print(f"DEBUG parse_deduction_value: Found D={d_value} in default_values")
            return d_value
        elif deduct_code_upper in ['DL', 'Dl']:
            # Return Dl value (which is D/2) - check both cases
            return default_values.get('Dl', default_values.get('DL', 0.5))
        elif deduct_code_upper in ['DR', 'Dr']:
            # Return Dr value (which is D/2) - check both cases
            return default_values.get('Dr', default_values.get('DR', 0.5))
    
    # No default values, use fallback defaults
    if deduct_code_upper == 'D':
        return 1.0  # Default total
    else:
        return 0.5  # Default per-side


def format_control_length(chain_value: float) -> Optional[str]:
    """Format chain value as ControlLength (e.g., 72.0 -> '72"LOOP')"""
    if pd.isna(chain_value):
        return None
    return f'{int(chain_value)}"LOOP'


def create_special_instructions(
    has_reverse_roll: bool,
    control_side: str,
    extra_fabric_deduction: Optional[float],
    deduction_code: Optional[str] = None
) -> str:
    """Generate SpecialInstructions based on reverse roll and deductions
    
    Formatting matches image exactly (line by line):
    - When ReverseRoll=Yes and no deduction: "Apply Punch Reverse Fascia !!!"
    - When ReverseRoll=Yes and has deduction:
      Line 1: "Apply Punch Reverse Fascia !!!"
      Line 2: "--"
      Line 3: "Fabric Deduction 1/2" from LEFT SIDE ONLY!!!" (or "from both sides" for D code)
    - When ReverseRoll=No: Only deduction info if present, no Punch Reverse Fascia
    
    Args:
        deduction_code: "D" (both sides), "Dl" (left), "Dr" (right), or None
    """
    instructions = []
    
    # Determine side text based on deduction code (not control_side)
    # If deduction code is Dl/Dr, use that; otherwise use control_side
    if deduction_code:
        if deduction_code.upper() == 'D':
            side_text = "both sides"
        elif deduction_code.upper() == 'DL':
            side_text = "LEFT SIDE ONLY"
        elif deduction_code.upper() == 'DR':
            side_text = "RIGHT SIDE ONLY"
        else:
            # Fallback to control_side if deduction_code is unexpected
            side_text = "LEFT SIDE ONLY" if control_side == "LEFT" else "RIGHT SIDE ONLY"
    else:
        # No deduction code, use control_side
        side_text = "LEFT SIDE ONLY" if control_side == "LEFT" else "RIGHT SIDE ONLY"
    
    # Only add Punch Reverse Fascia when ReverseRoll is Yes
    if has_reverse_roll:
        # Add header line matching user request
        instructions.append(".---")
        instructions.append("Apply Punch Reverse Fascia !!!")
        
        # Add deduction info if present
        if extra_fabric_deduction is not None and extra_fabric_deduction > 0:
            # Format deduction (handle fractions)
            if extra_fabric_deduction == 0.5:
                deduction_text = '1/2"'
            elif extra_fabric_deduction == 0.25:
                deduction_text = '1/4"'
            elif extra_fabric_deduction == 1.0:
                deduction_text = '1"'
            else:
                deduction_text = f'{extra_fabric_deduction}"'
            
            # Add separator (3 dashes) and deduction line
            instructions.append("---")
            instructions.append(f"Fabric Deduction {deduction_text} from {side_text} !!!")
    else:
        # When ReverseRoll is No, only add deduction info if present
        if extra_fabric_deduction is not None and extra_fabric_deduction > 0:
            if extra_fabric_deduction == 0.5:
                deduction_text = '1/2"'
            elif extra_fabric_deduction == 0.25:
                deduction_text = '1/4"'
            elif extra_fabric_deduction == 1.0:
                deduction_text = '1"'
            else:
                deduction_text = f'{extra_fabric_deduction}"'
            
            # Just add the deduction line (no header/branding for now unless requested)
            instructions.append(f"Fabric Deduction {deduction_text} from {side_text} !!!")
    
    return "\n".join(instructions) if instructions else ""


def determine_reverse_roll(roll_value: str) -> Optional[str]:
    """Convert Roll value to ReverseRoll (Rev -> Yes)"""
    if pd.isna(roll_value):
        return None
    if str(roll_value).strip().upper() == "REV":
        return "Yes"
    return None


def extract_color_codes_from_header(df_raw: pd.DataFrame, header_row: int = 8) -> Dict[str, str]:
    """
    Extract color codes from header/metadata section (rows 0-7).
    
    Looks for patterns like:
    - "Bed = COLOR123" or "Bedroom = COLOR123"
    - "Living = COLOR456" or "Liv = COLOR456"
    - "Bed/Bedroom = COLOR123"
    - "Living/Liv = COLOR456"
    - "Studio = COLOR789"
    - "Den = COLOR012"
    - "Kitchen = COLOR345"
    - "Bath = COLOR678"
    
    Returns dict with extracted colors: {'Bed': 'COLOR123', 'Liv': 'COLOR456', 'Studio': 'COLOR789', ...}
    """
    color_codes = {}
    
    # Define all fabric types and their patterns
    fabric_patterns = {
        'Bed': [
            r'Bed\s*[/=]\s*Bedroom\s*=\s*([A-Z0-9]+)',
            r'Bedroom\s*=\s*([A-Z0-9]+)',
            r'Bed\s*=\s*([A-Z0-9]+)',
        ],
        'Liv': [
            r'Living\s*[/=]\s*Liv\s*=\s*([A-Z0-9]+)',
            r'Living\s*=\s*([A-Z0-9]+)',
            r'Liv\s*=\s*([A-Z0-9]+)',
        ],
        'Studio': [
            r'Studio\s*=\s*([A-Z0-9]+)',
        ],
        'Den': [
            r'Den\s*=\s*([A-Z0-9]+)',
        ],
        'Kitchen': [
            r'Kitchen\s*=\s*([A-Z0-9]+)',
        ],
        'Bath': [
            r'Bath\s*=\s*([A-Z0-9]+)',
            r'Bathroom\s*=\s*([A-Z0-9]+)',
        ],
    }
    
    # Search header rows (0 to header_row-1)
    for idx in range(header_row):
        row_text = ' '.join([str(x) for x in df_raw.iloc[idx].tolist() if pd.notna(x)])
        
        # Check each fabric type
        for fabric_type, patterns in fabric_patterns.items():
            # Skip if already found
            if fabric_type in color_codes:
                continue
                
            for pattern in patterns:
                match = re.search(pattern, row_text, re.IGNORECASE)
                if match:
                    color_codes[fabric_type] = match.group(1).strip()
                    break
    
    return color_codes


def clean_excel_file(
    input_file: str,
    output_file: Optional[str] = None,
    bed_color: Optional[str] = None,
    liv_color: Optional[str] = None,
    fabric_colors: Optional[Dict[str, str]] = None,
    header_row: int = 8,
    notes_start_keywords: list = None,
    deduction_cell: Optional[str] = None
) -> pd.DataFrame:
    """
    Clean raw Excel file into standardized format.
    
    Parameters:
    -----------
    input_file : str
        Path to input Excel file
    output_file : str, optional
        Path to save cleaned file (if None, auto-generates name)
    bed_color : str, optional
        Color number for Bed fabric (if None, will try to detect or prompt)
    liv_color : str, optional
        Color number for Liv fabric (if None, will try to detect or prompt)
    fabric_colors : dict, optional
        Dictionary mapping fabric types to color numbers
        Example: {'Kitchen': 'KITCHEN_COLOR', 'Studio': 'STUDIO_COLOR', 'Den': 'DEN_COLOR'}
    header_row : int
        Row number (0-indexed) where data headers are located
    notes_start_keywords : list
        Keywords that indicate start of notes section (rows to exclude)
    
    Returns:
    --------
    pd.DataFrame : Cleaned dataframe
    """
    if notes_start_keywords is None:
        notes_start_keywords = ["Total", "all Finshed", "Punch Reverse", "Deducts"]
    
    # Read the Excel file
    print(f"Reading file: {input_file}")
    df_raw = pd.read_excel(input_file, sheet_name=0, header=None)
    
    # Extract header row
    headers = df_raw.iloc[header_row].tolist()
    
    # Extract color codes from header section
    header_colors = extract_color_codes_from_header(df_raw, header_row)
    
    # Use header colors if not provided as parameters
    if not bed_color and 'Bed' in header_colors:
        bed_color = header_colors['Bed']
        print(f"Found Bed color in header: {bed_color}")
    if not liv_color and 'Liv' in header_colors:
        liv_color = header_colors['Liv']
        print(f"Found Liv color in header: {liv_color}")
    
    # Merge header colors into fabric_colors dict (header colors take precedence)
    if fabric_colors is None:
        fabric_colors = {}
    
    # Add all header colors to fabric_colors (except Bed/Liv which are handled separately)
    for fabric_type, color_code in header_colors.items():
        if fabric_type not in ['Bed', 'Liv']:
            fabric_colors[fabric_type] = color_code
            print(f"Found {fabric_type} color in header: {color_code}")
    
    # Extract deduction values from notes section before reading data
    # Try specific cell first (default to I6)
    if deduction_cell is None:
        deduction_cell = 'I6'  # Default location
    deduction_values = extract_deduction_values_from_notes(df_raw, header_row, deduction_cell=deduction_cell)
    print(f"DEBUG: Notes extraction returned: {deduction_values}")
    if deduction_values:
        print(f"DEBUG: Extracted deduction values from notes:")
        for key, value in deduction_values.items():
            print(f"  {key}: {value}")
    else:
        print("DEBUG: No deduction values found in notes - will use defaults")
    
    # Read data starting from header_row + 1
    df = pd.read_excel(input_file, sheet_name=0, header=header_row)
    
    # Remove rows that are notes (check for keywords in Tag or other columns)
    print("Removing note rows...")
    mask = pd.Series([True] * len(df))
    
    for col in df.columns:
        if df[col].dtype == 'object':
            for keyword in notes_start_keywords:
                mask &= ~df[col].astype(str).str.contains(keyword, case=False, na=False)
    
    # Also remove rows where Tag is text (not numeric) and not NaN
    # Handle different column names: "Tag", "Tag/Unit", "Tag/unit", or "unit" (case-insensitive)
    tag_col = None
    for col in df.columns:
        col_upper = str(col).upper()
        if col_upper in ['TAG', 'TAG/UNIT', 'UNIT']:
            tag_col = col
            break
    
    # Separate rows with non-numeric tags (alphabets) into skipped_df
    skipped_df = pd.DataFrame()
    if tag_col:
        # Identify rows where Tag is NOT numeric AND NOT NaN (i.e., contains text/alphabets)
        # We want to keep NaNs in the main df (they get forward filled later)
        non_numeric_mask = df[tag_col].apply(lambda x: isinstance(x, str) and not pd.isna(x) and not str(x).replace('.','',1).isdigit())
        
        # rows that are notes/garbage are already filtered by 'mask' above, but we only want to save
        # "real" data rows that were skipped due to alphabets in tag.
        # But 'df' currently still contains everything because we haven't applied 'mask' yet.
        # Wait, we should apply 'mask' (notes filter) first to avoid filling the Skipped sheet with garbage.
        
        # 1. Apply notes/garbage filter first
        df = df[mask].copy()
        
        # 2. Now identify non-numeric tags in the CLEANED set of candidate rows
        # Re-evaluate non_numeric_mask on the filtered df
        non_numeric_mask = df[tag_col].apply(lambda x: isinstance(x, str) and not pd.isna(x) and not str(x).replace('.','',1).isdigit())
        
        # Capture skipped rows
        skipped_df = df[non_numeric_mask].copy()
        
        # Remove skipped rows from main df
        df = df[~non_numeric_mask].copy()
        print(f"DEBUG: Found {len(skipped_df)} rows with non-numeric tags (moved to Skipped sheet)")
    else:
        # Just apply the notes mask if no tag column
        df = df[mask].copy()
    
    # Forward fill Tag values (rows with NaN Tag belong to previous Tag)
    # Handle different column names: "Tag", "Tag/Unit", "Tag/unit", or "unit" (case-insensitive)
    tag_col = None
    for col in df.columns:
        col_upper = str(col).upper()
        if col_upper in ['TAG', 'TAG/UNIT', 'UNIT']:
            tag_col = col
            break
    
    if tag_col:
        df[tag_col] = df[tag_col].ffill()
        # Convert to numeric where possible, keeping non-numeric as object
        try:
            df[tag_col] = pd.to_numeric(df[tag_col])
        except (ValueError, TypeError):
            # If conversion fails, keep as is (may contain text notes)
            pass
    
    # Remove rows where essential columns are all NaN
    # Handle both 'Height' and 'Length' column names
    height_col = 'Height' if 'Height' in df.columns else ('Length' if 'Length' in df.columns else None)
    essential_cols = ['Width', 'Fabric']
    if height_col:
        essential_cols.append(height_col)
    if all(col in df.columns for col in essential_cols):
        df = df[df[essential_cols].notna().any(axis=1)].copy()
    
    # Debug: Check column names
    print(f"Available columns: {list(df.columns)}")
    print(f"Rows after filtering: {len(df)}")
    
    # Build cleaned dataframe
    cleaned_data = []
    
    # Determine tag column name - handle multiple possible names (case-insensitive)
    tag_col = None
    for col in df.columns:
        col_upper = str(col).upper()
        if col_upper in ['TAG', 'TAG/UNIT', 'UNIT']:
            tag_col = col
            break
    
    skipped_count = 0
    skipped_reasons = {'no_fabric': 0, 'no_width': 0, 'no_height': 0}
    
    for idx, row in df.iterrows():
        tag = row.get(tag_col, '') if tag_col else ''
        fabric_raw = row.get('Fabric', '')
        
        # Skip if essential data is missing
        if pd.isna(fabric_raw):
            skipped_reasons['no_fabric'] += 1
            skipped_count += 1
            continue
        
        fabric = str(fabric_raw).strip()
        
        # Clean and parse width (handle fractions including Unicode: ¾, ¼, ⅛, etc.)
        width_raw = row.get('Width')
        if not pd.isna(width_raw):
            width = parse_fraction_string(width_raw)
        else:
            width = None
        
        # Handle both 'Height' and 'Length' column names (handle fractions including Unicode)
        height_raw = row.get('Height') if 'Height' in row.index else row.get('Length', None)
        if not pd.isna(height_raw):
            height = parse_fraction_string(height_raw)
        else:
            height = None
        
        # Skip if width or height couldn't be parsed
        if width is None or height is None:
            if width is None:
                skipped_reasons['no_width'] += 1
            if height is None:
                skipped_reasons['no_height'] += 1
            skipped_count += 1
            continue
        
        # Handle both 'Control' and 'Controls' column names
        control = str(row.get('Control', row.get('Controls', ''))).strip().upper()
        # Handle both 'Deducts' and 'Deducts ' (with trailing space)
        deducts = row.get('Deducts ', '') if 'Deducts ' in row.index else row.get('Deducts', '')
        # Check if Unnamed: 12 column exists (some files don't have it)
        deduct_value = row.get('Unnamed: 12', None) if 'Unnamed: 12' in df.columns else None
        roll = row.get('Roll', '')
        chain = row.get('Chain ', None)
        
        # Determine color number based on fabric type
        # Handle Bed, Liv, and other fabric types (Kitchen, Studio, Den, etc.)
        fabric_upper = fabric.upper()
        color_number = None
        
        # Check fabric_colors dict first (for custom mappings)
        if fabric_colors and fabric in fabric_colors:
            color_number = fabric_colors[fabric]
        elif fabric_colors:
            # Try case-insensitive match
            for fab_key, fab_color in fabric_colors.items():
                if fab_key.upper() == fabric_upper:
                    color_number = fab_color
                    break
        
        # If not found in fabric_colors, use standard mappings
        if color_number is None:
            if 'BED' in fabric_upper:
                color_number = bed_color if bed_color else 'BED_COLOR_NEEDED'
            elif 'LIV' in fabric_upper:
                color_number = liv_color if liv_color else 'LIV_COLOR_NEEDED'
            else:
                # For other fabric types (Kitchen, Studio, Den, etc.), use placeholder
                # User can manually update these later
                color_number = f'{fabric.upper()}_COLOR_NEEDED'
        
        # Create Room (Tag-Fabric)
        # Handle tag: can be NaN, empty string, or numeric
        if pd.isna(tag) or tag == '' or str(tag).strip() == '':
            tag_str = 'UNKNOWN'
        else:
            try:
                tag_str = str(int(float(tag)))  # Convert to float first to handle decimals, then int
            except (ValueError, TypeError):
                tag_str = 'UNKNOWN'
        room = f"{tag_str}-{fabric}"
        
        # Map Control to ControlSide
        control_side_map = {'L': 'LEFT', 'R': 'RIGHT'}
        control_side = control_side_map.get(control, None)
        
        # Parse deduction - SIMPLE LOGIC
        # When notes contain "D=1":
        #   - Deducts = "D"  → ExtraFabricDeduction = 1.0 (total)
        #   - Deducts = "Dl" → ExtraFabricDeduction = 0.5 (half of D)
        #   - Deducts = "Dr" → ExtraFabricDeduction = 0.5 (half of D)
        extra_fabric_deduction = None
        deduction_code = None  # Track the deduction code (D, Dl, or Dr) for special instructions
        deduction_per_side_value = None  # Track per-side value for "D" deductions (for display in special instructions)
        
        if not pd.isna(deducts) and str(deducts).strip() != '':
            deducts_str = str(deducts).strip()
            
            # Extract code part (D, Dl, or Dr)
            code_match = re.match(r'^(D[LR]?)', deducts_str, re.IGNORECASE)
            if code_match:
                code_part = code_match.group(1).upper()
                deduction_code = code_part  # Store for special instructions
                
                # Get the deduction value using parse_deduction_value
                extra_fabric_deduction = parse_deduction_value(deducts_str, deduct_value, "", deduction_values)
                
                # Debug output for first few rows
                if idx < 3:
                    print(f"DEBUG Row {idx}: Deducts='{deducts_str}', code_part='{code_part}', extra_fabric_deduction={extra_fabric_deduction}")
                
                # Extract per-side value for "D" deductions (for display in special instructions only)
                if code_part == 'D' and extra_fabric_deduction is not None:
                    # Per-side value is half of total for display (e.g., D=0.5 → show "1/4" in special instructions)
                    deduction_per_side_value = extra_fabric_deduction / 2
                    if idx < 3:
                        print(f"DEBUG Row {idx}: deduction_per_side_value={deduction_per_side_value}")
        
        # ControlLength - set for all rows that have Chain value (present in raw file)
        control_length = format_control_length(chain) if chain else None
        
        # ReverseRoll - if Roll="Rev" then "Yes", otherwise "No" (not None)
        reverse_roll = determine_reverse_roll(roll)
        if reverse_roll is None:
            reverse_roll = "No"
        
        # Special Instructions
        # For "D" deductions, use per-side value for display (e.g., show "1/2"" instead of "1"")
        display_deduction_value = deduction_per_side_value if (deduction_code == 'D' and deduction_per_side_value is not None) else extra_fabric_deduction
        special_instructions = create_special_instructions(
            reverse_roll == "Yes",
            control_side or "",
            display_deduction_value,
            deduction_code
        )
        
        # Set ExtraFabricDeduction to 0 if None/NaN (no deduction)
        if extra_fabric_deduction is None:
            extra_fabric_deduction = 0.0
        
        cleaned_data.append({
            'Color Number': color_number,
            'Width': width,
            'Height': int(height) if height is not None else None,
            'ExtraFabricDeduction': extra_fabric_deduction,
            'ReverseRoll': reverse_roll,
            'ControlSide': control_side,
            'ControlLength': control_length,
            'Room': room,
            'SpecialInstructions': special_instructions
        })
    
    df_cleaned = pd.DataFrame(cleaned_data)
    
    # Clean up trailing spaces in Room and ControlSide
    if 'Room' in df_cleaned.columns:
        df_cleaned['Room'] = df_cleaned['Room'].str.replace(r'\s+$', '', regex=True)
    if 'ControlSide' in df_cleaned.columns:
        df_cleaned['ControlSide'] = df_cleaned['ControlSide'].str.replace(r'\s+$', '', regex=True)
    
    # Save to file if output_file specified
    if output_file:
        output_path = output_file
    else:
        # Auto-generate output filename
        input_path = Path(input_file)
        output_path = input_path.parent / f"{input_path.stem}-cleaned.xlsx"
    
    # Save with wrap text formatting for SpecialInstructions column and highlighting for width > 144
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_cleaned.to_excel(writer, index=False, sheet_name='Cleaned')
        
        # Get the worksheet for formatting
        worksheet = writer.sheets['Cleaned']
        
        # Yellow fill for highlighting
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        
        # Find column indices
        width_col_idx = df_cleaned.columns.get_loc('Width') + 1 if 'Width' in df_cleaned.columns else None
        special_instructions_col_idx = df_cleaned.columns.get_loc('SpecialInstructions') + 1 if 'SpecialInstructions' in df_cleaned.columns else None
        
        # Apply formatting to each row
        for row_num in range(2, len(df_cleaned) + 2):  # Start from row 2 (row 1 is header)
            # Highlight width cell if value > 144
            if width_col_idx:
                width_value = df_cleaned.iloc[row_num - 2]['Width']  # -2 because Excel is 1-indexed and has header
                if pd.notna(width_value) and width_value > 144:
                    width_col_letter = openpyxl.utils.get_column_letter(width_col_idx)
                    worksheet[f'{width_col_letter}{row_num}'].fill = yellow_fill
            
            # Enable wrap text for SpecialInstructions column
            if special_instructions_col_idx:
                col_letter = openpyxl.utils.get_column_letter(special_instructions_col_idx)
                cell = worksheet[f'{col_letter}{row_num}']
                cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        # Enable wrap text for SpecialInstructions header
        if special_instructions_col_idx:
            header_col_letter = openpyxl.utils.get_column_letter(special_instructions_col_idx)
            header_cell = worksheet[f'{header_col_letter}1']
            header_cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        
        # Write Skipped rows to a separate sheet if any exist
        if not skipped_df.empty:
            skipped_df.to_excel(writer, index=False, sheet_name='Skipped')
            print(f"Added 'Skipped' sheet with {len(skipped_df)} rows")
    
    print(f"Cleaned file saved to: {output_path}")
    
    if skipped_count > 0:
        print(f"\nDEBUG: Skipped {skipped_count} rows during processing:")
        for reason, count in skipped_reasons.items():
            if count > 0:
                print(f"  - {reason}: {count}")
    
    return df_cleaned


def detect_color_numbers(df_cleaned: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    """Try to detect color numbers from cleaned data (if available)"""
    bed_color = None
    liv_color = None
    
    if 'Color Number' in df_cleaned.columns and 'Room' in df_cleaned.columns:
        bed_rows = df_cleaned[df_cleaned['Room'].str.contains('-Bed', na=False, case=False)]
        liv_rows = df_cleaned[df_cleaned['Room'].str.contains('-Liv', na=False, case=False)]
        
        if len(bed_rows) > 0:
            bed_colors = bed_rows['Color Number'].unique()
            if len(bed_colors) == 1:
                bed_color = bed_colors[0]
        
        if len(liv_rows) > 0:
            liv_colors = liv_rows['Color Number'].unique()
            if len(liv_colors) == 1:
                liv_color = liv_colors[0]
    
    return bed_color, liv_color


if __name__ == "__main__":
    import sys
    
    # Simple command-line interface (backward compatible)
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        bed_color = sys.argv[3] if len(sys.argv) > 3 else None
        liv_color = sys.argv[4] if len(sys.argv) > 4 else None
        deduction_cell = sys.argv[5] if len(sys.argv) > 5 else 'I6'  # Default to I6
    else:
        # Default: clean the example file
        input_file = "Symington Mid Rise L1 to L3 .xlsx"
        output_file = None
        bed_color = "YUNOWH"  # From the cleaned example
        liv_color = "PWS3WHIT"  # From the cleaned example
        deduction_cell = 'I6'  # Default cell for D value
    
    fabric_colors = None
    
    print("=" * 60)
    print("Excel Cleaning Script")
    print("=" * 60)
    print(f"Reading D value from cell: {deduction_cell}")
    print("(You can change this by passing a 5th argument, e.g., 'I6')")
    print()
    
    if not bed_color or not liv_color:
        print("\nWARNING: Color numbers not provided.")
        print("   Bed color:", bed_color or "NOT SET")
        print("   Liv color:", liv_color or "NOT SET")
        print("   You can provide them as command line arguments:")
        print("   python clean_excel.py <input_file> [output_file] [bed_color] [liv_color] [deduction_cell]")
        print()
    
    df_cleaned = clean_excel_file(
        input_file=input_file,
        output_file=output_file,
        bed_color=bed_color,
        liv_color=liv_color,
        fabric_colors=fabric_colors,
        header_row=8,  # Default header row
        deduction_cell=deduction_cell
    )
    
    print("\n" + "=" * 60)
    print("Cleaning Complete!")
    print("=" * 60)
    print(f"\nCleaned {len(df_cleaned)} rows")
    print(f"\nFirst few rows:")
    print(df_cleaned.head(10).to_string())
    
    # Check for missing color numbers
    if 'Color Number' in df_cleaned.columns:
        missing_colors = df_cleaned[df_cleaned['Color Number'].str.contains('NEEDED', na=False)]
        if len(missing_colors) > 0:
            print(f"\nWARNING: {len(missing_colors)} rows have missing color numbers")
            unique_fabrics = missing_colors['Room'].str.split('-').str[1].unique()
            print(f"   Fabric types needing colors: {', '.join(unique_fabrics)}")
            print("   You can provide colors via fabric_colors parameter or update manually")

