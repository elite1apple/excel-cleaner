import pandas as pd
from clean_excel import clean_excel_file
import os

# Create a sample dataframe with some non-numeric tags
data = {
    'Tag': [101, '102A', 103, 'Office', 105],
    'Fabric': ['Bed', 'Liv', 'Bed', 'Liv', 'Bed'],
    'Width': [50, 60, 70, 80, 90],
    'Height': [50, 60, 70, 80, 90],
    'Control': ['L', 'R', 'L', 'R', 'L']
}

df = pd.DataFrame(data)

# Create a test Excel file
input_file = 'test_skipped_input.xlsx'
# Add some header rows to simulate real file
with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
    # Write empty header rows
    pd.DataFrame(['Header info']).to_excel(writer, startrow=0, index=False, header=False)
    # Write data at row 8 (0-indexed) -> row 9 in Excel
    df.to_excel(writer, startrow=8, index=False)

print(f"Created test file: {input_file}")

# Run the cleaner
output_file = 'test_skipped_output.xlsx'
# Note: clean_excel checks for "Tag" or "Tag/Unit" case insensitive
print("Running cleaner...")
try:
    clean_excel_file(
        input_file=input_file, 
        output_file=output_file, 
        header_row=8,
        bed_color="WHITE",
        liv_color="CREAM"
    )
    
    # Check the output
    print("\nVerifying output...")
    xls = pd.ExcelFile(output_file)
    print(f"Sheet names found: {xls.sheet_names}")
    
    if 'Skipped' in xls.sheet_names:
        skipped_df = pd.read_excel(output_file, sheet_name='Skipped')
        print(f"\nSkipped Sheet Content ({len(skipped_df)} rows):")
        print(skipped_df)
        
        # Verify specific rows are skipped
        skipped_tags = skipped_df['Tag'].astype(str).tolist()
        expected_skipped = ['102A', 'Office']
        
        all_found = True
        for tag in expected_skipped:
            if tag not in skipped_tags:
                print(f"❌ Error: Expected tag '{tag}' to be skipped, but it wasn't.")
                all_found = False
            else:
                print(f"✅ Verified tag '{tag}' is in Skipped sheet.")
                
        if len(skipped_df) == 2 and all_found:
             print("\n✅ SUCCESS: Skipped rows logic working correctly!")
        else:
             print(f"\n❌ FAILURE: Expected 2 skipped rows, found {len(skipped_df)}")
    else:
        print("\n❌ FAILURE: 'Skipped' sheet not created!")
        
    # Verify Cleaned sheet
    cleaned_df = pd.read_excel(output_file, sheet_name='Cleaned')
    print(f"\nCleaned Sheet Content ({len(cleaned_df)} rows):")
    # Expected: 101, 103, 105 (3 rows)
    if len(cleaned_df) == 3:
        print("✅ Correct number of cleaned rows.")
    else:
        print(f"❌ Incorrect number of cleaned rows. Expected 3, got {len(cleaned_df)}")

except Exception as e:
    print(f"An error occurred: {e}")
