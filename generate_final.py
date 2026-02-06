import sys
sys.path.insert(0, r'c:\Users\mdash\Downloads\Excel-clean')

# Read the template creation script
with open(r'c:\Users\mdash\Downloads\Excel-clean\create_standard_template.py', 'r') as f:
    code = f.read()

# Replace output filename
code = code.replace("wb.save('STANDARD_TEMPLATE.xlsx')", "wb.save(r'c:\\Users\\mdash\\Downloads\\Excel-clean\\FINAL_TEMPLATE.xlsx')")
code = code.replace("print('✅ Created STANDARD_TEMPLATE.xlsx')", "print('✅ Created FINAL_TEMPLATE.xlsx')")

# Execute
exec(code)
