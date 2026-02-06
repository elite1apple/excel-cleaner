import pandas as pd

df = pd.read_excel('FINAL_TEMPLATE-cleaned.xlsx')
print("--- Special Instructions Verification ---")
for idx, row in df.iterrows():
    instr = row['SpecialInstructions']
    if pd.notna(instr) and str(instr).strip() != '':
        print(f"Row {idx}:")
        print(instr)
        print("-" * 20)
        if idx >= 3:
            break
