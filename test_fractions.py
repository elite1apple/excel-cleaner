"""Test all 1/8 fraction variants"""
from clean_excel import parse_fraction_string

print("Testing ALL 1/8 Fractions:")
print("=" * 60)

# All fractions of 1/8
test_cases = [
    # Pure fractions
    ("⅛", 0.125, "1/8"),
    ("¼", 0.25, "2/8 = 1/4"),
    ("⅜", 0.375, "3/8"),
    ("½", 0.5, "4/8 = 1/2"),
    ("⅝", 0.625, "5/8"),
    ("¾", 0.75, "6/8 = 3/4"),
    ("⅞", 0.875, "7/8"),
    
    # Mixed numbers (whole + fraction)
    ("54⅛", 54.125, "54 and 1/8"),
    ("54¼", 54.25, "54 and 1/4"),
    ("54⅜", 54.375, "54 and 3/8"),
    ("54½", 54.5, "54 and 1/2"),
    ("54⅝", 54.625, "54 and 5/8"),
    ("54¾", 54.75, "54 and 3/4"),
    ("54⅞", 54.875, "54 and 7/8"),
    
    # Alternative text format
    ("51 1/8", 51.125, "51 and 1/8 (text)"),
    ("51 3/8", 51.375, "51 and 3/8 (text)"),
    ("51 5/8", 51.625, "51 and 5/8 (text)"),
    ("51 7/8", 51.875, "51 and 7/8 (text)"),
]

all_passed = True
for input_str, expected, description in test_cases:
    result = parse_fraction_string(input_str)
    passed = abs(result - expected) < 0.0001 if result else False
    status = "✅" if passed else "❌"
    
    if not passed:
        all_passed = False
    
    print(f"{status} {input_str:8s} → {result:6.3f} (expected {expected:6.3f}) - {description}")

print("=" * 60)
if all_passed:
    print("🎉 ALL 1/8 FRACTIONS WORK PERFECTLY!")
else:
    print("⚠️ Some tests failed")
