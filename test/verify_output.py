import os
import sys
import glob
import openpyxl

# Load module
import importlib.util
spec = importlib.util.spec_from_file_location("msl_v16", os.path.abspath("src/MSLdata_check_v16.py"))
msl_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(msl_mod)

sample_files = glob.glob('bench/*サンプル*.xlsx')
if not sample_files:
    sample_files = [x for x in glob.glob('bench/*.xlsx') if 'bench_op' not in x]
input_file = sample_files[0]
print(f"Testing MSLdata_check_v15 with input: {input_file}")

res = msl_mod.process_attendance_data(input_file, start_from_saturday=True)
output_file = res[4]
print(f"Target output file: {output_file}")

msl_mod.format_excel_fast(*res)
print("format_excel_fast completed successfully.")

# Verify with openpyxl
wb = openpyxl.load_workbook(output_file, read_only=False)
sheet_names = wb.sheetnames
print("Generated Sheets:", sheet_names)

assert "全集計" in sheet_names, "Missing 全集計"
assert "除外済み集計" in sheet_names, "Missing 除外済み集計"
assert "Lap改善ペアリスト" in sheet_names, "Missing Lap改善ペアリスト"

ws_lap = wb["Lap改善ペアリスト"]
print(f"Lap改善ペアリスト max_row: {ws_lap.max_row}, max_column: {ws_lap.max_column}")

headers = [ws_lap.cell(row=2, column=c).value for c in range(1, 13)]
print("Lap改善ペアリスト Headers (Row 2):", headers)

expected_headers = [
    "教室", "講師No", "担当講師名", "生徒名", "学年", "科目",
    "授業数", "Lap回数", "未実施数", "ペアLap％", "講師全体Lap％", "生徒全体Lap％"
]
assert headers == expected_headers, f"Headers mismatch! Expected: {expected_headers}, got: {headers}"

print("First 5 data rows in Lap改善ペアリスト:")
for r in range(4, min(9, ws_lap.max_row + 1)):
    row_vals = [ws_lap.cell(row=r, column=c).value for c in range(1, 13)]
    print(f"Row {r}: {row_vals}")

wb.close()
print("All validations PASSED successfully!")

