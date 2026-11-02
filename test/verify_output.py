import os
import sys
import glob
import openpyxl

# Load module
import importlib.util
spec = importlib.util.spec_from_file_location("msl_v17", os.path.abspath("src/MSLdata_check_v17.py"))
msl_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(msl_mod)

sample_files = glob.glob('bench/*sample*.xlsx') + glob.glob('bench/*サンプル*.xlsx')
if not sample_files:
    sample_files = [x for x in glob.glob('bench/*.xlsx') if 'bench_op' not in x]
if not sample_files:
    print("[ERROR] No sample Excel files found in bench/ directory.")
    sys.exit(1)
input_file = sample_files[0]
print(f"Testing MSLdata_check_v17 with input: {input_file}")

res = msl_mod.process_attendance_data(input_file, start_from_saturday=True)
output_file = res[4]
print(f"Target output file: {output_file}")

msl_mod.format_excel_fast(*res)
print("format_excel_fast completed successfully.")

# Verify with openpyxl
wb = openpyxl.load_workbook(output_file, read_only=False)
sheet_names = wb.sheetnames
print("Generated Sheets:", sheet_names)

expected_sheets = ["全集計", "除外済み集計", "全集計ペアリスト", "除外済みペアリスト"]
assert sheet_names == expected_sheets, f"Sheet order mismatch! Expected: {expected_sheets}, got: {sheet_names}"

expected_headers = [
    "教室", "講師No", "担当講師名", "生徒名", "学年", "科目",
    "授業数", "Lap回数", "未実施数", "ペアLap％", "講師全体Lap％", "生徒全体Lap％"
]

for s_name, criteria in [("全集計ペアリスト", "全集計データ基準"), ("除外済みペアリスト", "除外済みデータ基準")]:
    ws = wb[s_name]
    print(f"\n--- Checking {s_name} ---")
    print(f"max_row: {ws.max_row}, max_column: {ws.max_column}")
    
    title_val = ws.cell(row=1, column=1).value
    print(f"Title (Row 1): {title_val}")
    assert criteria in title_val, f"Title criteria mismatch in {s_name}: {title_val}"

    headers = [ws.cell(row=2, column=c).value for c in range(1, 13)]
    print(f"Headers (Row 2): {headers}")
    assert headers == expected_headers, f"Headers mismatch in {s_name}! Expected: {expected_headers}, got: {headers}"

    print(f"First 3 data rows in {s_name}:")
    for r in range(4, min(7, ws.max_row + 1)):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 13)]
        print(f"Row {r}: {row_vals}")

ws_all = wb["全集計ペアリスト"]
ws_excl = wb["除外済みペアリスト"]
print(f"\nRow comparison: 全集計ペアリスト rows = {ws_all.max_row}, 除外済みペアリスト rows = {ws_excl.max_row}")
assert ws_all.max_row >= ws_excl.max_row, "全集計ペアリスト should have >= rows than 除外済みペアリスト"

wb.close()
print("\nAll validations PASSED successfully!")
