import pandas as pd
import os
from datetime import datetime, timedelta
import importlib.util
from importlib.machinery import ModuleSpec
from typing import Optional

root = os.path.dirname(os.path.dirname(__file__))
input_path = os.path.join(root, 'test_input.xlsx')

# Create a simple, valid DataFrame with expected columns
today = datetime.today().date()
df = pd.DataFrame({
    '出欠': ['出席', '出席'],
    '担当講師名': ['講師A', '講師A'],
    '担当講師N0': [1, 1],
    '教室': ['教室1', '教室1'],
    '日付': [today, today + timedelta(days=7)],
    'ラップ1（分子）': [1, 0],
    '確認1（分子）': [None, 1],
    '宿題名': ['HW1', 'HW2'],
    '宿題開始ページ': [1, 5],
    '宿題終了ページ': [2, 6],
})

print('writing', input_path)
df.to_excel(input_path, index=False)

# Load module from src
spec: Optional[ModuleSpec] = importlib.util.spec_from_file_location(
    'msl_v5', os.path.join(root, 'src', 'MSLdata_check_v5.py')
)
if spec is None:
    raise ImportError(f"cannot create module spec for {os.path.join(root, 'src', 'MSLdata_check_v5.py')}")
mod = importlib.util.module_from_spec(spec)
loader = spec.loader
if loader is None:
    raise ImportError("module spec.loader is None; cannot load module")
loader.exec_module(mod)

try:
    print('calling process_attendance_data')
    res = mod.process_attendance_data(input_path)
    print('process returned, output file:', res[4])
    print('calling format_excel_fast')
    mod.format_excel_fast(*res[:9])
    print('format done')
except Exception as e:
    import traceback
    traceback.print_exc()
    print('error:', e)
