import importlib
import time
import os
import sys
import pandas as pd
import numpy as np

# 定数設定
SAMPLE = "sample.xlsx"
ROWS = 50000
DATE_RANGE = 90
ROOMS = 26
TEACHERS = 150

def make_sample(path: str, rows: int = ROWS):
    if os.path.exists(path):
        os.remove(path)

    print(f"generating sample {path} ({ROWS} rows)")
    dates = pd.date_range(end=pd.Timestamp.today(), periods=DATE_RANGE)
    data = []

    subject_map = {
        "小学": ["国語", "英語", "算数", "理科", "社会"],
        "中学": ["国語", "英語", "数学", "理科", "社会"],
        "高校": [
            "国語", "英語", "数学", "物理", "算数", "現代文", "化学", "理科",
            "古文", "日本史", "社会", "公共", "生物", "政治経済", "小論文", "歴史", "世界史"
        ]
    }

    for i in range(rows):
        d = np.random.choice(dates)
        date_str = pd.Timestamp(d).strftime("%Y-%m-%d")

        attendance = np.random.choice(np.array(["出席", "欠席"], dtype=object), p=[0.9, 0.1])
        teacher = f"講師{np.random.randint(1, TEACHERS+1)}"
        teacher_no = np.random.randint(1,10000)
        room = f"教室{chr(64 + np.random.randint(1, ROOMS+1))}"

        # 全学年（小1〜高卒）の抽選
        grade = np.random.choice(["小1", "小2", "小3", "小4", "小5", "小6", "中1", "中2", "中3", "高1", "高2", "高3", "高卒"])

        # 学年に応じて紐づく科目の抽選プールを動的に切り替え
        if grade.startswith("小"):
            subj_pool = subject_map["小学"]
        elif grade.startswith("中"):
            subj_pool = subject_map["中学"]
        else:
            subj_pool = subject_map["高校"]

        subject = np.random.choice(subj_pool)

        row = {
            "日付": date_str,
            "出欠": attendance,
            "担当講師名": teacher,
            "担当講師N0": teacher_no,
            "教室": room,
            "学年": grade,
            "科目": subject,
        }
        for k in range(1,6):
            row[f"宿題名_{k}"] = np.random.choice(np.array(["単語帳","問題集","練習", "Leap"], dtype=object), p=[0.4,0.3,0.25,0.05])
            s = np.random.randint(1,30)
            e = s + np.random.randint(0,3)
            row[f"宿題開始ページ_{k}"] = s
            row[f"宿題終了ページ_{k}"] = e
        for k in range(1,6):
            row[f"ラップ{k}（分子）"] = np.random.choice(np.array([1,0,None], dtype=object), p=[0.1,0.85,0.05])
            row[f"確認{k}（分子）"] = np.random.choice(np.array([1,0,None], dtype=object), p=[0.05,0.9,0.05])
        data.append(row)

    df = pd.DataFrame(data)
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    print("sample generated")

def bench_module(mod_name: str, sample_path: str, out_prefix: str):
    print(f"\n--- BENCH {mod_name} ---")
    mod = importlib.import_module(mod_name)
    t0 = time.perf_counter()
    res = mod.process_attendance_data(sample_path)
    t1 = time.perf_counter()
    print(f"process_attendance_data: {t1-t0:.2f}s")
    out_file = os.path.abspath(f"{out_prefix}_{mod_name}.xlsx")
    t2 = time.perf_counter()
    try:
        if hasattr(mod, "format_excel_fast"):
            args = list(res) if isinstance(res, (list, tuple)) else [res]
            replaced = False
            for i, v in enumerate(args):
                if isinstance(v, str) and v.lower().endswith(".xlsx"):
                    args[i] = out_file
                    replaced = True
                    break
            if not replaced:
                args.append(out_file)

            try:
                mod.format_excel_fast(*args)
            except TypeError:
                import inspect

                sig = inspect.signature(mod.format_excel_fast)
                params = len(sig.parameters)
                mod.format_excel_fast(*args[:params])
        else:
            saved = False
            if isinstance(res, (list, tuple)):
                for i, v in enumerate(res):
                    try:
                        import polars as pl

                        if isinstance(v, pl.DataFrame):
                            v.to_pandas().to_excel(out_file, index=False)
                            saved = True
                            break
                    except Exception:
                        pass
            if not saved:
                pd.DataFrame().to_excel(out_file)
    except Exception:
        raise
    t3 = time.perf_counter()
    print(f"format_excel_fast (write): {t3-t2:.2f}s")
    return (t1-t0, t3-t2, out_file)

def main():
    script_path = os.path.abspath(__file__)
    bench_dir = os.path.dirname(script_path)
    project_root = os.path.dirname(bench_dir)
    src_dir = os.path.join(project_root, "src")

    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)

    os.chdir(bench_dir)

    make_sample(SAMPLE, rows=ROWS)
    versions = [1, 2, 3, 4, 5, 6, 7]
    results = {}

    for v in versions:
        mod_name = f"MSLdata_check_v{v}"
        try:
            res = bench_module(mod_name, SAMPLE, "bench_output")
            results[mod_name] = res
        except Exception as e:
            print(f"error bench {mod_name}: {e}")
            import traceback
            traceback.print_exc()

    print("\n=== SUMMARY ===")
    for mod_name, metrics in results.items():
        if metrics is not None:
            t_proc, t_write, of = metrics
            print(f"{mod_name}: process {t_proc:.2f}s, write {t_write:.2f}s -> {of}")

if __name__ == "__main__":
    main()