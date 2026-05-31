import importlib
import time
import os
import pandas as pd
import numpy as np

SAMPLE = "sample.xlsx"

def make_sample(path: str, rows: int = 2000):
    if os.path.exists(path):
        print(f"sample exists: {path}")
        return
    print(f"generating sample {path} ({rows} rows)")
    dates = pd.date_range(end=pd.Timestamp.today(), periods=30).to_numpy(dtype=object)
    data = []
    for i in range(rows):
        d = np.random.choice(dates)
        attendance = np.random.choice(np.array(["出席", "欠席"], dtype=object), p=[0.9, 0.1])
        teacher = f"講師{np.random.randint(1,50)}"
        teacher_no = np.random.randint(1,200)
        room = f"教室{np.random.randint(1,20)}"
        # hw blocks: name, start, end (repeat 5)
        row = {
            "日付": d.date(),
            "出欠": attendance,
            "担当講師名": teacher,
            "担当講師N0": teacher_no,
            "教室": room,
        }
        for k in range(1,6):
            row[f"宿題名_{k}"] = np.random.choice(np.array(["単語帳","問題集","練習", "Leap"], dtype=object), p=[0.4,0.3,0.25,0.05])
            s = np.random.randint(1,30)
            e = s + np.random.randint(0,3)
            row[f"宿題開始ページ_{k}"] = s
            row[f"宿題終了ページ_{k}"] = e
        # lap and check columns
        for k in range(1,6):
            row[f"ラップ{k}（分子）"] = np.random.choice(np.array([1,0,None], dtype=object), p=[0.1,0.85,0.05])
            row[f"確認{k}（分子）"] = np.random.choice(np.array([1,0,None], dtype=object), p=[0.05,0.9,0.05])
        data.append(row)

    df = pd.DataFrame(data)
    df.to_excel(path, index=False)
    print("sample generated")

def bench_module(mod_name: str, sample_path: str, out_prefix: str):
    print(f"\n--- BENCH {mod_name} ---")
    mod = importlib.import_module(mod_name)
    t0 = time.perf_counter()
    res = mod.process_attendance_data(sample_path)
    t1 = time.perf_counter()
    print(f"process_attendance_data: {t1-t0:.2f}s")
    # write out
    out_file = os.path.abspath(f"{out_prefix}_{mod_name}.xlsx")
    t2 = time.perf_counter()
    # Try to safely call format function by replacing any internal output filename with our out_file
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
                # append out_file if no path found
                args.append(out_file)

            try:
                mod.format_excel_fast(*args)
            except TypeError:
                # fallback: try trimming args to match function signature
                import inspect

                sig = inspect.signature(mod.format_excel_fast)
                params = len(sig.parameters)
                mod.format_excel_fast(*args[:params])
        else:
            # No format function: try to save any dataframe-like objects from res
            saved = False
            if isinstance(res, (list, tuple)):
                for i, v in enumerate(res):
                    try:
                        # Polars DataFrame
                        import polars as pl

                        if isinstance(v, pl.DataFrame):
                            v.to_pandas().to_excel(out_file, index=False)
                            saved = True
                            break
                    except Exception:
                        pass
            if not saved:
                # last resort: touch the file
                pd.DataFrame().to_excel(out_file)
    except Exception:
        raise
    t3 = time.perf_counter()
    print(f"format_excel_fast (write): {t3-t2:.2f}s")
    return (t1-t0, t3-t2, out_file)

def main():
    make_sample(SAMPLE, rows=2000)
    modules = ["MSLdata_check_v5"]
    results = {}
    for m in modules:
        try:
            res = bench_module(m, SAMPLE, "bench_output")
            results[m] = res
        except Exception as e:
            print(f"error bench {m}: {e}")

    print("\n=== SUMMARY ===")
    for m, (t_proc, t_write, of) in results.items():
        print(f"{m}: process {t_proc:.2f}s, write {t_write:.2f}s -> {of}")

if __name__ == "__main__":
    main()
