import importlib
import time
import os
import sys
import pandas as pd
import numpy as np
import xlsxwriter

MIN_VERSION=12
MAX_VERSION=13

def call_set():
    global mk_sample, VERSIONS

    mk_sample=False

    VERSIONS=[str(v) for v in range(MIN_VERSION,MAX_VERSION+1)]

    # set_CONSTS(10000, 30, 5, 10)
    # set_CONSTS(20000, 30, 5, 10)
    # set_CONSTS(50000, 45, 10, 15)
    # set_CONSTS(100000, 45, 10, 15)
    set_CONSTS(500000, 30, 15, 20)

def set_CONSTS(R,D,C,T):
    global ROWS,DATE_RANGE,ROOMS,TEACHERS,SAMPLE_PATH
    ROWS=R
    DATE_RANGE=D
    ROOMS=C
    TEACHERS=T
    SAMPLE_PATH = f"s{ROWS}R_{DATE_RANGE}D_{ROOMS}C_{TEACHERS}T.xlsx"
    # SAMPLE_PATH = f"MSLサンプル.xlsx"

def print_progress(current,total,width=40):
    ratio=current/total
    filled=int(width*ratio)
    bar="█"*filled+"-"*(width-filled)
    print(f"\r[{bar}] {ratio*100:6.2f}% ({current:,}/{total:,})",end="",flush=True)

def make_sample():
    print("高速サンプル生成開始")
    wb=xlsxwriter.Workbook(SAMPLE_PATH,{"constant_memory":True})
    ws=wb.add_worksheet()

    headers=[
        "日付","出欠","担当講師名","担当講師N0",
        "教室","学年","科目"
    ]
    for i in range(1,6):
        headers+=[
            f"宿題名_{i}",
            f"宿題開始ページ_{i}",
            f"宿題終了ページ_{i}"
        ]
    for i in range(1,6):
        headers.append(f"ラップ{i}（分子）")
        headers.append(f"確認{i}（分子）")

    for c,h in enumerate(headers):
        ws.write(0,c,h)

    rng=np.random.default_rng()

    dates=np.datetime64("2026-01-01")+rng.integers(0,30,ROWS).astype("timedelta64[D]")
    attendance=rng.choice(["出席","欠席"],ROWS,p=[0.9,0.1])
    teachers=np.char.add("講師",rng.integers(1,16,ROWS).astype(str))
    teacher_no=rng.integers(1,10000,ROWS)
    rooms=np.char.add("教室",rng.choice(list("ABCDEFGHIJKLMNO"),ROWS))
    grades=rng.choice(["小1","小2","小3","中1","中2","中3","高1","高2","高3"],ROWS)
    subjects=rng.choice(["国語","英語","数学","理科","社会"],ROWS)

    cols=[
        dates.astype(str),
        attendance,
        teachers,
        teacher_no,
        rooms,
        grades,
        subjects
    ]

    hw_names=["単語帳","問題集","Lodestar","BUILDER","Leap"]

    for _ in range(5):
        cols.append(rng.choice(hw_names,ROWS))
        s=rng.integers(1,30,ROWS)
        cols.append(s)
        cols.append(s+rng.integers(0,3,ROWS))

    vals=np.array([1,0,None],dtype=object)
    for _ in range(5):
        cols.append(rng.choice(vals,ROWS,p=[0.6,0.3,0.1]))
        cols.append(rng.choice(vals,ROWS,p=[0.7,0.2,0.1]))

    for c,data in enumerate(cols):
        ws.write_column(1,c,data.tolist())
        print_progress(c+1,len(cols))

    print()
    wb.close()
    print("サンプル生成完了")

def bench_module(mod_name,sample_path,out_prefix):
    import inspect

    print(f"\n--- ベンチマーク {mod_name} ---")

    mod=importlib.import_module(mod_name)

    sig=inspect.signature(mod.process_attendance_data)
    kwargs={}

    if "start_from_saturday" in sig.parameters:
        kwargs["start_from_saturday"]=False

    t0=time.perf_counter()
    res=mod.process_attendance_data(sample_path,**kwargs)
    t1=time.perf_counter()

    print(f"データ処理: {t1-t0:.2f}s")

    out_file=os.path.abspath(f"{out_prefix}_{mod_name}.xlsx")

    t2=time.perf_counter()

    if hasattr(mod,"format_excel_fast"):
        args=list(res) if isinstance(res,(list,tuple)) else [res]

        replaced=False
        for i,v in enumerate(args):
            if isinstance(v,str) and v.lower().endswith(".xlsx"):
                args[i]=out_file
                replaced=True
                break

        if not replaced:
            sig=inspect.signature(mod.format_excel_fast)

            for i,(name,param) in enumerate(sig.parameters.items()):
                if "output" in name.lower() or "file" in name.lower():
                    while len(args)<=i:
                        args.append(None)
                    args[i]=out_file
                    replaced=True
                    break

        if not replaced:
            args.append(out_file)

        sig=inspect.signature(mod.format_excel_fast)
        params=list(sig.parameters.values())

        while len(args)<len(params):
            p=params[len(args)]

            if p.default is not inspect._empty:
                args.append(p.default)
            elif "output" in p.name.lower() or "file" in p.name.lower():
                args.append(out_file)
            elif p.annotation is bool or "flag" in p.name.lower():
                args.append(False)
            else:
                args.append(None)

        mod.format_excel_fast(*args[:len(params)])

    else:
        if isinstance(res,pd.DataFrame):
            res.to_excel(out_file,index=False)
        else:
            pd.DataFrame().to_excel(out_file,index=False)

    t3=time.perf_counter()

    print(f"書き込み: {t3-t2:.2f}s")

    file_size_kb=os.path.getsize(out_file)/1024 if os.path.exists(out_file) else 0

    return(
        t1-t0,
        t3-t2,
        out_file,
        file_size_kb
    )

def main():
    call_set()

    script_dir=os.path.dirname(os.path.abspath(__file__))
    project_root=os.path.dirname(script_dir)
    src_dir=os.path.join(project_root,"src")

    if src_dir not in sys.path:
        sys.path.insert(0,src_dir)

    os.chdir(script_dir)

    global SAMPLE_PATH
    SAMPLE_PATH=os.path.join(script_dir,os.path.basename(SAMPLE_PATH))

    if not os.path.exists(SAMPLE_PATH):
        print(f"サンプルファイルが存在しないため生成します。")
        make_sample()
    else:
        print(f"サンプル: {os.path.basename(SAMPLE_PATH)}")

    results={}

    print(f"対象バージョン: v{MIN_VERSION}～v{MAX_VERSION}")

    for v in VERSIONS:
        mod_name=f"MSLdata_check_v{v}"

        try:
            results[mod_name]=bench_module(mod_name,SAMPLE_PATH,"bench_op")
        except ModuleNotFoundError:
            print(f"[SKIP] {mod_name} が存在しません。")
        except Exception as e:
            print(f"[ERROR] {mod_name}: {e}")
            import traceback
            traceback.print_exc()

    print("\n=== SUMMARY ===")

    for mod_name,(proc,write,_,size) in results.items():
        total=proc+write
        print(f"{mod_name}: calc {proc:.2f}s + write {write:.2f}s = {total:.2f}s -> {size:.2f} KB")

if __name__=="__main__":
    main()