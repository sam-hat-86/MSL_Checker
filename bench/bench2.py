import os
import sys
import time
import importlib
import inspect
import numpy as np
import xlsxwriter

MIN_VERSION=11
MAX_VERSION=13

def print_progress(current,total,width=40):
    ratio=current/total
    filled=int(width*ratio)
    bar="█"*filled+"-"*(width-filled)
    print(f"\r[{bar}] {ratio*100:6.2f}% ({current:,}/{total:,})",end="",flush=True)

def make_sample(path,rows):
    print("高速サンプル生成開始")
    wb=xlsxwriter.Workbook(path,{"constant_memory":True})
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

    dates=np.datetime64("2026-01-01")+rng.integers(0,30,rows).astype("timedelta64[D]")
    attendance=rng.choice(["出席","欠席"],rows,p=[0.9,0.1])
    teachers=np.char.add("講師",rng.integers(1,16,rows).astype(str))
    teacher_no=rng.integers(1,10000,rows)
    rooms=np.char.add("教室",rng.choice(list("ABCDEFGHIJKLMNO"),rows))
    grades=rng.choice(["小1","小2","小3","中1","中2","中3","高1","高2","高3"],rows)
    subjects=rng.choice(["国語","英語","数学","理科","社会"],rows)

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
        cols.append(rng.choice(hw_names,rows))
        s=rng.integers(1,30,rows)
        cols.append(s)
        cols.append(s+rng.integers(0,3,rows))

    vals=np.array([1,0,None],dtype=object)
    for _ in range(5):
        cols.append(rng.choice(vals,rows,p=[0.6,0.3,0.1]))
        cols.append(rng.choice(vals,rows,p=[0.7,0.2,0.1]))

    for c,data in enumerate(cols):
        ws.write_column(1,c,data.tolist())
        print_progress(c+1,len(cols))

    print()
    wb.close()
    print("サンプル生成完了")

if __name__=="__main__":
    make_sample("MSLサンプル.xlsx",100000)
