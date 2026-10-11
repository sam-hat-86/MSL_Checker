import os
import sys
import glob
import re
import sqlite3
import xlsxwriter
import fastexcel
from datetime import datetime, timedelta

def test_aggregation():
    sample_files = glob.glob('bench/*サンプル*.xlsx')
    if not sample_files:
        sample_files = [x for x in glob.glob('bench/*.xlsx') if 'bench_op' not in x]
    input_file = sample_files[0]
    print(f"Testing with file: {input_file}")

    excel_reader = fastexcel.read_excel(input_file)
    sheet = excel_reader.load_sheet(0, header_row=None)
    arrow_table = sheet.to_arrow()
    raw_data = arrow_table.to_pylist()
    header_raw = list(raw_data[0].values())
    data_rows = [list(r.values()) for r in raw_data[1:]]

    cols = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(header_raw)]
    col_to_idx = {c: idx for idx, c in enumerate(cols)}

    c_att_idx = col_to_idx.get("出欠")
    c_room_idx = col_to_idx.get("教室")
    c_tname_idx = col_to_idx.get("担当講師名")
    c_tno_idx = col_to_idx.get("担当講師N0")
    c_sname_idx = col_to_idx.get("生徒名")
    c_grade_idx = col_to_idx.get("学年")
    c_subj_idx = col_to_idx.get("科目")

    lap_col_names = [c for c in cols if "ラップ" in c and "（分子）" in c]
    lap_indices = [col_to_idx[c] for c in lap_col_names if c in col_to_idx]

    re_fs = re.compile(r"(ＦＳ|FS|ｆｓ|fs)")
    re_teacher_mark = re.compile(r"[Ａ-Ｚ_]")
    re_excl_grade = re.compile(r"(高3|高３|高卒)")
    re_excl_subj = re.compile(r"(国語|小論文)")

    records = []
    for row in data_rows:
        if not row: continue
        att_val = row[c_att_idx] if c_att_idx is not None else ""
        if att_val != "出席" and str(att_val).strip() != "出席": continue

        cls = re_fs.sub("", str(row[c_room_idx] or "")).strip()
        name = re_teacher_mark.sub("", str(row[c_tname_idx] or "")).strip()
        no = str(row[c_tno_idx] or "").strip()
        sname = str(row[c_sname_idx] or "").strip()
        grade = str(row[c_grade_idx] or "").strip()
        subj = str(row[c_subj_idx] or "").strip()

        excl_grade = 1 if re_excl_grade.search(grade) else 0
        excl_subj = 1 if re_excl_subj.search(subj) else 0

        has_lap = 1 if any(row[idx] is not None and str(row[idx]).strip() and str(row[idx]).strip().lower() != "none" for idx in lap_indices) else 0
        records.append((no, name, cls, sname, grade, subj, has_lap, excl_grade, excl_subj))

    print(f"Total attended records: {len(records)}")

    conn = sqlite3.connect(":memory:")
    cur = conn.cursor()
    cur.execute("CREATE TABLE att (no TEXT, name TEXT, cls TEXT, sname TEXT, grade TEXT, subj TEXT, has_lap INTEGER, excl_grade INTEGER, excl_subj INTEGER)")
    cur.executemany("INSERT INTO att VALUES (?,?,?,?,?,?,?,?,?)", records)

    # 1. 講師全体Lap率
    cur.execute("""
        SELECT cls, no, COUNT(*) as tot, SUM(has_lap) as lap
        FROM att
        WHERE excl_grade=0 AND excl_subj=0
        GROUP BY cls, no
    """)
    t_stats = {(r[0], r[1]): (r[3] / r[2] if r[2] > 0 else 0.0) for r in cur.fetchall()}

    # 2. 生徒全体Lap率
    cur.execute("""
        SELECT cls, sname, COUNT(*) as tot, SUM(has_lap) as lap
        FROM att
        WHERE excl_grade=0 AND excl_subj=0
        GROUP BY cls, sname
    """)
    s_stats = {(r[0], r[1]): (r[3] / r[2] if r[2] > 0 else 0.0) for r in cur.fetchall()}

    # 3. ペア別集計
    cur.execute("""
        SELECT cls, no, name, sname, grade, subj, COUNT(*) as lessons, SUM(has_lap) as lap_cnt
        FROM att
        WHERE excl_grade=0 AND excl_subj=0
        GROUP BY cls, no, name, sname, grade, subj
    """)
    pair_rows = cur.fetchall()
    print(f"Total distinct excluded pairs: {len(pair_rows)}")

    pairs = []
    for r in pair_rows:
        cls, no, name, sname, grade, subj, lessons, lap_cnt = r
        unlap_cnt = lessons - lap_cnt
        lap_rate = (lap_cnt / lessons) if lessons > 0 else 0.0
        t_rate = t_stats.get((cls, no), 0.0)
        s_rate = s_stats.get((cls, sname), 0.0)
        pairs.append({
            "cls": cls,
            "no": no,
            "name": name,
            "sname": sname,
            "grade": grade,
            "subj": subj,
            "lessons": lessons,
            "lap_cnt": lap_cnt,
            "unlap_cnt": unlap_cnt,
            "lap_rate": lap_rate,
            "t_all_rate": t_rate,
            "s_all_rate": s_rate,
        })

    pairs.sort(key=lambda x: (x["cls"], x["lap_rate"], -x["unlap_cnt"], -x["lessons"]))

    print("Sample top 5 worst pairs:")
    for p in pairs[:5]:
        print(f"[{p['cls']}] 講師: {p['name']}({p['no']}), 生徒: {p['sname']}({p['grade']}/{p['subj']}) | 授業:{p['lessons']}, Lap:{p['lap_cnt']}, 未実:{p['unlap_cnt']}, ペア率:{p['lap_rate']:.1%}, 講師率:{p['t_all_rate']:.1%}, 生徒率:{p['s_all_rate']:.1%}")

if __name__ == '__main__':
    test_aggregation()

