import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter
import os
from datetime import datetime, timedelta
import re
import unicodedata
import sys
import ctypes
import tkinter.font as tkfont
import traceback
import time
import gc
import pyarrow
import fastexcel

# ==========================================
# 【設定セクション】 アプリケーション全般・Excel出力の統合管理
# ここを変更するだけで、全体の動作やデザインを一括変更できます
# ==========================================
class AppConfig:
    # --- 1. 除外キーワード・フォント設定 ---
    EXCLUDED_HW_TERMS = [
        "タンゴ", "単語", "シスタン", "シス単", "ターゲット", "リープ",
        "leap", "ゲットスルー", "パスタン", "パス単", "マドンナ"
    ]
    FONT_PREFERRED = ["Noto Sans JP", "Arial Unicode MS", "Arial"]

    # --- 2. 出力ファイルの列名（表示名）設定 ---
    COL_NO = "No"
    COL_NAME = "氏名"
    COL_CLASSROOM = "教室"
    COL_TOTAL = "総授業"

    COL_LESSON = "授業"
    COL_HW_AVG = "平均HW"
    COL_LAP_CNT = "Lap回数"
    COL_LAP_RATE = "Lap％"
    COL_TEST_CNT = "テスト回数"
    COL_TEST_RATE = "テスト％"

    # --- 3. アラート（条件付き書式）のしきい値設定 ---
    ALERT_HW_UNDER = 6.0    # 宿題平均がこの値未満なら警告
    ALERT_RATE_UNDER = 0.7  # ラップ率・テスト率がこの値(70%)未満なら警告

    # --- 4. Excel ページ・印刷設定 ---
    PAPER_SIZE = 9          # 9 = A4用紙, 8 = A3用紙, 11 = A5用紙
    IS_LANDSCAPE = False    # True: 横向き印刷, False: 縦向き印刷
    MARGIN_LR = 0.5 / 2.54  # 左右のマージン（インチ単位。0.5cm / 2.54）
    MARGIN_TB = 0.75 / 2.54 # 上下のマージン（インチ単位。0.75cm / 2.54）

    # --- 5. テーマカラー設定（HEXコード） ---
    # シートタブの色
    COLOR_TAB_ALL = "#4F81BD"     # 全集計シート
    COLOR_TAB_EXCL = "#4B6F44"    # 除外済み集計シート

    # 見出し・ヘッダーの色
    COLOR_HEADER = "#D3D3D3"      # 列ヘッダー背景色 (グレー)
    COLOR_WEEK_TITLE = "#E8F6F3"  # 週ごとの大見出し背景色 (薄い青緑)
    COLOR_SYS_TEXT = "#A0A0A0"    # 2行目のシステム列番号の文字色

    # データセルの背景色
    COLOR_CLASS_BG = "#F0E68C"    # 授業数列の背景色 (カーキ)
    COLOR_TOTAL_BG = "#E6E6FA"    # 総授業数列の背景色 (ラベンダー)

    # アラート（条件付き書式）の色
    COLOR_ALERT_HW = "#FFA500"    # 宿題平均の警告色 (オレンジ)
    COLOR_ALERT_RATE = "#B22222"  # テスト率等の警告色 (濃い赤/Firebrick)
    COLOR_ALERT_TEXT = "#FFFFFF"  # テスト率等の警告時の文字色 (白)
# ==========================================


def progress(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{ts}] {msg}")

def dismiss_splash() -> None:
    try:
        if sys.platform == "win32":
            import ctypes
            if hasattr(ctypes, "windll") and hasattr(ctypes.windll, "user32"):
                ctypes.windll.user32.PostMessageW(0xFFFF, 0x0400 + 5, 0, 0)
    except Exception:
        pass

def select_font(preferred: list[str] | None = None) -> str:
    prefs = preferred or AppConfig.FONT_PREFERRED
    try:
        root_tmp = tk.Tk()
        root_tmp.withdraw()
        fams = set(tkfont.families())
        root_tmp.destroy()
        for f in prefs:
            if f in fams: return f
    except Exception:
        pass
    for f in prefs: return f
    return "Calibri"

def normalize_text_for_matching(s: object) -> str:
    if s is None: return ""
    t = str(s)
    t = unicodedata.normalize("NFKC", t).lower()
    t = re.sub(r"\s+", "", t)
    t = "".join(chr(ord(c) + 96) if "ぁ" <= c <= "ん" else c for c in t)
    return t

def enable_windows_dpi_awareness() -> None:
    if sys.platform != "win32": return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

def configure_tk_scaling(root: tk.Tk) -> None:
    try:
        scale = float(root.tk.call("tk", "scaling"))
        root.tk.call("tk", "scaling", scale)
    except Exception:
        pass

def get_jp_weekday(dt_obj):
    return ["月", "火", "水", "木", "金", "土", "日"][dt_obj.weekday()]

def calculate_visual_width(text: str) -> float:
    if not text: return 0.0
    width = 0.0
    for char in text:
        status = unicodedata.east_asian_width(char)
        if status in ("W", "F", "A"): width += 2.0
        else: width += 1.0
    return width

def process_attendance_data(input_file: str, start_from_saturday: bool):
    progress("ファイル読み込み開始...")

    header_raw = []
    data_rows = []

    progress("fastexcel読み込み")
    excel_reader = fastexcel.read_excel(input_file)
    sheet = excel_reader.load_sheet(0, header_row=None)

    arrow_table = sheet.to_arrow()
    raw_data = arrow_table.to_pylist()

    if raw_data:
        header_raw = list(raw_data[0].values())
        data_rows = [list(r.values()) for r in raw_data[1:]]

    if not header_raw:
        return [], [], {}, {}, "", "", "", [], [], {}, []

    headers = []
    hw_base_idx = -1
    for i, h in enumerate(header_raw):
        h_str = str(h).strip() if h is not None else ""
        headers.append(h_str)
        if "宿題名" in h_str and hw_base_idx == -1:
            hw_base_idx = i

    if hw_base_idx != -1:
        pattern = ["宿題名", "宿題開始ページ", "宿題終了ページ"]
        k, pattern_idx = 2, 0
        for i in range(hw_base_idx + 1, len(headers)):
            if headers[i].strip() == "":
                headers[i] = f"{pattern[pattern_idx]}_{k}"
                pattern_idx += 1
                if pattern_idx >= 3:
                    pattern_idx, k = 0, k + 1

    unique_headers = []
    seen = set()
    for h in headers:
        h = h.strip()
        if h == "": h = "無名列"
        orig, c = h, 1
        while h in seen:
            h = f"{orig}_{c}"
            c += 1
        seen.add(h)
        unique_headers.append(h)

    cols = unique_headers

    possible_date_cols = ["日付", "授業日"]
    date_col = next((dc for dc in possible_date_cols if dc in cols), None)
    lap_col_names = [c for c in cols if "ラップ" in c and "（分子）" in c]
    kaku_col_names = [c for c in cols if "確認" in c and "（分子）" in c]

    hw_start_idx = next((i for i, c in enumerate(cols) if c and "宿題名" in c), len(cols))
    hw_cols_groups = []
    for i in range(hw_start_idx, len(cols), 3):
        if i + 2 < len(cols):
            hw_cols_groups.append((cols[i], cols[i+1], cols[i+2]))

    progress("出席データの抽出およびデータベース用レコード構築中...")
    records = []
    week_period_by_week_number = {}
    min_dt_overall = None
    max_dt_overall = None

    # 正規表現オブジェクトをループ外でコンパイルして徹底的に高速化
    re_fs = re.compile(r"(ＦＳ|FS|ｆｓ|fs)")
    re_teacher_mark = re.compile(r"[Ａ-Ｚ_]")
    re_digits = re.compile(r"(\d+)")
    re_excl_grade = re.compile(r"(高3|高３|高卒)")
    re_excl_subj = re.compile(r"(国語|小論文)")

    # 高速にマッピングするためにインデックスリストを作成
    col_to_idx = {c: idx for idx, c in enumerate(cols)}
    c_att_idx = col_to_idx.get("出欠")
    c_room_idx = col_to_idx.get("教室")
    c_tname_idx = col_to_idx.get("担当講師名")
    c_tno_idx = col_to_idx.get("担当講師N0")
    c_date_idx = col_to_idx.get(date_col) if date_col else None
    c_grade_idx = col_to_idx.get("学年")
    c_subj_idx = col_to_idx.get("科目")

    lap_indices = [col_to_idx[c] for c in lap_col_names if c in col_to_idx]
    kaku_indices = [col_to_idx[c] for c in kaku_col_names if c in col_to_idx]

    hw_groups_indices = []
    for n_c, s_c, e_c in hw_cols_groups:
        if n_c in col_to_idx and s_c in col_to_idx and e_c in col_to_idx:
            hw_groups_indices.append((col_to_idx[n_c], col_to_idx[s_c], col_to_idx[e_c]))

    for row in data_rows:
        if not row: continue

        # 出欠チェックの高速化
        att_val = row[c_att_idx] if c_att_idx is not None else ""
        if att_val != "出席" and str(att_val).strip() != "出席": continue

        cls_val = row[c_room_idx] if c_room_idx is not None else ""
        cls = re_fs.sub("", str(cls_val)).strip()

        name_val = row[c_tname_idx] if c_tname_idx is not None else ""
        name = re_teacher_mark.sub("", str(name_val)).strip()

        no = str(row[c_tno_idx]).strip() if c_tno_idx is not None else ""

        dt_val = row[c_date_idx] if c_date_idx is not None else None
        week_num = None
        dt_obj = None

        if isinstance(dt_val, datetime): dt_obj = dt_val
        elif isinstance(dt_val, str):
            try: dt_obj = datetime.strptime(dt_val, "%Y/%m/%d")
            except:
                try: dt_obj = datetime.strptime(dt_val, "%Y-%m-%d")
                except: pass

        if dt_obj:
            if not min_dt_overall or dt_obj < min_dt_overall: min_dt_overall = dt_obj
            if not max_dt_overall or dt_obj > max_dt_overall: max_dt_overall = dt_obj

            if start_from_saturday:
                offset_dt = dt_obj + timedelta(days=2)
                week_num = str(offset_dt.isocalendar()[1])
            else:
                week_num = str(dt_obj.isocalendar()[1])

            if week_num not in week_period_by_week_number:
                week_period_by_week_number[week_num] = {"min": dt_obj, "max": dt_obj}
            else:
                if dt_obj < week_period_by_week_number[week_num]["min"]: week_period_by_week_number[week_num]["min"] = dt_obj
                if dt_obj > week_period_by_week_number[week_num]["max"]: week_period_by_week_number[week_num]["max"] = dt_obj

        hw_pages = 0
        for n_idx, s_idx, e_idx in hw_groups_indices:
            hw_name = str(row[n_idx]).strip() if row[n_idx] is not None else ""
            if not hw_name or hw_name.lower() == "none": continue

            norm_name = normalize_text_for_matching(hw_name)
            if any(t in norm_name for t in AppConfig.EXCLUDED_HW_TERMS if t): continue

            try:
                s_m = re_digits.search(str(row[s_idx]))
                e_m = re_digits.search(str(row[e_idx]))
                if s_m and e_m:
                    pages = int(e_m.group(1)) - int(s_m.group(1)) + 1
                    if pages < 0: pages = 0
                    elif pages > 30: pages = 30
                    hw_pages += pages
            except Exception: pass

        has_lap = 1 if any(row[idx] is not None and str(row[idx]).strip() and str(row[idx]).strip().lower() != "none" for idx in lap_indices) else 0
        has_kaku = 1 if any(row[idx] is not None and str(row[idx]).strip() and str(row[idx]).strip().lower() != "none" for idx in kaku_indices) else 0
        has_test = 1 if (has_lap or has_kaku) else 0

        grade = str(row[c_grade_idx]) if c_grade_idx is not None else ""
        subj = str(row[c_subj_idx]) if c_subj_idx is not None else ""
        excl_grade = 1 if re_excl_grade.search(grade) else 0
        excl_subj = 1 if re_excl_subj.search(subj) else 0

        records.append((no, name, cls, week_num, hw_pages, has_lap, has_test, excl_grade, excl_subj))

    if min_dt_overall and max_dt_overall:
        min_date_file = min_dt_overall.strftime("%Y%m%d")
        max_date_file = max_dt_overall.strftime("%Y%m%d")
        min_yy = min_dt_overall.strftime("%y")
        max_yy = max_dt_overall.strftime("%y")
        date_range_str = f"期間:{min_yy}/{min_dt_overall.month}/{min_dt_overall.day}({get_jp_weekday(min_dt_overall)}) ～ {max_yy}/{max_dt_overall.month}/{max_dt_overall.day}({get_jp_weekday(max_dt_overall)})"
    else:
        min_date_file, max_date_file = "不明", "不明"
        date_range_str = "期間:不明"

    sheet_name = "集計結果"
    output_filename = f"集計結果_[{min_date_file}-{max_date_file}]_{datetime.now().strftime('%y%m%d-%H%M%S')}.xlsx"
    output_file = os.path.join(os.path.dirname(input_file), output_filename)

    progress("SQLiteインメモリデータベースによる超高速ピボット集計を実行中...")
    conn = sqlite3.connect(":memory:")
    cur = conn.cursor()
    cur.execute("CREATE TABLE att (no TEXT, name TEXT, cls TEXT, week_num TEXT, hw_pages REAL, has_lap INTEGER, has_test INTEGER, excl_grade INTEGER, excl_subj INTEGER)")
    cur.executemany("INSERT INTO att VALUES (?,?,?,?,?,?,?,?,?)", records)

    distinct_weeks = sorted(list(set(r[3] for r in records if r[3] is not None)))
    lesson_nums = list(range(1, len(distinct_weeks) + 1))

    week_labels_map = {}
    for w in distinct_weeks:
        dts = week_period_by_week_number.get(w)
        if dts:
            md, xd = dts["min"], dts["max"]
            week_labels_map[w] = f"{md.strftime('%Y/%m/%d')}({get_jp_weekday(md)}) ～ {xd.strftime('%Y/%m/%d')}({get_jp_weekday(xd)})"
        else:
            week_labels_map[w] = f"Week {w}"

    # === 出力フォーマットの定義 ===
    export_keys = ["base_no", "base_name", "base_cls", "base_total"]
    for i in lesson_nums:
        export_keys.extend([f"lesson_{i}", f"hw_{i}", f"lap_cnt_{i}", f"lap_rate_{i}", f"test_cnt_{i}", f"test_rate_{i}"])

    # 内部キー -> 表示名(設定セクションから取得) のマッピングを作成
    display_map = {
        "base_no": AppConfig.COL_NO,
        "base_name": AppConfig.COL_NAME,
        "base_cls": AppConfig.COL_CLASSROOM,
        "base_total": AppConfig.COL_TOTAL
    }
    for i in lesson_nums:
        display_map[f"lesson_{i}"] = f"{AppConfig.COL_LESSON}"
        display_map[f"hw_{i}"] = f"{AppConfig.COL_HW_AVG}"
        display_map[f"lap_cnt_{i}"] = f"{AppConfig.COL_LAP_CNT}"
        display_map[f"lap_rate_{i}"] = f"{AppConfig.COL_LAP_RATE}"
        display_map[f"test_cnt_{i}"] = f"{AppConfig.COL_TEST_CNT}"
        display_map[f"test_rate_{i}"] = f"{AppConfig.COL_TEST_RATE}"

    def get_pivot_data(condition=""):
        where = f"WHERE {condition}" if condition else ""
        cols_sql = ["cls", "no", "name", "COUNT(*) as total"]
        for w in distinct_weeks:
            cols_sql.append(f"SUM(CASE WHEN week_num='{w}' THEN 1 ELSE 0 END) as w_cnt_{w}")
            cols_sql.append(f"SUM(CASE WHEN week_num='{w}' THEN hw_pages ELSE 0 END) as w_hw_{w}")
            cols_sql.append(f"SUM(CASE WHEN week_num='{w}' THEN has_lap ELSE 0 END) as w_lap_{w}")
            cols_sql.append(f"SUM(CASE WHEN week_num='{w}' THEN has_test ELSE 0 END) as w_test_{w}")

        sql = f"SELECT {','.join(cols_sql)} FROM att {where} GROUP BY cls, no, name ORDER BY cls, no"
        cur.execute(sql)
        rows = cur.fetchall()
        if not rows: return []

        col_names = [d[0] for d in cur.description]
        res = []
        for r in rows:
            d = dict(zip(col_names, r))
            record = {
                "base_cls": d["cls"],
                "base_no": d["no"],
                "base_name": d["name"],
                "base_total": d["total"]
            }

            for i, w in enumerate(distinct_weeks, start=1):
                cnt = d[f"w_cnt_{w}"]
                record[f"lesson_{i}"] = cnt
                if cnt > 0:
                    record[f"hw_{i}"] = round(d[f"w_hw_{w}"] / cnt, 1)
                    record[f"lap_rate_{i}"] = d[f"w_lap_{w}"] / cnt
                    record[f"test_rate_{i}"] = d[f"w_test_{w}"] / cnt
                else:
                    record[f"hw_{i}"] = None
                    record[f"lap_rate_{i}"] = None
                    record[f"test_rate_{i}"] = None

                record[f"lap_cnt_{i}"] = d[f"w_lap_{w}"]
                record[f"test_cnt_{i}"] = d[f"w_test_{w}"]
                record[f"prod_hw_{i}"] = d[f"w_hw_{w}"] # 積和計算用の退避

            res.append(record)
        return res

    def get_summary_data(condition=""):
        where = f"WHERE {condition}" if condition else ""
        sql = f"SELECT cls, COUNT(*) as cnt FROM att {where} GROUP BY cls"
        cur.execute(sql)
        return {r[0]: r[1] for r in cur.fetchall()}

    records_all = get_pivot_data("")
    records_excl = get_pivot_data("excl_grade=0 AND excl_subj=0")
    summary_all = get_summary_data("")
    summary_excl = get_summary_data("excl_grade=0 AND excl_subj=0")

    week_labels = [week_labels_map[w] for w in distinct_weeks]

    conn.close()
    gc.collect()

    return records_all, records_excl, summary_all, summary_excl, output_file, date_range_str, sheet_name, week_labels, export_keys, display_map, lesson_nums

def show_error_dialog(title: str, message: str, parent: tk.Tk | None = None) -> None:
    dlg = tk.Toplevel(parent) if parent else tk.Toplevel()
    dlg.title(title)
    try: scale = float(parent.tk.call("tk", "scaling")) if parent else float(dlg.tk.call("tk", "scaling"))
    except Exception: scale = 1.0
    base_font = tkfont.nametofont("TkDefaultFont")
    dlg_font = tkfont.Font(family=base_font.cget("family"), size=max(int(base_font.cget("size") * scale), 10))
    dlg.geometry(f"{int(700*scale)}x{int(360*scale)}")
    txt = tk.Text(dlg, wrap="word", font=dlg_font)
    txt.insert("1.0", message)
    txt.configure(state="disabled")
    txt.pack(expand=True, fill="both", padx=int(6 * scale), pady=int(6 * scale))
    frm = tk.Frame(dlg); frm.pack(fill="x", padx=int(6 * scale), pady=int(6 * scale))
    def _copy():
        try: dlg.clipboard_clear(); dlg.clipboard_append(message)
        except Exception: pass
    tk.Button(frm, text="エラーメッセージをコピー", command=_copy, font=dlg_font).pack(side="left")
    tk.Button(frm, text="閉じる", command=dlg.destroy, font=dlg_font).pack(side="right")
    try: dlg.transient(parent); dlg.grab_set(); dlg.wait_window()
    except Exception: pass

# ==========================================
# スタイル設定ロジック (表示名に依存せず、内部キーで色を判定)
# ==========================================
def get_column_style_config(internal_key: str, display_name: str) -> tuple[float, str, str | None]:
    header_visual_width = calculate_visual_width(display_name)

    # 固定列のスタイル判定
    if internal_key == "base_no": return max(header_visual_width + 1.5, 5.0), "white", None
    if internal_key == "base_name": return max(header_visual_width + 1.5, 11.0), "white", None
    if internal_key == "base_cls": return max(header_visual_width + 1.5, 9.0), "white", None
    if internal_key == "base_total": return max(header_visual_width + 2.0, 9.0), "total_bg", None

    # 変動列（週ごと）のスタイル判定
    col_width = max(header_visual_width + 2.0, 8.0)
    if internal_key.startswith("lesson_"): return col_width, "class_bg", None
    if internal_key.startswith("lap_rate_") or internal_key.startswith("test_rate_"): return col_width, "white_pct", "pct"

    return col_width, "white", None

def format_excel_fast(
    records_all,
    records_excl,
    summary_all,
    summary_excl,
    output_file,
    date_range_str,
    sheet_name,
    week_period_labels,
    export_keys,
    display_map,
    lesson_nums
):
    progress("xlsxwriterによる高速レイアウト処理を開始")
    workbook = xlsxwriter.Workbook(output_file, {'constant_memory': True})
    ud_font = select_font()

    # テーマカラーをAppConfigから適用
    fmt_title = workbook.add_format({"font_name": ud_font, "bold": True, "font_size": 11, "align": "left", "valign": "vcenter"})
    fmt_title_center = workbook.add_format({"font_name": ud_font, "bold": True, "font_size": 11, "align": "center", "valign": "vcenter", "bg_color": AppConfig.COLOR_WEEK_TITLE, "border": 1})
    fmt_header = workbook.add_format({"font_name": ud_font, "border": 1, "bg_color": AppConfig.COLOR_HEADER, "bold": True, "align": "center", "valign": "vcenter"})
    fmt_system_sub = workbook.add_format({"font_name": ud_font, "border": 1, "font_size": 8, "font_color": AppConfig.COLOR_SYS_TEXT, "align": "center", "valign": "vcenter"})
    fmt_cond_orange = workbook.add_format({"font_name": ud_font, "bg_color": AppConfig.COLOR_ALERT_HW, "font_color": "#000000"})
    fmt_cond_firebrick = workbook.add_format({"font_name": ud_font, "bg_color": AppConfig.COLOR_ALERT_RATE, "font_color": AppConfig.COLOR_ALERT_TEXT})

    formats_matrix = {
        "base": {
            "white": workbook.add_format({"font_name": ud_font, "border": 1, "valign": "vcenter"}),
            "class_bg": workbook.add_format({"font_name": ud_font, "border": 1, "bg_color": AppConfig.COLOR_CLASS_BG, "valign": "vcenter"}),
            "total_bg": workbook.add_format({"font_name": ud_font, "border": 1, "bg_color": AppConfig.COLOR_TOTAL_BG, "valign": "vcenter"}),
            "white_pct": workbook.add_format({"font_name": ud_font, "border": 1, "num_format": "0.0%", "valign": "vcenter"})
        },
        "avg": {
            "white": workbook.add_format({"font_name": ud_font, "border": 1, "bold": True, "valign": "vcenter"}),
            "class_bg": workbook.add_format({"font_name": ud_font, "border": 1, "bg_color": AppConfig.COLOR_CLASS_BG, "bold": True, "valign": "vcenter"}),
            "total_bg": workbook.add_format({"font_name": ud_font, "border": 1, "bg_color": AppConfig.COLOR_TOTAL_BG, "bold": True, "valign": "vcenter"}),
            "white_pct": workbook.add_format({"font_name": ud_font, "border": 1, "bold": True, "num_format": "0.0%", "valign": "vcenter"})
        }
    }

    def write_sheet(records_list, summary_dict, sheet_title, tab_color):
        progress(f"[{sheet_title}] のデータ合成中...")

        final_col_values: dict[str, list] = {k: [] for k in export_keys}

        if len(records_list) > 0:
            unique_classrooms = sorted(list(set(r.get("base_cls", "") for r in records_list)))

            progress(f"[{sheet_title}] ステップ1: 全体(FS)のAVG集計を計算")
            final_col_values["base_no"].append("0")
            final_col_values["base_name"].append("AVG")
            final_col_values["base_cls"].append("FS")

            total_lessons_fs = sum(r.get("base_total", 0) for r in records_list if isinstance(r.get("base_total"), (int, float)))
            final_col_values["base_total"].append(total_lessons_fs)

            for i in lesson_nums:
                w_lessons = sum(r.get(f"lesson_{i}", 0) for r in records_list if isinstance(r.get(f"lesson_{i}"), (int, float)))
                w_lap = sum(r.get(f"lap_cnt_{i}", 0) for r in records_list if isinstance(r.get(f"lap_cnt_{i}"), (int, float)))
                w_test = sum(r.get(f"test_cnt_{i}", 0) for r in records_list if isinstance(r.get(f"test_cnt_{i}"), (int, float)))
                thw = sum(r.get(f"prod_hw_{i}", 0.0) for r in records_list)

                final_col_values[f"lesson_{i}"].append(w_lessons)
                if w_lessons == 0:
                    final_col_values[f"hw_{i}"].append(None)
                    final_col_values[f"lap_rate_{i}"].append(None)
                    final_col_values[f"test_rate_{i}"].append(None)
                    final_col_values[f"lap_cnt_{i}"].append(0.0)
                    final_col_values[f"test_cnt_{i}"].append(0.0)
                else:
                    final_col_values[f"lap_cnt_{i}"].append(w_lap)
                    final_col_values[f"test_cnt_{i}"].append(w_test)
                    final_col_values[f"lap_rate_{i}"].append(w_lap / w_lessons)
                    final_col_values[f"test_rate_{i}"].append(w_test / w_lessons)
                    final_col_values[f"hw_{i}"].append(round(thw / w_lessons, 1))

            progress(f"[{sheet_title}] ステップ2: 各教室ごとのAVG・個別データマージ中...")
            for target_cls in unique_classrooms:
                cls_records = [r for r in records_list if r.get("base_cls") == target_cls]

                cls_total = summary_dict.get(target_cls, 0)
                if cls_total == 0:
                    cls_total = sum(r.get("base_total", 0) for r in cls_records if isinstance(r.get("base_total"), (int, float)))

                final_col_values["base_no"].append("0")
                final_col_values["base_name"].append("AVG")
                final_col_values["base_cls"].append(target_cls)
                final_col_values["base_total"].append(cls_total)

                for i in lesson_nums:
                    w_lessons_cls = sum(r.get(f"lesson_{i}", 0) for r in cls_records if isinstance(r.get(f"lesson_{i}"), (int, float)))
                    wl_c = sum(r.get(f"lap_cnt_{i}", 0) for r in cls_records if isinstance(r.get(f"lap_cnt_{i}"), (int, float)))
                    wt_c = sum(r.get(f"test_cnt_{i}", 0) for r in cls_records if isinstance(r.get(f"test_cnt_{i}"), (int, float)))
                    thw_c = sum(r.get(f"prod_hw_{i}", 0.0) for r in cls_records)

                    final_col_values[f"lesson_{i}"].append(w_lessons_cls)
                    if w_lessons_cls == 0:
                        final_col_values[f"hw_{i}"].append(None)
                        final_col_values[f"lap_rate_{i}"].append(None)
                        final_col_values[f"test_rate_{i}"].append(None)
                        final_col_values[f"lap_cnt_{i}"].append(0.0)
                        final_col_values[f"test_cnt_{i}"].append(0.0)
                    else:
                        final_col_values[f"lap_cnt_{i}"].append(wl_c)
                        final_col_values[f"test_cnt_{i}"].append(wt_c)
                        final_col_values[f"lap_rate_{i}"].append(wl_c / w_lessons_cls)
                        final_col_values[f"test_rate_{i}"].append(wt_c / w_lessons_cls)
                        final_col_values[f"hw_{i}"].append(round(thw_c / w_lessons_cls, 1))

                for r in cls_records:
                    for k in export_keys:
                        val = r.get(k)
                        if k == "base_no": val = None if val is None else (str(val).strip().lstrip("0") or str(val).strip())
                        elif k == "base_total" or k.startswith("lesson_") or k.startswith("lap_cnt_") or k.startswith("test_cnt_"):
                            val = 0.0 if val is None else float(val)
                        elif k.startswith("hw_") or k.startswith("lap_rate_") or k.startswith("test_rate_"):
                            val = None if val is None else float(val)
                        else:
                            val = None if val is None else str(val)
                        final_col_values[k].append(val)

        nrows = len(final_col_values["base_no"])

        progress(f"[{sheet_title}] ステップ3: 空白値のハイフン変換処理")
        for k in export_keys:
            vals = final_col_values[k]
            for j, v in enumerate(vals):
                if v is None: vals[j] = "-"

        progress(f"[{sheet_title}] ステップ4: 列幅AutoFit計算とシート生成")
        worksheet = workbook.add_worksheet(sheet_title)
        worksheet.set_tab_color(tab_color)

        # ページ設定をAppConfigから適用
        worksheet.center_horizontally()
        worksheet.set_paper(AppConfig.PAPER_SIZE)
        worksheet.set_margins(AppConfig.MARGIN_LR, AppConfig.MARGIN_LR, AppConfig.MARGIN_TB, AppConfig.MARGIN_TB)
        if AppConfig.IS_LANDSCAPE:
            worksheet.set_landscape()
        else:
            worksheet.set_portrait()

        display_headers = [display_map[k] for k in export_keys]
        system_col_numbers = [col_idx + 1 for col_idx in range(len(export_keys))]
        max_col = len(export_keys) - 1

        for col_idx, internal_key in enumerate(export_keys):
            d_name = display_map[internal_key]
            max_val_len = 0
            for v in final_col_values[internal_key]:
                w = calculate_visual_width(str(v))
                if w > max_val_len: max_val_len = int(w)

            width, format_key, _ = get_column_style_config(internal_key, d_name)
            header_w = calculate_visual_width(d_name)
            final_w = max(max_val_len + 4, header_w + 4, 12, width)

            base_format = formats_matrix["base"][format_key]
            worksheet.set_column(col_idx, col_idx, final_w, base_format)

        try: worksheet.merge_range(0, 0, 0, min(3, max_col), date_range_str, fmt_title)
        except Exception: worksheet.write(0, 0, date_range_str, fmt_title)
        worksheet.set_row(0, 20)

        week_block_size = len(lesson_nums)
        first_week_col = 4 # base_* の4列分
        cols_per_week = 6  # 1週あたりの列数

        for i, period_str in enumerate(week_period_labels):
            start_col = first_week_col + i * cols_per_week
            if start_col + (cols_per_week - 1) <= max_col:
                worksheet.merge_range(0, start_col, 0, start_col + (cols_per_week - 1), period_str, fmt_title_center)

        header_row_visual = 1
        header_row_system = 2
        data_start = 3

        worksheet.write_row(header_row_visual, 0, display_headers, fmt_header)
        worksheet.set_row(header_row_system, 12)
        worksheet.write_row(header_row_system, 0, system_col_numbers, fmt_system_sub)

        progress(f"[{sheet_title}] ステップ5: セル書き込み開始 (計 {nrows} 行)")
        for r_idx in range(nrows):
            is_avg = (final_col_values["base_name"][r_idx] == "AVG")
            current_row_idx = data_start + r_idx

            if is_avg:
                row_type = "avg"
                for col_idx, internal_key in enumerate(export_keys):
                    val = final_col_values[internal_key][r_idx]
                    d_name = display_map[internal_key]
                    _, format_key, _ = get_column_style_config(internal_key, d_name)
                    worksheet.write(current_row_idx, col_idx, val, formats_matrix[row_type][format_key])
            else:
                row_data = [final_col_values[k][r_idx] for k in export_keys]
                worksheet.write_row(current_row_idx, 0, row_data)

        progress(f"[{sheet_title}] ステップ6: 条件付き書式・印刷設定の適用")
        last_row = data_start + nrows - 1
        worksheet.autofilter(header_row_system, 0, last_row, max_col)

        # アラート設定をAppConfigから適用
        for c_idx, internal_key in enumerate(export_keys):
            if internal_key.startswith("hw_"):
                worksheet.conditional_format(data_start, c_idx, last_row, c_idx, {"type": "cell", "criteria": "<", "value": AppConfig.ALERT_HW_UNDER, "format": fmt_cond_orange})
            elif internal_key.startswith("lap_rate_") or internal_key.startswith("test_rate_"):
                worksheet.conditional_format(data_start, c_idx, last_row, c_idx, {"type": "cell", "criteria": "<", "value": AppConfig.ALERT_RATE_UNDER, "format": fmt_cond_firebrick})

        worksheet.freeze_panes(data_start, first_week_col)
        worksheet.repeat_columns(0, first_week_col - 1)
        worksheet.repeat_rows(0, header_row_system)

        v_breaks = []
        current_break_col = first_week_col + cols_per_week
        while current_break_col < len(export_keys):
            v_breaks.append(current_break_col)
            current_break_col += cols_per_week
        worksheet.set_v_pagebreaks(v_breaks)
        worksheet.print_area(0, 0, last_row, max_col)
        progress(f"[{sheet_title}] の書き込み処理完了")

    write_sheet(records_all, summary_all, "全集計", AppConfig.COLOR_TAB_ALL)
    write_sheet(records_excl, summary_excl, "除外済み集計", AppConfig.COLOR_TAB_EXCL)

    progress("Excelブックのクローズ処理中...")
    workbook.close()

def main():
    dismiss_splash()
    enable_windows_dpi_awareness()
    root = tk.Tk()
    configure_tk_scaling(root)
    root.withdraw()
    try:
        root.attributes("-topmost", True); root.lift(); root.focus_force()
        input_file = filedialog.askopenfilename(title="集計元のExcelファイルを選択してください", filetypes=[("集計データ","scube2-lesson-result_*.xlsx"), ("Excel", "*.xlsx"), ("すべてのファイル", "*.*")], parent=root)
        if not input_file:
            root.destroy()
            return

        start_from_saturday = messagebox.askyesno(
            "集計基準の選択",
            "土曜始まり（土〜金）で集計しますか？\n\n【はい】: 土曜始まり\n【いいえ】: 月曜始まり",
            parent=root
        )
        root.attributes("-topmost", False)

        total_start_time = time.perf_counter()

        (
            records_all,
            records_excl,
            summary_all,
            summary_excl,
            output_file,
            date_range_str,
            sheet_name,
            week_period_labels,
            export_keys,
            display_map,
            lesson_nums
        ) = process_attendance_data(input_file, start_from_saturday)

        if not records_all:
            root.attributes("-topmost", True)
            messagebox.showwarning("データなし", "集計対象のデータが見つかりませんでした。", parent=root)
            root.destroy()
            return

        try:
            if os.path.exists(output_file):
                with open(output_file, "r+"): pass
        except IOError:
            root.attributes("-topmost", True)
            messagebox.showerror("実行エラー", f"出力先のExcelファイルが開かれたままです。\nファイルを閉じてから再度実行してください。\n\n対象ファイル:\n{os.path.basename(output_file)}", parent=root)
            root.destroy()
            return

        format_excel_fast(
            records_all,
            records_excl,
            summary_all,
            summary_excl,
            output_file,
            date_range_str,
            sheet_name,
            week_period_labels,
            export_keys,
            display_map,
            lesson_nums
        )

        total_elapsed = time.perf_counter() - total_start_time
        processed_classroom_count = len(summary_all)
        try: processed_lesson_count = int(sum(v for v in summary_all.values()))
        except Exception: processed_lesson_count = 0

        result_msg = (
            f"処理が完了しました！\n実行時間: {total_elapsed:.2f}秒\n\n"
            f"処理教室数: {processed_classroom_count}教室\n授業数: {processed_lesson_count}件\n\n"
            f"出力先:\n{output_file}\n\n"
            f"作成されたファイルを開きますか？"
        )

        root.attributes("-topmost", True)
        open_file = messagebox.askyesno("集計完了", result_msg, parent=root)
        root.attributes("-topmost", False)

        if open_file:
            suggest_msg=(
                f"列幅を調整するために、以下を入力してください。\n"
                f"絶対ではないので、不要であれば無視してください\n\n"
                f"全選択：[Ctrl]+[A]\n"
                f"↓\n"
                f"列の自動調整：[Alt]→[H]→[O]→[I]"
            )
            os.startfile(output_file)

            messagebox.showinfo("おすすめの操作",suggest_msg, parent=root)

        progress(f"処理完了(実行時間: {total_elapsed:.2f}秒)")
    except KeyboardInterrupt: pass
    except Exception as e:
        tb = traceback.format_exc()
        try: root.attributes("-topmost", True); show_error_dialog("処理中にエラーが発生しました", tb, parent=root)
        except Exception: messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}", parent=root)
    finally: root.destroy()

if __name__ == "__main__": main()