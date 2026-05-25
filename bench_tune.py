import time
import multiprocessing as mp
from MSLdata_check_v4 import normalize_text_for_matching
import os


def main() -> None:
    # サンプル生成（bench_normalize.py と同じ設定）
    N = 20000
    samples = [
        "ターゲット １０ページ",
        "たんご",
        "Sample　テスト",
        "ＡＢＣ123",
        "カタカナ",
        "ひらがな",
        "見本 サンプル",
        "確認用フォーマット",
        "リープLEAP",
    ]
    arr = [samples[i % len(samples)] + str(i) for i in range(N)]

    cpu_total = mp.cpu_count()
    # 候補: 明示的な小さめ値と自動(cpu-1)を含める
    worker_candidates = [max(1, cpu_total - 1), 2, 4, 6]
    # 除去重複・範囲
    worker_candidates = sorted(set([w for w in worker_candidates if w >= 1 and w <= cpu_total]))
    chunks_candidates = [None, 100, 500, 1000, 2000]

    results = []

    for workers in worker_candidates:
        for chunks in chunks_candidates:
            # 自動 chunksize
            if chunks is None:
                chunksize = max(1, len(arr) // (workers * 4))
                label_chunks = f"auto({chunksize})"
            else:
                chunksize = chunks
                label_chunks = str(chunks)

            # ベンチ
            try:
                start = time.perf_counter()
                with mp.Pool(processes=workers) as pool:
                    pool.map(normalize_text_for_matching, arr, chunksize)
                elapsed = time.perf_counter() - start
                results.append((elapsed, workers, chunksize))
                print(f"workers={workers:2d}, chunksize={label_chunks:10s} -> {elapsed:.3f}s")
            except Exception as e:
                print(f"workers={workers}, chunks={label_chunks} -> ERROR: {e}")

    # 最良結果を表示
    results_sorted = sorted(results, key=lambda x: x[0])
    if results_sorted:
        best = results_sorted[0]
        print("\nBest:")
        print(f"{best[0]:.3f}s with workers={best[1]} chunksize={best[2]}")
    else:
        print("No successful runs")


if __name__ == '__main__':
    # Windows の場合は freeze_support を呼ぶ
    try:
        mp.freeze_support()
    except Exception:
        pass
    main()
