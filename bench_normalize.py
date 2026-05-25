import time
import random
from MSLdata_check_v4 import normalize_text_for_matching

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

arr = [random.choice(samples) + str(i) for i in range(N)]

start = time.perf_counter()
for s in arr:
    normalize_text_for_matching(s)
end = time.perf_counter()

print(f"Normalized {N} strings in {end-start:.3f}s (per={ (end-start)/N :.6f}s)")
