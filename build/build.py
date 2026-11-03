"""
MSL Checker ビルドスクリプト (互換用エントリポイント)
現行の build2.py を呼び出します。
"""
from __future__ import annotations
import sys
from pathlib import Path

# build2.py の main を呼び出す
script_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(script_dir))
import build2

if __name__ == "__main__":
    sys.exit(build2.main())
