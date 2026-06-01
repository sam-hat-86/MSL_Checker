from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import shutil
from pathlib import Path


APP_NAME_BASE = "MSL集計ソフト"
SCRIPT_BASE_PREFIX = "MSLdata_check_v"
DEFAULT_VERSION = "5"
ICON_NAME = "logo.ico"
BUILD_REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements-build.txt"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MSL Checker を Nuitka でビルドします。")
    parser.add_argument("-v", "--version", help="使用するバージョン番号 (1-5)")
    parser.add_argument(
        "--jobs",
        type=int,
        default=os.cpu_count() or 6,
        help="Nuitka の並列ジョブ数 (既定: CPU コア数)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="除外されるモジュール一覧を表示してビルドを実行しません（安全確認用）。",
    )
    return parser.parse_args()


def normalize_version(raw_version: str | None) -> str:
    if raw_version is None or raw_version.strip() == "":
        return DEFAULT_VERSION

    version = raw_version.strip()
    if version.lower().startswith("v"):
        version = version[1:]

    if not version.isdigit():
        raise ValueError("数字を入力してください。")

    return version


def prompt_version() -> str:
    print("使用するバージョンを入力してください。")
    return normalize_version(input("番号を入力してください: "))


def available_versions(project_root: Path) -> list[str]:
    src_dir = project_root / "src"
    versions: list[str] = []
    for path in src_dir.glob(f"{SCRIPT_BASE_PREFIX}*.py"):
        match = re.fullmatch(rf"{re.escape(SCRIPT_BASE_PREFIX)}(\d+)\.py", path.name)
        if match:
            versions.append(match.group(1))
    return sorted(set(versions), key=int)


def requirements_file_for_version(project_root: Path, version: str) -> Path:
    return project_root / "src" / f"requirements_v{version}.txt"


def build_command(
    project_root: Path,
    script_path: Path,
    exe_name: str,
    jobs: int,
) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "nuitka",
        "--onefile",
        "--lto=yes",
        "--remove-docstrings",
        "--enable-plugin=tk-inter",
        "--windows-console-mode=disable",
        f"--windows-icon-from-ico={project_root / 'build' / ICON_NAME}",
        f"--jobs={max(1, jobs)}",
        "--assume-yes-for-downloads",
        f"--output-dir={project_root}",
        f"--output-filename={exe_name}",
        "--remove-output",
        "--no-deployment-flag=self-execution",
        "--nofollow-import-to=unittest",
        "--nofollow-import-to=pydoc",
        "--nofollow-import-to=IPython",
        "--nofollow-import-to=notebook",
        "--nofollow-import-to=numpy.random",
        "--nofollow-import-to=matplotlib",
        "--nofollow-import-to=PIL",
        str(script_path),
    ]

    # LTO は常時有効化して実行時性能を優先します。
    return command


def detect_top_level_imports(script_path: Path, project_root: Path) -> set[str]:
    """Parse the target script and return a set of top-level module names imported.

    This function will recursively scan local modules under `project_root/src` to
    include imports that are declared in project modules.
    """
    try:
        import ast
        from collections import deque

        src_root = project_root / "src"

        def parse_file(path: Path) -> list[str]:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    src = f.read()
                tree = ast.parse(src)
            except Exception:
                return []

            found: list[str] = []
            for node in tree.body:
                if isinstance(node, ast.Import):
                    for n in node.names:
                        found.append(n.name)
                elif isinstance(node, ast.ImportFrom):
                    if node.module:
                        found.append(node.module)
            return found

        names: set[str] = set()
        visited_files: set[Path] = set()
        queue = deque()
        queue.append(script_path.resolve())

        while queue:
            fp = queue.popleft()
            if fp in visited_files:
                continue
            visited_files.add(fp)
            parts = parse_file(fp)
            for full_name in parts:
                top = full_name.split(".")[0]
                names.add(top)
                # if this import corresponds to a local module, enqueue it for parsing
                candidate = src_root / (full_name.replace(".", os.sep) + ".py")
                if candidate.exists() and candidate not in visited_files:
                    queue.append(candidate.resolve())
                else:
                    # also try top-level module file (module.py)
                    candidate2 = src_root / (top + ".py")
                    if candidate2.exists() and candidate2 not in visited_files:
                        queue.append(candidate2.resolve())

        return names
    except Exception:
        return set()


def clean_outputs(project_root: Path, build_dir: Path, exe_name: str) -> None:
    output_exe = project_root / f"{exe_name}.exe"
    if output_exe.exists():
        output_exe.unlink()

    for suffix in (".build", ".onefile-build"):
        target = build_dir / f"{exe_name}{suffix}"
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)


def main() -> int:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent
    build_dir = script_dir
    versions = available_versions(project_root)

    if not versions:
        print("[ERROR] 対象のバージョンファイルが見つかりません。")
        return 1

    print("使用できるバージョン:")
    for version in versions:
        print(f"  {version} -> v{version}")

    try:
        version = normalize_version(args.version) if args.version is not None else prompt_version()
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    if version not in versions:
        print("[ERROR] 指定されたバージョンのスクリプトが見つかりません。")
        return 1

    script_name = f"{SCRIPT_BASE_PREFIX}{version}.py"
    script_path = project_root / "src" / script_name
    requirements_file = requirements_file_for_version(project_root, version)

    exe_name = f"{APP_NAME_BASE}_v{version}"

    print("=== STEP 1: CLEANING ===")
    clean_outputs(project_root, build_dir, exe_name)

    print("=== STEP 2: NUITKA BUILD ===")
    if requirements_file.exists():
        print(f"使用する依存関係: {requirements_file}")
    if BUILD_REQUIREMENTS_FILE.exists():
        print(f"ビルド依存関係: {BUILD_REQUIREMENTS_FILE}")
    print("LTO: 有効（実行時性能を最優先にしています）")
    command = build_command(project_root, script_path, exe_name, args.jobs)

    # 最小限のインポートに絞るため、ターゲットとプロジェクト内モジュールを解析し、
    # 指定された requirements_v{version}.txt をホワイトリストとして使用します。
    top_imports = detect_top_level_imports(script_path, project_root)
    print(f"検出されたトップレベル imports: {sorted(top_imports)}")

    # requirements ファイルを読み、パッケージ名リストを作る
    reqs: set[str] = set()
    try:
        if requirements_file.exists():
            for line in requirements_file.read_text(encoding="utf-8").splitlines():
                s = line.strip()
                if not s or s.startswith("#"):
                    continue
                # バージョン指定や extras を除去
                s = re.split(r"[<=>!\[\];]", s)[0]
                s = s.strip().lower()
                if s:
                    reqs.add(s)
    except Exception:
        reqs = set()

    # project-level exclude/allow 設定があれば読み込む
    exclude_cfg_file = project_root / "build" / "exclude.json"
    cfg_allow: set[str] = set()
    cfg_deny: set[str] = set()
    if exclude_cfg_file.exists():
        try:
            cfg = json.loads(exclude_cfg_file.read_text(encoding="utf-8"))
            for a in cfg.get("allow", []):
                if isinstance(a, str):
                    cfg_allow.add(a)
            for d in cfg.get("deny", []):
                if isinstance(d, str):
                    cfg_deny.add(d)
            print(f"読み込んだ除外設定: allow={sorted(cfg_allow)}, deny={sorted(cfg_deny)}")
        except Exception:
            print("exclude.json の読み込みに失敗しました。無視します。")

    # 一部パッケージ名と import 名の差分マッピング
    alias_map = {
        "pillow": "PIL",
        "python-dateutil": "dateutil",
        "pyyaml": "yaml",
    }

    allowed: set[str] = set()
    for r in reqs:
        if r in alias_map:
            allowed.add(alias_map[r])
        else:
            allowed.add(r)
    # 設定ファイルの allow を追加
    for a in cfg_allow:
        allowed.add(a)

    # プロジェクト内モジュールは常に許可
    src_root = project_root / "src"
    for imp in list(top_imports):
        candidate = src_root / (imp + ".py")
        candidate_pkg = src_root / imp / "__init__.py"
        if candidate.exists() or candidate_pkg.exists():
            allowed.add(imp)

    # 組み込みモジュールは許可
    import sys as _sys

    builtin = set([n for n in _sys.builtin_module_names if isinstance(n, str)])

    # 除外対象: top_imports のうち allowed に含まれないもの
    excluded = []
    for mod in sorted(top_imports):
        if mod in cfg_deny:
            command.append(f"--nofollow-import-to={mod}")
            excluded.append(mod)
            continue
        low = mod.lower()
        if mod in allowed or low in allowed or mod in builtin:
            continue
        # mod が requirements に記載されていなければ除外フラグを付ける
        command.append(f"--nofollow-import-to={mod}")
        excluded.append(mod)

    if excluded:
        print(f"requirements と照合し除外した import: {sorted(excluded)}")
    # PYTHON 最適化フラグを有効に（-OO 相当）
    env = os.environ.copy()
    env["PYTHONOPTIMIZE"] = "2"

    # dry-run: 除外一覧を表示してコマンドは実行しない
    if getattr(args, "dry_run", False):
        print()
        print("=== DRY RUN ===")
        print("生成予定の Nuitka コマンド:")
        print(" ".join(command))
        print(f"環境: PYTHONOPTIMIZE={env.get('PYTHONOPTIMIZE')}")
        print()
        print("ビルドはスキップしました（dry-run）。")
        return 0

    print("PYTHON 最適化: -OO を適用します (PYTHONOPTIMIZE=2)")
    completed = subprocess.run(command, cwd=build_dir, env=env)
    if completed.returncode != 0:
        print("[ERROR] ビルドに失敗しました。")
        return completed.returncode

    output_exe = project_root / f"{exe_name}.exe"
    print()
    print("=== STEP 3: RESULT ===")
    if output_exe.exists():
        print("[SUCCESS] ビルドが完了しました。")
        print(f"保存先: {output_exe}")
        return 0

    print("[ERROR] ビルドに失敗しました。")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())