from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import shutil
from pathlib import Path


APP_NAME_BASE = "MSL集計ソフト"
SCRIPT_BASE_PREFIX = "MSLdata_check_v"
DEFAULT_VERSION = "9"
ICON_NAME = "logo.ico"
BUILD_REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements-build.txt"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MSL Checker を Nuitka でビルドします。")
    parser.add_argument("-v", "--version", help="使用するバージョン番号 (1-9)")
    parser.add_argument(
        "--release",
        action="store_true",
        help="リリースモードでビルドします（指定がない場合は確認用デバッグモード）。",
    )
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


def prompt_build_mode() -> bool:
    print("ビルドモードを選択してください。")
    print("  1 -> 確認用（除外最大・コンソール表示・高速ビルド）")
    print("  2 -> リリース用（除外最大・コンソール非表示・最適化実行）")
    choice = input("番号を入力してください (既定: 1): ").strip()
    return choice == "2"


def available_versions(project_root: Path) -> list[str]:
    src_dir = project_root / "src"
    versions: list[str] = []
    if not src_dir.exists():
        return versions
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
    is_release: bool,
) -> list[str]:
    # どちらのモードでも共通して適用する、最大限の不要モジュール除外リスト
    common_exclusions = [
        "--nofollow-import-to=unittest",
        "--nofollow-import-to=IPython",
        "--nofollow-import-to=notebook",
        "--nofollow-import-to=matplotlib",
        "--nofollow-import-to=scipy",
        "--nofollow-import-to=scipy.stats",
        "--nofollow-import-to=numba",
        "--nofollow-import-to=jinja2",
        "--nofollow-import-to=PIL",
        "--nofollow-import-to=sqlite3",
        "--nofollow-import-to=distutils",
        "--nofollow-import-to=setuptools",
        "--nofollow-import-to=pytest",
        "--nofollow-import-to=numpy.testing",
    ]

    command = [
        sys.executable,
        "-m",
        "nuitka",
        "--onefile",
        "--follow-imports",
        "--enable-plugin=tk-inter",
        f"--windows-icon-from-ico={project_root / 'build' / ICON_NAME}",
        f"--jobs={max(1, jobs)}",
        "--assume-yes-for-downloads",
        f"--output-dir={project_root}",
        f"--output-filename={exe_name}",
        "--remove-output",
        "--no-deployment-flag=self-execution",
    ]

    # 除外設定をベースコマンドに追加
    command.extend(common_exclusions)

    if is_release:
        # リリース用：最適化を最大化し、コンソールは隠す
        command.extend([
            "--lto=yes",
            "--windows-console-mode=disable",
        ])
    else:
        # 確認用：Cのコンパイル速度を最優先し、エラー特定用にコンソールを強制表示
        command.extend([
            "--lto=no",
            "--windows-console-mode=force",
        ])

    command.append(str(script_path))
    return command


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

    is_release = args.release if args.release else prompt_build_mode()

    script_name = f"{SCRIPT_BASE_PREFIX}{version}.py"
    script_path = project_root / "src" / script_name
    requirements_file = requirements_file_for_version(project_root, version)

    exe_name = f"{APP_NAME_BASE}_v{version}"
    if not is_release:
        exe_name += "_dev"

    print("=== STEP 1: CLEANING ===")
    clean_outputs(project_root, build_dir, exe_name)

    print("=== STEP 2: NUITKA BUILD ===")
    if requirements_file.exists():
        print(f"使用する依存関係: {requirements_file}")
    if BUILD_REQUIREMENTS_FILE.exists():
        print(f"ビルド依存関係: {BUILD_REQUIREMENTS_FILE}")

    print(f"ビルドモード: {'リリース (最適化・コンソール非表示)' if is_release else '確認用 (高速ビルド・コンソール表示)'}")
    print("不要モジュール除外設定: 最大化（両モード共通）")
    print(f"LTO: {'有効' if is_release else '無効'}")

    command = build_command(project_root, script_path, exe_name, args.jobs, is_release)

    env = os.environ.copy()
    env["PYTHONOPTIMIZE"] = "1" if is_release else "0"

    if getattr(args, "dry_run", False):
        print()
        print("=== DRY RUN ===")
        print("生成予定の Nuitka コマンド:")
        print(" ".join(command))
        print(f"環境: PYTHONOPTIMIZE={env.get('PYTHONOPTIMIZE')}")
        print()
        print("ビルドはスキップしました（dry-run）。")
        return 0

    print(f"PYTHON 最適化: {'-O (PYTHONOPTIMIZE=1)' if is_release else '標準 (PYTHONOPTIMIZE=0)'}")
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