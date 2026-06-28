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
DEFAULT_VERSION = "10"
ICON_NAME = "logo.ico"
BUILD_REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements-build.txt"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MSL Checker を Nuitka でビルドします。")
    parser.add_argument("-v", "--version", help=f"使用するバージョン番号 (1-{DEFAULT_VERSION})")
    parser.add_argument(
        "-m",
        "--mode",
        choices=["1", "2", "3", "dev", "pre", "release"],
        help="ビルドモード (1: dev, 2: pre, 3: release)",
    )
    parser.add_argument(
        "--jobs",
        type=int,
        default=os.cpu_count() or 6,
        help="Nuitka の並列ジョブ数 (既定: CPUコア数)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="実際にビルドせず、生成されるコマンドのみを確認します",
    )
    return parser.parse_args()


def normalize_version(s: object) -> str:
    if s is None: return ""
    return str(s).strip()


def prompt_version() -> str:
    v = input(f"ビルドするバージョン番号を入力してください (既定: {DEFAULT_VERSION}): ").strip()
    return v if v else DEFAULT_VERSION


def prompt_build_mode() -> str:
    print("ビルドモードを選択してください:")
    print("  1: dev (LTOなし / コンソール強制表示)")
    print("  2: pre (LTOあり / コンソール強制表示)")
    print("  3: release (LTOあり / コンソール非表示) 【本番用】")
    choice = input("選択 (1-3, 既定: 3): ").strip()
    if choice == "1": return "dev"
    if choice == "2": return "pre"
    return "release"


def available_versions(project_root: Path) -> list[str]:
    src_dir = project_root / "src"
    if not src_dir.exists(): return []
    versions = []
    for f in src_dir.glob(f"{SCRIPT_BASE_PREFIX}*.py"):
        m = re.match(rf"{SCRIPT_BASE_PREFIX}(\d+)\.py", f.name)
        if m:
            versions.append(m.group(1))
    return sorted(versions, key=lambda x: int(x))


def find_requirements_file(project_root: Path, version: str) -> Path | None:
    req_v = project_root / "src" / f"requirements_v{version}.txt"
    if req_v.exists(): return req_v

    req_git = project_root / "src" / "requirements_git-act.txt"
    if req_git.exists(): return req_git

    req_std = project_root / "src" / "requirements.txt"
    if req_std.exists(): return req_std

    return None


def build_command(project_root: Path, script_path: Path, exe_name: str, jobs: int, mode: str, version: str) -> list[str]:
    common_exclusions = [
        "--nofollow-import-to=pandas",
        "--nofollow-import-to=numpy",
        "--nofollow-import-to=fastexcel",
        "--nofollow-import-to=polars",
        "--nofollow-import-to=matplotlib",
        "--nofollow-import-to=scipy",
        "--nofollow-import-to=tangowidth",
        "--nofollow-import-to=unittest",
        "--nofollow-import-to=IPython",
        "--nofollow-import-to=notebook",
        "--nofollow-import-to=pytest",
        "--nofollow-import-to=sqlite3",
    ]

    command = [
        sys.executable,
        "-m",
        "nuitka",
        "--onefile",
        "--follow-imports",
        "--enable-plugin=tk-inter",
        f"--jobs={max(1, jobs)}",
        "--assume-yes-for-downloads",
        f"--output-dir={project_root}",
        f"--output-filename={exe_name}",
        "--remove-output",
        "--no-deployment-flag=self-execution",
    ]

    ico_path = project_root / 'build' / ICON_NAME
    if ico_path.exists():
        command.append(f"--windows-icon-from-ico={ico_path}")

    command.extend(common_exclusions)

    if mode == "dev":
        command.extend([
            "--lto=no",
            "--windows-console-mode=force",
        ])
    elif mode == "pre":
        command.extend([
            "--lto=yes",
            "--windows-console-mode=force",
        ])
    elif mode == "release":
        command.extend([
            "--lto=yes",
            "--windows-console-mode=disable",
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

    for suffix in (".build", ".onefile-build"):
        target = project_root / f"{exe_name}{suffix}"
        if target.exists():
            shutil.rmtree(target, ignore_errors=True)


def main() -> int:
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent
    build_dir = script_dir
    versions = available_versions(project_root)

    if not versions:
        print("[ERROR] 対象のバージョンファイルが見つかりません。 (src/MSLdata_check_v*.py)")
        return 1

    print("検出されたバージョン一覧:")
    for v in versions:
        print(f"  {v} -> v{v}")

    try:
        version = normalize_version(args.version) if args.version is not None else prompt_version()
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    if version not in versions:
        print(f"[ERROR] 指定されたバージョン v{version} のスクリプトが見つかりません。")
        return 1

    mode = "release"
    if args.mode:
        if args.mode in ("1", "dev"):
            mode = "dev"
        elif args.mode in ("2", "pre"):
            mode = "pre"
        elif args.mode in ("3", "release"):
            mode = "release"
    else:
        mode = prompt_build_mode()

    script_name = f"{SCRIPT_BASE_PREFIX}{version}.py"
    script_path = project_root / "src" / script_name
    requirements_file = find_requirements_file(project_root, version)

    exe_name = f"{APP_NAME_BASE}_v{version}"
    if mode == "dev":
        exe_name += "_dev"
    elif mode == "pre":
        exe_name += "_pre"

    print("\n=== STEP 1: CLEANING ===")
    clean_outputs(project_root, build_dir, exe_name)

    print("\n=== STEP 2: NUITKA BUILD ===")
    if requirements_file and requirements_file.exists():
        print(f"解析されたパッケージ依存関係: {requirements_file}")
    if BUILD_REQUIREMENTS_FILE.exists():
        print(f"ビルド支援依存関係: {BUILD_REQUIREMENTS_FILE}")

    mode_descriptions = {
        "dev": "dev (LTOなし / コンソール強制表示 / 高速ビルド)",
        "pre": "pre (LTOあり / コンソール強制表示 / 最適化実行)",
        "release": "release (LTOあり / コンソール非表示 / 本番配布用)",
    }
    print(f"ビルドターゲットモード: {mode_descriptions[mode]}")

    command = build_command(project_root, script_path, exe_name, args.jobs, mode, version)

    env = os.environ.copy()
    env["PYTHONOPTIMIZE"] = "1" if mode in ("pre", "release") else "0"

    if args.dry_run:
        print()
        print("=== DRY RUN MODE ===")
        print("生成されるコマンド:")
        print(" ".join(command))
        print("==========================================")
        return 0

    print("コンパイルを実行しています。数分かかる場合があります...")
    try:
        result = subprocess.run(command, shell=True)
        if result.returncode != 0:
            print("[エラー] Nuitkaコンパイル中に不具合が発生しました。")
            return 1
    except Exception as e:
        print(f"[エラー] ビルド実行中に例外が発生しました: {e}")
        return 1

    output_exe = project_root / f"{exe_name}.exe"
    print("\n==========================================")
    if output_exe.exists():
        print("[SUCCESS] ビルドが正常に完了しました！")
        print(f"保存先: {output_exe}")
        return 0

    print("[エラー] EXEファイルの生成が確認できませんでした。")
    return 1


if __name__ == "__main__":
    main()