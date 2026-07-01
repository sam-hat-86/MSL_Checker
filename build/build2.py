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
MIN_VERSION="11"
MAX_VERSION="15"
DEFAULT_VERSION=MAX_VERSION
ICON_NAME = "logo.ico"
SPLASH_NAME = "splash.png"
BUILD_REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements-build.txt"

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MSL Checker を Nuitka でビルドします。")
    parser.add_argument("-v", "--version", help=f"使用するバージョン番号 ({MIN_VERSION}-{MAX_VERSION})")
    parser.add_argument(
        "-m",
        "--mode",
        choices=["1", "2", "dev","release"],
        help="ビルドモード (1: dev, 2: release)",
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

def prompt_version()->str:
    v=input(f"ビルドバージョンを入力 ({MIN_VERSION}-{MAX_VERSION},既定:{DEFAULT_VERSION}): ").strip()
    return v if v else DEFAULT_VERSION

def prompt_build_mode() -> str:
    print("ビルドモードを選択してください:")
    print("  1: dev (コンソールあり)")
    print("  2: release (コンソールなし)")
    choice = input("選択 (1-2, 既定: 1): ").strip()
    if choice == "2": return "release"
    return "dev"

def available_versions(project_root:Path)->list[str]:
    search_dirs=[project_root/"src",project_root]
    versions=set()
    min_ver=int(MIN_VERSION)
    max_ver=int(MAX_VERSION)

    for d in search_dirs:
        if d.exists():
            for f in d.glob(f"{SCRIPT_BASE_PREFIX}*.py"):
                m=re.match(rf"{SCRIPT_BASE_PREFIX}(\d+)\.py",f.name)
                if m:
                    v=int(m.group(1))
                    if min_ver<=v<=max_ver:
                        versions.add(str(v))

    return sorted(versions,key=int)

def find_script_path(project_root: Path, version: str) -> Path | None:
    candidates = [
        project_root / "src" / f"{SCRIPT_BASE_PREFIX}{version}.py",
        project_root / f"{SCRIPT_BASE_PREFIX}{version}.py"
    ]
    for c in candidates:
        if c.exists():
            return c
    return None

def find_requirements_file(project_root: Path, version: str) -> Path | None:
    candidates = [
        project_root / "src" / f"requirements_v{version}.txt",
        project_root / f"requirements_v{version}.txt",
        project_root / "src" / "requirements_git-act.txt",
        project_root / "requirements_git-act.txt",
        project_root / "src" / "requirements.txt",
        project_root / "requirements.txt"
    ]
    for c in candidates:
        if c.exists():
            return c
    return None

def build_command(project_root: Path, script_path: Path, exe_name: str, jobs: int, mode: str, version: str) -> list[str]:
    common_exclusions = [
        "--nofollow-import-to=tzdata",
        "--nofollow-import-to=unittest",
        "--nofollow-import-to=IPython",
        "--nofollow-import-to=notebook",
        "--nofollow-import-to=pytest",
        "--nofollow-import-to=distutils",
        "--nofollow-import-to=setuptools",
        "--nofollow-import-to=ssl",
        "--nofollow-import-to=asyncio",
        "--nofollow-import-to=http",
        "--nofollow-import-to=ftplib",
        "--nofollow-import-to=encodings.idna",
        "--nofollow-import-to=stringprep",
    ]

    v_num = 0
    try:
        v_num = int(version)
    except ValueError:
        pass

    version_exclusions = []
    version_allowances = []

    common_unused = [
        "--nofollow-import-to=pandas",
        "--nofollow-import-to=numpy",
        "--nofollow-import-to=matplotlib",
        "--nofollow-import-to=scipy",
        "--nofollow-import-to=cryptography",
        "--nofollow-import-to=jinja2",
        "--nofollow-import-to=pytz",
        "--nofollow-import-to=pyarrow.tests.*",
    ]
    version_exclusions.extend(common_unused)

    if v_num >= 13:
        version_exclusions.extend([
            "--nofollow-import-to=polars",
        ])
        print(f"[INFO] v13以降向けビルド定義：Polarsを除外します。")

    if v_num == 13:
        version_exclusions.extend([
            "--nofollow-import-to=pyarrow",
        ])
    else:
        version_allowances.extend([
            "--include-package=pyarrow",
            "--include-package=fastexcel",
        ])

    common_exclusions.extend(version_exclusions)

    command = [
        sys.executable,
        "-m",
        "nuitka",
        "--onefile",
        "--follow-imports",
        "--enable-plugin=tk-inter",
        f"--jobs={max(4, jobs)}",
        "--assume-yes-for-downloads",
        f"--output-dir={project_root}",
        f"--output-filename={exe_name}",
        "--remove-output",
        "--no-deployment-flag=self-execution",
        "--python-flag=no_docstrings",
        "--python-flag=no_asserts",
        "--lto=no",
    ]

    ico_path = project_root / 'build' / ICON_NAME
    if ico_path.exists():
        command.append(f"--windows-icon-from-ico={ico_path}")

    splash_path = project_root / 'build' / SPLASH_NAME
    if splash_path.exists():
        command.append(f"--onefile-windows-splash-screen-image={splash_path}")
        print(f"[INFO] スプラッシュ画面を適用します: {splash_path}")

    command.extend(common_exclusions)
    command.extend(version_allowances)

    if mode == "dev":
        command.extend([
            "--windows-console-mode=force",
        ])
    else:
        command.extend([
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
        print("[ERROR] 対象のバージョンファイルが見つかりません。")
        print(f"期待される命名規則: {SCRIPT_BASE_PREFIX}*.py")
        return 1

    print("検出されたバージョン一覧")
    for v in versions:
        print(f"  {v} -> v{v}")

    try:
        version = normalize_version(args.version) if args.version is not None else prompt_version()
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    if version not in versions:
        print(f"[ERROR] 指定されたバージョン v{version} は対象外です。")
        print(f"対象バージョン: v{MIN_VERSION}～v{MAX_VERSION}")
        return 1

    mode = "dev"
    if args.mode:
        if args.mode in ("1", "dev"):
            mode = "dev"
        elif args.mode in ("2", "release"):
            mode = "release"
    else:
        mode = prompt_build_mode()

    script_path = find_script_path(project_root, version)
    if not script_path:
        print(f"[ERROR] スクリプトのフルパス取得に失敗しました。")
        return 1

    requirements_file = find_requirements_file(project_root, version)

    exe_name = f"{APP_NAME_BASE}_v{version}"
    if mode == "dev":
        exe_name += "_dev"

    print("\n=== STEP 1: CLEANING ===")
    clean_outputs(project_root, build_dir, exe_name)

    print("\n=== STEP 2: NUITKA BUILD ===")
    if requirements_file and requirements_file.exists():
        print(f"解析されたパッケージ依存関係: {requirements_file}")
    if BUILD_REQUIREMENTS_FILE.exists():
        print(f"ビルド支援依存関係: {BUILD_REQUIREMENTS_FILE}")

    mode_descriptions = {
        "dev": "dev (コンソールあり)",
        "release": "release (コンソールなし / 本番配布用)",
    }
    print(f"ビルドターゲットモード: {mode_descriptions[mode]}")

    command = build_command(project_root, script_path, exe_name, args.jobs, mode, version)

    env = os.environ.copy()
    env["PYTHONOPTIMIZE"] = "1" if mode in ("release") else "0"

    if args.dry_run:
        print()
        print("=== DRY RUN MODE ===")
        print("生成されるコマンド:")
        print(" ".join(command))
        print("==========================================")
        return 0

    print("コンパイルを実行しています。これには数分かかる場合があります...")
    try:
        result = subprocess.run(command, check=True)
        if result.returncode != 0:
            print("[エラー] Nuitkaコンパイル中に不具合が発生しました。")
            return 1
    except Exception as e:
        print(f"[エラー] ビルド実行中に例外が発生しました: {e}")
        return 1

    output_exe = project_root / f"{exe_name}.exe"
    print("\n==========================================")
    if output_exe.exists():
        print("[SUCCESS] ビルドが正常に完了しました。")
        print(f"保存先: {output_exe}")
        return 0

    print("[エラー] EXEファイルの生成が確認できませんでした。")
    return 1

if __name__ == "__main__":
    main()