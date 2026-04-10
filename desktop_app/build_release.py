from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent
ROOT = THIS_DIR.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from desktop_app.runtime import APP_NAME, APP_VERSION, UPDATER_EXE_NAME


BUILD_ROOT = ROOT / "desktop_release_build"
PYI_WORK_ROOT = BUILD_ROOT / "pyinstaller-work"
PYI_DIST_ROOT = BUILD_ROOT / "pyinstaller-dist"
PYI_SPEC_ROOT = BUILD_ROOT / "pyinstaller-spec"
APP_FOLDER_NAME = "AMS-Assistant-Desktop"
UPDATER_APP_NAME = "AMS-Assistant-Updater"
PREVIEW_ROOT = ROOT / "普通用户体验区-桌面应用版-release预览"
PREVIEW_APP_ROOT = PREVIEW_ROOT / "app"
PREVIEW_USER_DATA_ROOT = PREVIEW_ROOT / "user-data"
GUIDE_SOURCE = ROOT / "desktop_app" / "release_assets" / "应用使用说明.html"
HELP_SOURCE = ROOT / "desktop_app" / "release_assets" / "help"
EXCLUDED_MODULES = [
    "IPython",
    "boto3",
    "botocore",
    "contourpy",
    "h5py",
    "jedi",
    "kiwisolver",
    "matplotlib",
    "ml_dtypes",
    "optree",
    "pandas",
    "parso",
    "prompt_toolkit",
    "scipy",
    "tensorflow",
    "torch",
    "torchaudio",
    "torchvision",
    "tornado",
    "zmq",
]


def run(command: list[str]) -> None:
    print("[RUN]", " ".join(str(part) for part in command))
    subprocess.run(command, check=True, cwd=ROOT)


def remove_if_exists(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def build_pyinstaller_bundle() -> Path:
    remove_if_exists(PYI_WORK_ROOT)
    remove_if_exists(PYI_DIST_ROOT)
    remove_if_exists(PYI_SPEC_ROOT)
    PYI_WORK_ROOT.mkdir(parents=True, exist_ok=True)
    PYI_DIST_ROOT.mkdir(parents=True, exist_ok=True)
    PYI_SPEC_ROOT.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        APP_FOLDER_NAME,
        "--workpath",
        str(PYI_WORK_ROOT),
        "--distpath",
        str(PYI_DIST_ROOT),
        "--specpath",
        str(PYI_SPEC_ROOT),
        "--add-data",
        f"{ROOT / 'maritime-service'};maritime-service",
        "--add-data",
        f"{ROOT / 'desktop_app' / 'release_assets'};desktop_app/release_assets",
        "--add-data",
        f"{ROOT / 'private_pack_source.example'};private_pack_source.example",
        "--collect-all",
        "ttkbootstrap",
        "--collect-all",
        "openpyxl",
        "--collect-all",
        "docx",
        "--collect-all",
        "playwright",
        "--collect-all",
        "requests",
        "--collect-all",
        "cryptography",
        str(ROOT / "launch_ams_desktop_app.py"),
    ]
    for module_name in EXCLUDED_MODULES:
        command.extend(["--exclude-module", module_name])
    run(command)
    return PYI_DIST_ROOT / APP_FOLDER_NAME


def build_pyinstaller_updater() -> Path:
    work_root = PYI_WORK_ROOT / "updater"
    dist_root = PYI_DIST_ROOT / "updater"
    spec_root = PYI_SPEC_ROOT / "updater"
    remove_if_exists(work_root)
    remove_if_exists(dist_root)
    remove_if_exists(spec_root)
    work_root.mkdir(parents=True, exist_ok=True)
    dist_root.mkdir(parents=True, exist_ok=True)
    spec_root.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onefile",
        "--name",
        UPDATER_APP_NAME,
        "--workpath",
        str(work_root),
        "--distpath",
        str(dist_root),
        "--specpath",
        str(spec_root),
        "--add-data",
        f"{ROOT / 'maritime-service'};maritime-service",
        "--add-data",
        f"{ROOT / 'desktop_app' / 'release_assets'};desktop_app/release_assets",
        "--add-data",
        f"{ROOT / 'private_pack_source.example'};private_pack_source.example",
        "--collect-all",
        "ttkbootstrap",
        "--collect-all",
        "openpyxl",
        "--collect-all",
        "docx",
        "--collect-all",
        "playwright",
        "--collect-all",
        "requests",
        "--collect-all",
        "cryptography",
        str(ROOT / "launch_ams_update_helper.py"),
    ]
    for module_name in EXCLUDED_MODULES:
        command.extend(["--exclude-module", module_name])
    run(command)
    return dist_root / f"{UPDATER_APP_NAME}.exe"


def launcher_text() -> str:
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "set \"BASE_DIR=%~dp0\"\r\n"
        "set \"APP_DIR=%BASE_DIR%app\\AMS-Assistant-Desktop\"\r\n"
        "set \"USER_DATA_DIR=%BASE_DIR%user-data\"\r\n"
        "set \"AMS_ASSISTANT_SETTINGS_DIR=%USER_DATA_DIR%\\settings\"\r\n"
        "set \"AMS_ASSISTANT_DEFAULT_WORKSPACE=%USER_DATA_DIR%\\workspace\"\r\n"
        "if not exist \"%USER_DATA_DIR%\\settings\" mkdir \"%USER_DATA_DIR%\\settings\"\r\n"
        "if not exist \"%USER_DATA_DIR%\\workspace\" mkdir \"%USER_DATA_DIR%\\workspace\"\r\n"
        "start \"\" \"%APP_DIR%\\AMS-Assistant-Desktop.exe\"\r\n"
        "exit /b 0\r\n"
    )


def open_workspace_bat() -> str:
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "set \"TARGET=%~dp0user-data\\workspace\"\r\n"
        "if not exist \"%TARGET%\" mkdir \"%TARGET%\"\r\n"
        "start \"\" \"%TARGET%\"\r\n"
        "exit /b 0\r\n"
    )


def open_user_data_bat() -> str:
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "set \"TARGET=%~dp0user-data\"\r\n"
        "if not exist \"%TARGET%\" mkdir \"%TARGET%\"\r\n"
        "start \"\" \"%TARGET%\"\r\n"
        "exit /b 0\r\n"
    )


def open_readme_bat() -> str:
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "start \"\" \"%~dp0README.html\"\r\n"
        "exit /b 0\r\n"
    )


def self_test_bat() -> str:
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "set \"BASE_DIR=%~dp0\"\r\n"
        "set \"APP_DIR=%BASE_DIR%app\\AMS-Assistant-Desktop\"\r\n"
        "set \"USER_DATA_DIR=%BASE_DIR%user-data\"\r\n"
        "set \"AMS_ASSISTANT_SETTINGS_DIR=%USER_DATA_DIR%\\settings\"\r\n"
        "set \"AMS_ASSISTANT_DEFAULT_WORKSPACE=%USER_DATA_DIR%\\workspace\"\r\n"
        "if not exist \"%USER_DATA_DIR%\\settings\" mkdir \"%USER_DATA_DIR%\\settings\"\r\n"
        "if not exist \"%USER_DATA_DIR%\\workspace\" mkdir \"%USER_DATA_DIR%\\workspace\"\r\n"
        "\"%APP_DIR%\\AMS-Assistant-Desktop.exe\" --self-test --self-test-output \"%USER_DATA_DIR%\\self-test\\desktop-self-test.json\" --self-test-workspace \"%USER_DATA_DIR%\\workspace\"\r\n"
        "if errorlevel 1 (\r\n"
        "  echo [ERROR] Self test failed.\r\n"
        "  pause\r\n"
        "  exit /b 1\r\n"
        ")\r\n"
        "start \"\" \"%USER_DATA_DIR%\\self-test\"\r\n"
        "exit /b 0\r\n"
    )


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def proxy_bat(target_name: str) -> str:
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        f"call \"%~dp0{target_name}\"\r\n"
        "exit /b %errorlevel%\r\n"
    )


def copy_preview_assets(bundle_dir: Path, updater_exe: Path) -> None:
    remove_if_exists(PREVIEW_ROOT)
    PREVIEW_APP_ROOT.mkdir(parents=True, exist_ok=True)
    PREVIEW_USER_DATA_ROOT.mkdir(parents=True, exist_ok=True)
    shutil.copytree(bundle_dir, PREVIEW_APP_ROOT / APP_FOLDER_NAME)
    shutil.copy2(updater_exe, PREVIEW_APP_ROOT / APP_FOLDER_NAME / UPDATER_EXE_NAME)
    shutil.copytree(HELP_SOURCE, PREVIEW_ROOT / "help")

    shutil.copy2(GUIDE_SOURCE, PREVIEW_ROOT / "README.html")
    shutil.copy2(GUIDE_SOURCE, PREVIEW_ROOT / "00-从这里开始.html")
    shutil.copy2(GUIDE_SOURCE, PREVIEW_ROOT / "00-先看这里.html")

    write_text(PREVIEW_ROOT / "Start AMS Assistant.bat", launcher_text())
    write_text(PREVIEW_ROOT / "Open Workspace.bat", open_workspace_bat())
    write_text(PREVIEW_ROOT / "Open User Data.bat", open_user_data_bat())
    write_text(PREVIEW_ROOT / "Open Guide.bat", open_readme_bat())
    write_text(PREVIEW_ROOT / "Run Desktop Self Test.bat", self_test_bat())
    write_text(PREVIEW_ROOT / "1-启动AMS桌面应用.bat", proxy_bat("Start AMS Assistant.bat"))
    write_text(PREVIEW_ROOT / "2-运行桌面版自检.bat", proxy_bat("Run Desktop Self Test.bat"))
    write_text(PREVIEW_ROOT / "3-打开工作区.bat", proxy_bat("Open Workspace.bat"))
    write_text(PREVIEW_ROOT / "4-打开说明.bat", proxy_bat("Open Guide.bat"))
    write_text(PREVIEW_ROOT / "5-打开用户数据目录.bat", proxy_bat("Open User Data.bat"))

    write_text(
        PREVIEW_ROOT / "VERSION.txt",
        (
            f"{APP_NAME} Desktop Preview\r\n"
            f"Version: {APP_VERSION}\r\n"
            "\r\n"
            "This folder is a local release preview.\r\n"
            "App binaries are under app\\AMS-Assistant-Desktop.\r\n"
            "User data is stored under user-data.\r\n"
        ),
    )

    write_text(
        PREVIEW_USER_DATA_ROOT / "README.txt",
        (
            "Do not delete this folder if you want to keep your settings, workspace files, "
            "sync tasks, or clearance site session.\r\n"
        ),
    )


def make_zip() -> Path:
    archive_base = BUILD_ROOT / f"AMS-Assistant-Desktop-v{APP_VERSION}"
    remove_if_exists(archive_base.with_suffix(".zip"))
    archive_path = shutil.make_archive(str(archive_base), "zip", PREVIEW_ROOT)
    return Path(archive_path)


def main() -> int:
    print(f"[INFO] Building {APP_NAME} Desktop release preview...")
    bundle_dir = build_pyinstaller_bundle()
    updater_exe = build_pyinstaller_updater()
    copy_preview_assets(bundle_dir, updater_exe)
    archive_path = make_zip()
    print(f"[OK] Preview folder: {PREVIEW_ROOT}")
    print(f"[OK] Preview archive: {archive_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
