"""
Makerspace Card Scanner - Installer Build Script
=================================================
Automates assembling the installer package:
  1. Downloads Python 3.10 embeddable zip
  2. Enables pip and installs runtime dependencies
  3. Copies application source files and assets
  4. Writes the .version file
  5. Optionally compiles the Inno Setup installer

Prerequisites:
  - Python 3.x (to run THIS script)
  - Internet access (to download Python + packages)
  - Inno Setup 6 installed (optional, for compiling the installer .exe)

Usage:
    python installer/build_installer.py                 (build everything)
    python installer/build_installer.py --skip-inno     (skip Inno Setup compile)
    python installer/build_installer.py --clean         (remove build dir first)
    python installer/build_installer.py --python-ver 3.10.11  (specific version)
"""

import os
import sys
import shutil
import zipfile
import subprocess
import argparse
import hashlib
from urllib.request import urlopen, Request
from urllib.error import URLError

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)

BUILD_DIR = os.path.join(SCRIPT_DIR, "build")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")
PYTHON_DIR = os.path.join(BUILD_DIR, "python")

DEFAULT_PYTHON_VERSION = "3.10.11"

# Files to include from the project root
APP_SOURCE_FILES = [
    "MakerspaceSignInTablet.py",
    "CardReaderMakerspace.py",
    "database.py",
    "database_sync.py",
    "excel_db_sync.py",
    "excel_utils.py",
    "bridge_api.py",
    "config_examples.py",
    "auto_updater.py",
    "fetch_missing_training.py",
    "fetch_missing_training.bat",
    "MakerspaceScanner.bat",
]

APP_ASSET_FILES = [
    "BackgroundTablet.png",
    "BackgroundWatt.png",
    "BackgroundAdobe.png",
    "MakerspaceLogoIcon.ico",
]

REQUIREMENTS_FILE = os.path.join(SCRIPT_DIR, "requirements-installer.txt")

# Inno Setup compiler -- common install locations
ISCC_PATHS = [
    r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    r"C:\Program Files\Inno Setup 6\ISCC.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe"),
    r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def banner(msg):
    print()
    print("=" * 60)
    print(f"  {msg}")
    print("=" * 60)


def download_file(url, dest_path):
    """Download a file from url to dest_path with progress indication."""
    print(f"  Downloading: {url}")
    print(f"  Destination: {dest_path}")
    req = Request(url, headers={"User-Agent": "MakerspaceInstaller/1.0"})
    with urlopen(req, timeout=120) as resp:
        total = resp.headers.get("Content-Length")
        total = int(total) if total else None
        downloaded = 0
        with open(dest_path, "wb") as fh:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                fh.write(chunk)
                downloaded += len(chunk)
                if total:
                    pct = downloaded * 100 // total
                    print(f"\r  Progress: {pct}% ({downloaded:,} / {total:,} bytes)", end="", flush=True)
        print()
    return dest_path


def get_git_head_sha():
    """Get the current HEAD commit SHA from git."""
    try:
        result = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=PROJECT_ROOT,
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "unknown"


def find_iscc():
    """Locate the Inno Setup compiler."""
    for path in ISCC_PATHS:
        if os.path.exists(path):
            return path
    # Try PATH
    try:
        result = subprocess.run(
            ["where", "ISCC.exe"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            return result.stdout.strip().splitlines()[0]
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Build steps
# ---------------------------------------------------------------------------

def step_clean(build_dir):
    """Remove existing build directory."""
    if os.path.exists(build_dir):
        print(f"  Removing existing build directory: {build_dir}")
        shutil.rmtree(build_dir)
    os.makedirs(build_dir, exist_ok=True)


def step_download_python(python_version, build_dir):
    """Download and extract the Python embeddable zip."""
    python_dir = os.path.join(build_dir, "python")

    if os.path.exists(python_dir) and os.path.exists(os.path.join(python_dir, "python.exe")):
        print("  Embedded Python already present. Skipping download.")
        return python_dir

    os.makedirs(python_dir, exist_ok=True)

    # Build download URL for the Windows AMD64 embeddable zip
    zip_name = f"python-{python_version}-embed-amd64.zip"
    url = f"https://www.python.org/ftp/python/{python_version}/{zip_name}"

    zip_path = os.path.join(build_dir, zip_name)
    download_file(url, zip_path)

    print(f"  Extracting to {python_dir}...")
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(python_dir)

    os.remove(zip_path)
    print(f"  Python {python_version} extracted.")
    return python_dir


def step_enable_pip(python_dir, python_version):
    """Enable pip in the embeddable Python by modifying ._pth and installing pip."""
    # Determine the _pth file name (e.g., python310._pth for 3.10.x)
    major_minor = python_version.rsplit(".", 1)[0].replace(".", "")  # "310"
    pth_file = os.path.join(python_dir, f"python{major_minor}._pth")

    if not os.path.exists(pth_file):
        print(f"  WARNING: {pth_file} not found. Trying to find ._pth file...")
        for f in os.listdir(python_dir):
            if f.endswith("._pth"):
                pth_file = os.path.join(python_dir, f)
                print(f"  Found: {pth_file}")
                break
        else:
            print("  ERROR: No ._pth file found!")
            return False

    # Uncomment 'import site' in the ._pth file
    print(f"  Enabling pip: modifying {os.path.basename(pth_file)}...")
    with open(pth_file, "r") as fh:
        lines = fh.readlines()

    with open(pth_file, "w") as fh:
        for line in lines:
            if line.strip() == "#import site":
                fh.write("import site\n")
                print("    Uncommented 'import site'")
            else:
                fh.write(line)
        # Ensure Lib/ and Lib/site-packages are on the path
        fh.write("Lib\n")
        fh.write("Lib\\site-packages\n")
        # Add parent directory so app .py files are importable
        # (python/ is a subdirectory of the install dir where .py files live)
        fh.write("..\n")

    # Download get-pip.py
    get_pip_path = os.path.join(python_dir, "get-pip.py")
    if not os.path.exists(get_pip_path):
        download_file("https://bootstrap.pypa.io/get-pip.py", get_pip_path)

    # Run get-pip.py
    python_exe = os.path.join(python_dir, "python.exe")
    print("  Installing pip...")
    result = subprocess.run(
        [python_exe, get_pip_path, "--no-warn-script-location"],
        cwd=python_dir,
        capture_output=True, text=True, timeout=300
    )
    if result.returncode != 0:
        print(f"  ERROR installing pip:\n{result.stderr}")
        return False

    # Clean up get-pip.py
    os.remove(get_pip_path)
    print("  pip installed successfully.")
    return True


def step_install_deps(python_dir, requirements_file):
    """Install dependencies using the embedded pip."""
    python_exe = os.path.join(python_dir, "python.exe")
    print(f"  Installing dependencies from {os.path.basename(requirements_file)}...")

    result = subprocess.run(
        [python_exe, "-m", "pip", "install",
         "--no-warn-script-location",
         "--disable-pip-version-check",
         "-r", requirements_file],
        cwd=python_dir,
        capture_output=True, text=True, timeout=600
    )

    if result.returncode != 0:
        print(f"  ERROR installing dependencies:\n{result.stderr}")
        return False

    if result.stdout:
        # Show just the summary
        lines = result.stdout.strip().splitlines()
        for line in lines[-5:]:
            print(f"    {line}")

    print("  Dependencies installed successfully.")

    # Clean pip cache to save space
    subprocess.run(
        [python_exe, "-m", "pip", "cache", "purge"],
        capture_output=True, timeout=60
    )

    return True


def step_copy_tkinter(python_dir):
    """
    Copy tkinter from the system Python into the embedded Python.
    The embeddable zip doesn't include tkinter, but customtkinter requires it.
    """
    import tkinter
    sys_python_prefix = sys.prefix

    # Source locations
    tkinter_lib_src = os.path.dirname(tkinter.__file__)  # .../lib/tkinter/
    dlls_dir = os.path.join(sys_python_prefix, "DLLs")
    tcl_dir = os.path.join(sys_python_prefix, "tcl")

    # Destination locations
    lib_dir = os.path.join(python_dir, "Lib")
    os.makedirs(lib_dir, exist_ok=True)

    # 1. Copy tkinter package (Lib/tkinter/)
    tkinter_dst = os.path.join(lib_dir, "tkinter")
    if os.path.exists(tkinter_dst):
        print("  tkinter already present. Skipping.")
        return True

    print(f"  Copying tkinter from {tkinter_lib_src}")
    shutil.copytree(tkinter_lib_src, tkinter_dst)

    # 2. Copy _tkinter.pyd and Tcl/Tk DLLs
    dll_files = ["_tkinter.pyd", "tcl86t.dll", "tk86t.dll"]
    for dll_name in dll_files:
        src = os.path.join(dlls_dir, dll_name)
        dst = os.path.join(python_dir, dll_name)
        if os.path.exists(src) and not os.path.exists(dst):
            shutil.copy2(src, dst)
            print(f"    Copied {dll_name}")
        elif not os.path.exists(src):
            print(f"    WARNING: {dll_name} not found at {src}")

    # 3. Copy tcl/ directory (tcl8.6 and tk8.6 data files)
    tcl_dst = os.path.join(python_dir, "tcl")
    if os.path.exists(tcl_dir) and not os.path.exists(tcl_dst):
        print(f"  Copying Tcl/Tk data files from {tcl_dir}")
        shutil.copytree(tcl_dir, tcl_dst)

    # 4. Set TCL/TK environment in the ._pth file
    # (The launcher batch file sets PYTHONPATH, but we also need TCL_LIBRARY/TK_LIBRARY)
    print("  tkinter copied successfully.")
    return True


def step_copy_app_files(build_dir, project_root):
    """Copy application source files and assets to the build directory."""
    print("  Copying application files...")
    copied = 0

    for fname in APP_SOURCE_FILES:
        src = os.path.join(project_root, fname)
        dst = os.path.join(build_dir, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            copied += 1
        else:
            print(f"    WARNING: {fname} not found in project root")

    for fname in APP_ASSET_FILES:
        src = os.path.join(project_root, fname)
        dst = os.path.join(build_dir, fname)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            copied += 1
        else:
            print(f"    NOTICE: {fname} not found (optional asset)")

    print(f"  Copied {copied} files.")
    return copied > 0


def step_write_version(build_dir):
    """Write the .version file with the current git commit SHA."""
    sha = get_git_head_sha()
    version_file = os.path.join(build_dir, ".version")
    with open(version_file, "w") as fh:
        fh.write(sha + "\n")
    print(f"  Version file written: {sha[:12]}")
    return sha


def step_compile_inno(iss_file, output_dir):
    """Compile the Inno Setup script into an installer .exe."""
    iscc = find_iscc()
    if not iscc:
        print("  WARNING: Inno Setup compiler (ISCC.exe) not found.")
        print("  Install Inno Setup 6 from https://jrsoftware.org/isinfo.php")
        print("  The build directory is ready for manual compilation.")
        return False

    os.makedirs(output_dir, exist_ok=True)

    print(f"  Compiling with: {iscc}")
    result = subprocess.run(
        [iscc, f"/O{output_dir}", iss_file],
        capture_output=True, text=True, timeout=300
    )

    if result.returncode != 0:
        print(f"  ERROR compiling installer:\n{result.stderr}")
        return False

    print("  Installer compiled successfully!")
    # Find the output file
    for f in os.listdir(output_dir):
        if f.endswith(".exe"):
            full = os.path.join(output_dir, f)
            size_mb = os.path.getsize(full) / (1024 * 1024)
            print(f"  Output: {full} ({size_mb:.1f} MB)")
    return True


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Build the Makerspace Card Scanner installer."
    )
    parser.add_argument(
        "--clean", action="store_true",
        help="Remove existing build directory before starting."
    )
    parser.add_argument(
        "--skip-inno", action="store_true",
        help="Skip Inno Setup compilation (just prepare the build dir)."
    )
    parser.add_argument(
        "--python-ver", default=DEFAULT_PYTHON_VERSION,
        help=f"Python version to embed (default: {DEFAULT_PYTHON_VERSION})."
    )
    args = parser.parse_args()

    banner("MAKERSPACE CARD SCANNER - INSTALLER BUILD")
    print(f"  Project root:    {PROJECT_ROOT}")
    print(f"  Build directory: {BUILD_DIR}")
    print(f"  Python version:  {args.python_ver}")
    print()

    # Step 0: Clean
    if args.clean:
        banner("Step 0/7: Cleaning build directory")
        step_clean(BUILD_DIR)
    else:
        os.makedirs(BUILD_DIR, exist_ok=True)

    # Step 1: Download embedded Python
    banner("Step 1/7: Downloading embedded Python")
    python_dir = step_download_python(args.python_ver, BUILD_DIR)

    # Step 2: Enable pip
    banner("Step 2/7: Enabling pip")
    if not step_enable_pip(python_dir, args.python_ver):
        print("FATAL: Could not enable pip. Aborting.")
        sys.exit(1)

    # Step 3: Install dependencies
    banner("Step 3/7: Installing dependencies")
    if not step_install_deps(python_dir, REQUIREMENTS_FILE):
        print("FATAL: Could not install dependencies. Aborting.")
        sys.exit(1)

    # Step 4: Copy tkinter (not included in embeddable Python)
    banner("Step 4/7: Copying tkinter from system Python")
    if not step_copy_tkinter(python_dir):
        print("WARNING: tkinter copy failed. GUI may not work.")

    # Step 5: Copy app files
    banner("Step 5/7: Copying application files")
    if not step_copy_app_files(BUILD_DIR, PROJECT_ROOT):
        print("FATAL: No application files copied. Aborting.")
        sys.exit(1)

    # Step 6: Write version
    banner("Step 6/7: Writing version file")
    step_write_version(BUILD_DIR)

    # Step 7: Compile Inno Setup
    if not args.skip_inno:
        banner("Step 7/7: Compiling Inno Setup installer")
        iss_file = os.path.join(SCRIPT_DIR, "setup.iss")
        if not os.path.exists(iss_file):
            print(f"  ERROR: {iss_file} not found.")
            print("  Build directory is ready for manual compilation.")
        else:
            step_compile_inno(iss_file, OUTPUT_DIR)
    else:
        print("\n  Step 7 skipped (--skip-inno).")

    # Summary
    banner("BUILD COMPLETE")
    print(f"  Build directory: {BUILD_DIR}")
    if os.path.exists(OUTPUT_DIR):
        for f in os.listdir(OUTPUT_DIR):
            if f.endswith(".exe"):
                print(f"  Installer:       {os.path.join(OUTPUT_DIR, f)}")
    print()
    print("  To test the build without installing:")
    print(f'    cd "{BUILD_DIR}"')
    print(f"    python\\python.exe MakerspaceSignInTablet.py")
    print()


if __name__ == "__main__":
    main()
