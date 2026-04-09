"""
Makerspace Card Scanner - Auto Updater
=======================================
Checks the GitHub repository for updates and downloads changed source files.
Designed to run as a Windows Scheduled Task at midnight.

Protected files (NEVER overwritten):
    config.py, *.db, *.xlsx, backups/, python/, update.log

Usage:
    python auto_updater.py              (normal update)
    python auto_updater.py --check      (check only, no download)
    python auto_updater.py --force      (re-download all updatable files)
    python auto_updater.py --verbose    (extra detail)
"""

import os
import sys
import json
import time
import shutil
import hashlib
import argparse
import tempfile
from datetime import datetime

# Use urllib from stdlib so the updater works even if requests isn't installed
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GITHUB_OWNER = "clemsonMakerspace"
GITHUB_REPO = "Makerspace-CardScanner"
GITHUB_BRANCH = "main"

GITHUB_API_BASE = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
GITHUB_RAW_BASE = f"https://raw.githubusercontent.com/{GITHUB_OWNER}/{GITHUB_REPO}/{GITHUB_BRANCH}"

# Resolve paths relative to THIS script's directory (the install dir)
INSTALL_DIR = os.path.dirname(os.path.abspath(__file__))
VERSION_FILE = os.path.join(INSTALL_DIR, ".version")
LOG_FILE = os.path.join(INSTALL_DIR, "update.log")

# Maximum log file size before rotation (1 MB)
MAX_LOG_SIZE = 1_048_576

# HTTP request timeout in seconds
HTTP_TIMEOUT = 30

# Files and directories that must NEVER be overwritten by updates
PROTECTED_PATTERNS = {
    # Exact filenames (case-insensitive comparison)
    "config.py",
    "update.log",
    ".version",
}

PROTECTED_EXTENSIONS = {
    ".db",
    ".xlsx",
}

PROTECTED_DIRS = {
    "python",
    "backups",
    "_internal",
    "__pycache__",
    ".git",
    "installer",
}

# Only these file extensions are eligible for update
UPDATABLE_EXTENSIONS = {
    ".py",
    ".bat",
    ".png",
    ".ico",
    ".pdn",
    ".md",
    ".txt",
}


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

_log_lines = []


def log(message, level="INFO"):
    """Log a message to both stdout and the log buffer."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] [{level}] {message}"
    print(line)
    _log_lines.append(line)


def flush_log():
    """Append buffered log lines to the persistent log file."""
    if not _log_lines:
        return
    try:
        # Rotate log if it's too large
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > MAX_LOG_SIZE:
            rotated = LOG_FILE + ".old"
            if os.path.exists(rotated):
                os.remove(rotated)
            os.rename(LOG_FILE, rotated)

        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            for line in _log_lines:
                fh.write(line + "\n")
            fh.write("\n")
    except Exception as exc:
        print(f"WARNING: Could not write to log file: {exc}")


# ---------------------------------------------------------------------------
# HTTP helpers (stdlib only -- no requests dependency)
# ---------------------------------------------------------------------------

def _github_get(url, accept="application/vnd.github.v3+json"):
    """Make a GET request to the GitHub API. Returns parsed JSON or bytes."""
    headers = {
        "Accept": accept,
        "User-Agent": f"MakerspaceCardScanner-AutoUpdater/1.0",
    }

    # Optional: use a GitHub token if available (avoids 60 req/hr rate limit)
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        headers["Authorization"] = f"token {token}"

    req = Request(url, headers=headers)
    with urlopen(req, timeout=HTTP_TIMEOUT) as resp:
        data = resp.read()
        content_type = resp.headers.get("Content-Type", "")
        if "json" in content_type or "json" in accept:
            return json.loads(data.decode("utf-8"))
        return data


def _download_raw(path):
    """Download a single file from the raw GitHub content URL."""
    url = f"{GITHUB_RAW_BASE}/{path}"
    req = Request(url, headers={"User-Agent": "MakerspaceCardScanner-AutoUpdater/1.0"})
    with urlopen(req, timeout=HTTP_TIMEOUT) as resp:
        return resp.read()


# ---------------------------------------------------------------------------
# Version tracking
# ---------------------------------------------------------------------------

def get_local_version():
    """Read the stored commit SHA from .version file."""
    if os.path.exists(VERSION_FILE):
        try:
            with open(VERSION_FILE, "r") as fh:
                return fh.read().strip()
        except Exception:
            pass
    return None


def set_local_version(sha):
    """Write the commit SHA to .version file."""
    with open(VERSION_FILE, "w") as fh:
        fh.write(sha + "\n")


def get_remote_version():
    """Get the latest commit SHA on the target branch from GitHub API."""
    url = f"{GITHUB_API_BASE}/commits/{GITHUB_BRANCH}"
    data = _github_get(url)
    return data["sha"]


# ---------------------------------------------------------------------------
# File protection logic
# ---------------------------------------------------------------------------

def is_protected(filepath):
    """
    Check whether a file path should be protected from updates.
    filepath is relative to the install directory (e.g. "config.py", "backups/foo.xlsx").
    """
    name = os.path.basename(filepath).lower()
    _, ext = os.path.splitext(name)

    # Check exact filename matches
    if name in {p.lower() for p in PROTECTED_PATTERNS}:
        return True

    # Check extension
    if ext in PROTECTED_EXTENSIONS:
        return True

    # Check if file is inside a protected directory
    parts = filepath.replace("\\", "/").split("/")
    for part in parts[:-1]:  # all directory components except the filename
        if part.lower() in {d.lower() for d in PROTECTED_DIRS}:
            return True

    return False


def is_updatable(filepath):
    """Check whether a file's extension makes it eligible for updates."""
    _, ext = os.path.splitext(filepath.lower())
    return ext in UPDATABLE_EXTENSIONS


# ---------------------------------------------------------------------------
# Tree comparison
# ---------------------------------------------------------------------------

def get_remote_tree():
    """
    Fetch the full file tree for the target branch.
    Returns a list of dicts with 'path', 'sha', 'size', 'type' keys.
    """
    url = f"{GITHUB_API_BASE}/git/trees/{GITHUB_BRANCH}?recursive=1"
    data = _github_get(url)
    return [
        item for item in data.get("tree", [])
        if item["type"] == "blob"  # files only, not tree entries
    ]


def get_local_file_sha(filepath):
    """
    Compute the git-compatible SHA-1 for a local file.
    Git SHA = SHA1("blob <size>\0<content>")
    """
    abs_path = os.path.join(INSTALL_DIR, filepath)
    if not os.path.exists(abs_path):
        return None
    try:
        with open(abs_path, "rb") as fh:
            content = fh.read()
        header = f"blob {len(content)}\0".encode("utf-8")
        return hashlib.sha1(header + content).hexdigest()
    except Exception:
        return None


def find_changed_files(remote_tree, force=False):
    """
    Compare remote tree with local files.
    Returns list of file paths that need updating.
    """
    changed = []
    for item in remote_tree:
        path = item["path"]

        # Skip files that aren't updatable
        if not is_updatable(path):
            continue

        # Skip protected files
        if is_protected(path):
            continue

        if force:
            changed.append(path)
            continue

        # Compare SHA
        local_sha = get_local_file_sha(path)
        if local_sha != item["sha"]:
            changed.append(path)

    return changed


# ---------------------------------------------------------------------------
# Update execution
# ---------------------------------------------------------------------------

def download_and_replace(file_path):
    """
    Download a file from GitHub and write it to the install directory.
    Uses a temp file + rename for atomic replacement.
    """
    abs_path = os.path.join(INSTALL_DIR, file_path)

    # Ensure parent directory exists
    parent = os.path.dirname(abs_path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)

    # Download to a temp file in the same directory (same filesystem for rename)
    fd, tmp_path = tempfile.mkstemp(dir=parent, prefix=".update_", suffix=".tmp")
    try:
        content = _download_raw(file_path)
        os.write(fd, content)
        os.close(fd)
        fd = None

        # Atomic replace (on Windows, need to remove target first if it exists)
        if os.path.exists(abs_path):
            # Keep a backup just in case
            backup_path = abs_path + ".update_backup"
            try:
                if os.path.exists(backup_path):
                    os.remove(backup_path)
                shutil.copy2(abs_path, backup_path)
            except Exception:
                pass
            os.remove(abs_path)

        os.rename(tmp_path, abs_path)
        return True

    except Exception as exc:
        log(f"  FAILED to update {file_path}: {exc}", "ERROR")
        # Clean up temp file
        try:
            if fd is not None:
                os.close(fd)
        except Exception:
            pass
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False


def cleanup_backup_files():
    """Remove .update_backup files left from successful updates."""
    for root, dirs, files in os.walk(INSTALL_DIR):
        # Don't walk into protected directories
        dirs[:] = [d for d in dirs if d.lower() not in {p.lower() for p in PROTECTED_DIRS}]
        for fname in files:
            if fname.endswith(".update_backup"):
                try:
                    os.remove(os.path.join(root, fname))
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# Main update flow
# ---------------------------------------------------------------------------

def run_update(check_only=False, force=False, verbose=False):
    """
    Main update sequence.
    Returns True if updates were applied (or available if check_only).
    """
    log("=" * 60)
    log("Makerspace Card Scanner - Auto Updater")
    log(f"Install directory: {INSTALL_DIR}")
    log(f"Mode: {'check-only' if check_only else 'force' if force else 'normal'}")

    # Step 1: Get local version
    local_sha = get_local_version()
    log(f"Local version:  {local_sha or '(none)'}")

    # Step 2: Get remote version
    try:
        remote_sha = get_remote_version()
    except (URLError, HTTPError) as exc:
        log(f"Cannot reach GitHub: {exc}", "ERROR")
        log("Update check skipped. Will retry next run.")
        return False

    log(f"Remote version: {remote_sha}")

    # Step 3: Compare
    if local_sha == remote_sha and not force:
        log("Already up to date. No changes needed.")
        return False

    if local_sha != remote_sha:
        log("New version available on GitHub.")
    elif force:
        log("Force mode: re-checking all files regardless of version.")

    # Step 4: Get file tree and find changes
    log("Fetching file tree from GitHub...")
    try:
        remote_tree = get_remote_tree()
    except (URLError, HTTPError) as exc:
        log(f"Cannot fetch file tree: {exc}", "ERROR")
        return False

    log(f"Remote tree contains {len(remote_tree)} files.")

    changed = find_changed_files(remote_tree, force=force)

    if not changed:
        log("No updatable files have changed.")
        set_local_version(remote_sha)
        return False

    log(f"Found {len(changed)} file(s) to update:")
    for path in changed:
        log(f"  -> {path}")

    if check_only:
        log("Check-only mode. No files downloaded.")
        return True

    # Step 5: Download and replace
    log("Downloading updates...")
    success_count = 0
    fail_count = 0

    for path in changed:
        if verbose:
            log(f"  Downloading: {path}")
        if download_and_replace(path):
            success_count += 1
        else:
            fail_count += 1

    log(f"Update complete: {success_count} updated, {fail_count} failed.")

    # Step 6: Update version tracker
    if fail_count == 0:
        set_local_version(remote_sha)
        log(f"Version updated to {remote_sha[:12]}")
        cleanup_backup_files()
    else:
        log("Some files failed to update. Version NOT advanced.", "WARN")
        log("Backup files (.update_backup) preserved for recovery.", "WARN")

    return True


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Makerspace Card Scanner auto-updater."
    )
    parser.add_argument(
        "--check", action="store_true",
        help="Check for updates without downloading."
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Re-download all updatable files regardless of version."
    )
    parser.add_argument(
        "--verbose", action="store_true",
        help="Show extra detail during update."
    )
    args = parser.parse_args()

    try:
        run_update(
            check_only=args.check,
            force=args.force,
            verbose=args.verbose,
        )
    except Exception as exc:
        log(f"Unhandled error: {exc}", "ERROR")
        import traceback
        log(traceback.format_exc(), "ERROR")
    finally:
        flush_log()


if __name__ == "__main__":
    main()
