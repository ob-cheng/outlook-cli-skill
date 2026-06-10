"""WSL detection and Windows Python auto-discovery."""

import os
import subprocess
import sys
from pathlib import Path


def is_wsl() -> bool:
    """Check if running under Windows Subsystem for Linux."""
    if sys.platform != 'linux':
        return False
    try:
        with open('/proc/version', 'r') as f:
            return 'microsoft' in f.read().lower()
    except (FileNotFoundError, PermissionError):
        return False


def find_windows_python() -> str | None:
    """Auto-detect Windows Python executable from WSL.

    Probes common installation paths and returns the first working one.

    Returns:
        Path to Windows python.exe, or None if not found.
    """
    if not is_wsl():
        return None

    # Common Windows Python installation paths
    common_paths = []

    # Try to get the Windows username for user-specific paths
    try:
        result = subprocess.run(
            ['cmd.exe', '/c', 'echo %USERNAME%'],
            capture_output=True, text=True, timeout=5
        )
        win_user = result.stdout.strip()
        if win_user and win_user != '%USERNAME%':
            # User-specific Python installations (most common)
            for ver in ['313', '312', '311', '310', '39']:
                common_paths.append(f'/mnt/c/Users/{win_user}/AppData/Local/Programs/Python/Python{ver}/python.exe')
            # Also check Microsoft Store Python
            common_paths.append(f'/mnt/c/Users/{win_user}/AppData/Local/Microsoft/WindowsApps/python.exe')
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    # System-wide installations
    for ver in ['313', '312', '311', '310', '39']:
        common_paths.append(f'/mnt/c/Program Files/Python{ver}/python.exe')
        common_paths.append(f'/mnt/c/Python{ver}/python.exe')

    # Check each path
    for path in common_paths:
        if Path(path).exists():
            # Verify it actually works
            try:
                result = subprocess.run(
                    [path, '--version'],
                    capture_output=True, text=True, timeout=5
                )
                if result.returncode == 0 and 'Python' in result.stdout:
                    return path
            except (subprocess.TimeoutExpired, FileNotFoundError, PermissionError):
                continue

    # Fallback: try 'where python' via cmd.exe
    try:
        result = subprocess.run(
            ['cmd.exe', '/c', 'where python'],
            capture_output=True, text=True, timeout=5
        )
        if result.returncode == 0:
            # Convert Windows path to WSL path
            win_path = result.stdout.strip().split('\n')[0]
            if win_path:
                # C:\... -> /mnt/c/...
                wsl_path = '/mnt/' + win_path[0].lower() + win_path[2:].replace('\\', '/')
                if Path(wsl_path).exists():
                    return wsl_path
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    return None


def get_wsl_python_error_message() -> str:
    """Generate a helpful error message for WSL users without OUTLOOK_CLI_PYTHON set."""
    detected = find_windows_python()

    msg = """
Running from WSL but OUTLOOK_CLI_PYTHON is not set.

WSL's Python cannot use COM automation (pywin32). You need to use Windows Python.
"""

    if detected:
        msg += f"""
Auto-detected Windows Python at:
  {detected}

Quick fix (add to your shell profile):
  export OUTLOOK_CLI_PYTHON="{detected}"

Or for Hermes, add to ~/.hermes/config.yaml:
  env:
    OUTLOOK_CLI_PYTHON: "{detected}"

Or for Claude Code, add to .claude/settings.json:
  {{"env": {{"OUTLOOK_CLI_PYTHON": "{detected}"}}}}
"""
    else:
        msg += """
Could not auto-detect Windows Python. Common locations:
  /mnt/c/Users/<username>/AppData/Local/Programs/Python/Python312/python.exe
  /mnt/c/Program Files/Python312/python.exe

Find your Windows Python path and set OUTLOOK_CLI_PYTHON:
  export OUTLOOK_CLI_PYTHON="/mnt/c/path/to/python.exe"
"""

    return msg.strip()


def check_wsl_environment() -> tuple[bool, str | None]:
    """Check if WSL environment is properly configured.

    Returns:
        (ok, error_message): ok=True if good to proceed, otherwise error_message explains the issue.
    """
    if not is_wsl():
        return True, None

    # Check if OUTLOOK_CLI_PYTHON is set
    if os.environ.get('OUTLOOK_CLI_PYTHON'):
        return True, None

    # WSL without OUTLOOK_CLI_PYTHON - this won't work
    return False, get_wsl_python_error_message()
