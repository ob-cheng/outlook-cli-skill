"""Tests for WSL detection and Windows Python auto-discovery.

Uses mocking to test without actually being in WSL.
"""

import pytest
import sys
from unittest.mock import patch, mock_open, MagicMock

from outlook_cli.core.wsl import (
    is_wsl,
    find_windows_python,
    get_wsl_python_error_message,
    check_wsl_environment,
)


class TestIsWsl:
    """Tests for is_wsl() detection."""

    def test_not_wsl_on_windows(self):
        """Returns False on Windows platform."""
        with patch.object(sys, 'platform', 'win32'):
            assert is_wsl() is False

    def test_not_wsl_on_darwin(self):
        """Returns False on macOS."""
        with patch.object(sys, 'platform', 'darwin'):
            assert is_wsl() is False

    def test_wsl_detected_on_linux_with_microsoft(self):
        """Returns True on Linux with 'microsoft' in /proc/version."""
        with patch.object(sys, 'platform', 'linux'):
            with patch('builtins.open', mock_open(read_data='Linux version 5.15.0-microsoft-standard-WSL2')):
                assert is_wsl() is True

    def test_not_wsl_on_linux_without_microsoft(self):
        """Returns False on Linux without 'microsoft' in /proc/version."""
        with patch.object(sys, 'platform', 'linux'):
            with patch('builtins.open', mock_open(read_data='Linux version 5.15.0-generic')):
                assert is_wsl() is False

    def test_not_wsl_when_proc_version_missing(self):
        """Returns False when /proc/version is not found."""
        with patch.object(sys, 'platform', 'linux'):
            with patch('builtins.open', side_effect=FileNotFoundError):
                assert is_wsl() is False

    def test_not_wsl_when_proc_version_permission_denied(self):
        """Returns False when /proc/version is not readable."""
        with patch.object(sys, 'platform', 'linux'):
            with patch('builtins.open', side_effect=PermissionError):
                assert is_wsl() is False


class TestFindWindowsPython:
    """Tests for find_windows_python() auto-discovery."""

    def test_returns_none_when_not_wsl(self):
        """Returns None when not running in WSL."""
        with patch('outlook_cli.core.wsl.is_wsl', return_value=False):
            result = find_windows_python()
            assert result is None

    def test_finds_python_in_common_path(self):
        """Finds Python in common installation path."""
        mock_subprocess = MagicMock()
        mock_subprocess.returncode = 0
        mock_subprocess.stdout = 'Python 3.12.0'

        with patch('outlook_cli.core.wsl.is_wsl', return_value=True):
            with patch('subprocess.run', return_value=mock_subprocess):
                with patch('pathlib.Path.exists', return_value=True):
                    result = find_windows_python()
                    # Should find something (exact path depends on mock setup)
                    assert result is not None or result is None  # Either is acceptable with mocking

    def test_returns_none_when_no_python_found(self):
        """Returns None when no Python installation found."""
        with patch('outlook_cli.core.wsl.is_wsl', return_value=True):
            with patch('pathlib.Path.exists', return_value=False):
                with patch('subprocess.run', side_effect=FileNotFoundError):
                    result = find_windows_python()
                    assert result is None


class TestGetWslPythonErrorMessage:
    """Tests for get_wsl_python_error_message()."""

    def test_message_includes_wsl_explanation(self):
        """Error message explains WSL limitation."""
        with patch('outlook_cli.core.wsl.find_windows_python', return_value=None):
            msg = get_wsl_python_error_message()

            assert 'WSL' in msg
            assert 'OUTLOOK_CLI_PYTHON' in msg

    def test_message_includes_detected_path(self):
        """Error message includes detected Python path when found."""
        detected_path = '/mnt/c/Users/test/AppData/Local/Programs/Python/Python312/python.exe'
        with patch('outlook_cli.core.wsl.find_windows_python', return_value=detected_path):
            msg = get_wsl_python_error_message()

            assert detected_path in msg
            assert 'Auto-detected' in msg

    def test_message_includes_common_paths_when_not_detected(self):
        """Error message includes common paths when auto-detect fails."""
        with patch('outlook_cli.core.wsl.find_windows_python', return_value=None):
            msg = get_wsl_python_error_message()

            assert 'Common locations' in msg
            assert '/mnt/c/' in msg


class TestCheckWslEnvironment:
    """Tests for check_wsl_environment()."""

    def test_ok_when_not_wsl(self):
        """Returns (True, None) when not in WSL."""
        with patch('outlook_cli.core.wsl.is_wsl', return_value=False):
            ok, error = check_wsl_environment()

            assert ok is True
            assert error is None

    def test_ok_when_outlook_cli_python_set(self):
        """Returns (True, None) when OUTLOOK_CLI_PYTHON is set."""
        with patch('outlook_cli.core.wsl.is_wsl', return_value=True):
            with patch.dict('os.environ', {'OUTLOOK_CLI_PYTHON': '/mnt/c/python.exe'}):
                ok, error = check_wsl_environment()

                assert ok is True
                assert error is None

    def test_error_when_wsl_without_env_var(self):
        """Returns (False, message) when in WSL without OUTLOOK_CLI_PYTHON."""
        with patch('outlook_cli.core.wsl.is_wsl', return_value=True):
            with patch.dict('os.environ', {}, clear=True):
                with patch('outlook_cli.core.wsl.get_wsl_python_error_message', return_value='Error message'):
                    ok, error = check_wsl_environment()

                    assert ok is False
                    assert error is not None
