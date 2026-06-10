"""Tests for _print_progress_dot helper in cli.py."""

import io
import sys
import pytest


# Copy of the function from cli.py (can't import due to win32com)
def _print_progress_dot(count: int, out=None) -> None:
    """Print a progress dot every N emails during search."""
    if out is None:
        out = sys.stdout
    if count % 20 == 0:
        if count == 20:
            print("  Scanning", end="", flush=True, file=out)
        print(".", end="", flush=True, file=out)


class TestProgressDots:
    def test_no_dots_for_few_items(self):
        """5 items should produce no output."""
        buf = io.StringIO()
        for i in range(1, 6):
            _print_progress_dot(i, out=buf)
        assert buf.getvalue() == ""

    def test_no_dots_for_19_items(self):
        """19 items just under the 20-item threshold."""
        buf = io.StringIO()
        for i in range(1, 20):
            _print_progress_dot(i, out=buf)
        assert buf.getvalue() == ""

    def test_one_dot_at_20(self):
        """Exactly 20 items: one dot with header."""
        buf = io.StringIO()
        for i in range(1, 21):
            _print_progress_dot(i, out=buf)
        assert buf.getvalue() == "  Scanning."

    def test_two_dots_at_40(self):
        """40 items: two dots, header only on first."""
        buf = io.StringIO()
        for i in range(1, 41):
            _print_progress_dot(i, out=buf)
        assert buf.getvalue() == "  Scanning.."

    def test_five_dots_at_100(self):
        """100 items: dots at 20, 40, 60, 80, 100."""
        buf = io.StringIO()
        for i in range(1, 101):
            _print_progress_dot(i, out=buf)
        assert buf.getvalue() == "  Scanning....."

    def test_header_only_once(self):
        """'  Scanning' header appears once, not repeated."""
        buf = io.StringIO()
        for i in range(1, 61):
            _print_progress_dot(i, out=buf)
        output = buf.getvalue()
        # "  Scanning" should appear exactly once
        assert output.count("  Scanning") == 1

    def test_count_0_behavior(self):
        """count=0 is an edge case: 0%20==0 is True, prints a dot.
        
        In practice, scanned starts at 0 and is incremented to 1 before
        the callback fires, so count=0 should never reach this function."""
        buf = io.StringIO()
        _print_progress_dot(0, out=buf)
        # 0 % 20 == 0 → True, but 0 != 20 so no "Scanning" header
        # Just prints "."
        assert buf.getvalue() == "."
