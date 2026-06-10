"""Tests for formatting utilities — email parsing, date handling, sanitization.

These are pure functions with no COM dependency.
"""

import pytest
from datetime import datetime

from outlook_cli.utils.formatting import (
    extract_email_address,
    extract_all_email_addresses,
    extract_display_name,
    format_outlook_date,
    sanitize_filename,
    normalize_subject,
    parse_date,
)


class TestExtractEmailAddress:
    """Tests for extracting a single email address."""

    def test_simple_email(self):
        """Extracts simple email address."""
        assert extract_email_address("alice@example.com") == "alice@example.com"

    def test_email_with_name(self):
        """Extracts email from 'Name <email>' format."""
        result = extract_email_address("Alice Smith <alice@example.com>")
        assert result == "alice@example.com"

    def test_email_with_quotes(self):
        """Extracts email from quoted name format."""
        result = extract_email_address('"Smith, Alice" <alice@example.com>')
        assert result == "alice@example.com"

    def test_email_lowercased(self):
        """Returns email in lowercase."""
        result = extract_email_address("Alice@Example.COM")
        assert result == "alice@example.com"

    def test_email_with_subdomain(self):
        """Extracts email with subdomain."""
        result = extract_email_address("alice@mail.example.com")
        assert result == "alice@mail.example.com"

    def test_email_with_dots_in_local(self):
        """Extracts email with dots in local part."""
        result = extract_email_address("alice.smith@example.com")
        assert result == "alice.smith@example.com"

    def test_empty_string(self):
        """Returns None for empty string."""
        assert extract_email_address("") is None

    def test_none_input(self):
        """Returns None for None input."""
        assert extract_email_address(None) is None

    def test_no_email_in_string(self):
        """Returns None when no email found."""
        assert extract_email_address("Alice Smith") is None

    def test_exchange_format(self):
        """Extracts email from Exchange X.500 path."""
        result = extract_email_address("/O=COMPANY/OU=USERS/CN=alice@example.com")
        assert result == "alice@example.com"


class TestExtractAllEmailAddresses:
    """Tests for extracting multiple email addresses."""

    def test_single_email(self):
        """Extracts single email."""
        result = extract_all_email_addresses("alice@example.com")
        assert result == ["alice@example.com"]

    def test_multiple_emails_semicolon(self):
        """Extracts multiple emails separated by semicolon."""
        result = extract_all_email_addresses("alice@example.com; bob@example.com")
        assert result == ["alice@example.com", "bob@example.com"]

    def test_multiple_emails_with_names(self):
        """Extracts emails from 'Name <email>' format."""
        result = extract_all_email_addresses(
            "Alice <alice@example.com>; Bob <bob@example.com>"
        )
        assert result == ["alice@example.com", "bob@example.com"]

    def test_empty_string(self):
        """Returns empty list for empty string."""
        assert extract_all_email_addresses("") == []

    def test_none_input(self):
        """Returns empty list for None input."""
        assert extract_all_email_addresses(None) == []

    def test_all_lowercased(self):
        """All emails returned in lowercase."""
        result = extract_all_email_addresses("Alice@EXAMPLE.com; BOB@example.COM")
        assert all(e == e.lower() for e in result)


class TestExtractDisplayName:
    """Tests for extracting display names from addresses."""

    def test_simple_name(self):
        """Extracts simple name."""
        result = extract_display_name("Alice Smith")
        assert result == "Alice Smith"

    def test_name_with_email(self):
        """Extracts name from 'Name <email>' format."""
        result = extract_display_name("Alice Smith <alice@example.com>")
        assert result == "Alice Smith"

    def test_exchange_cn_format(self):
        """Extracts name from Exchange CN format."""
        result = extract_display_name("/O=COMPANY/OU=USERS/CN=Alice Smith")
        assert result == "Alice Smith"

    def test_last_first_format(self):
        """Converts 'Last, First' to 'First Last'."""
        result = extract_display_name("Smith, Alice")
        assert result == "Alice Smith"

    def test_empty_string(self):
        """Returns None for empty string."""
        assert extract_display_name("") is None

    def test_none_input(self):
        """Returns None for None input."""
        assert extract_display_name(None) is None

    def test_name_with_parentheses(self):
        """Removes parenthetical content."""
        result = extract_display_name("Alice Smith (Marketing)")
        assert result == "Alice Smith"

    def test_whitespace_handling(self):
        """Strips leading/trailing whitespace."""
        result = extract_display_name("  Alice Smith  ")
        assert result == "Alice Smith"


class TestFormatOutlookDate:
    """Tests for formatting dates for Outlook Restrict filter."""

    def test_format_date(self):
        """Formats datetime correctly."""
        dt = datetime(2024, 3, 15, 14, 30)
        result = format_outlook_date(dt)
        assert result == "03/15/2024 14:30 PM"

    def test_format_midnight(self):
        """Formats midnight correctly."""
        dt = datetime(2024, 1, 1, 0, 0)
        result = format_outlook_date(dt)
        assert result == "01/01/2024 00:00 AM"

    def test_format_noon(self):
        """Formats noon correctly."""
        dt = datetime(2024, 6, 15, 12, 0)
        result = format_outlook_date(dt)
        assert result == "06/15/2024 12:00 PM"


class TestSanitizeFilename:
    """Tests for sanitizing filenames."""

    def test_simple_filename(self):
        """Keeps simple filename unchanged."""
        assert sanitize_filename("document") == "document"

    def test_removes_invalid_chars(self):
        """Removes invalid filesystem characters."""
        result = sanitize_filename('file<>:"/\\|?*name')
        assert result == "filename"

    def test_collapses_whitespace(self):
        """Collapses multiple spaces to single space."""
        result = sanitize_filename("file   name   here")
        assert result == "file name here"

    def test_truncates_long_filename(self):
        """Truncates filename to max_length."""
        long_name = "a" * 100
        result = sanitize_filename(long_name, max_length=50)
        assert len(result) == 50

    def test_strips_dots_and_spaces(self):
        """Strips leading/trailing dots and spaces."""
        result = sanitize_filename("...filename...")
        assert result == "filename"

    def test_empty_becomes_untitled(self):
        """Empty filename becomes 'untitled'."""
        assert sanitize_filename("") == "untitled"

    def test_only_invalid_chars_becomes_untitled(self):
        """Filename with only invalid chars becomes 'untitled'."""
        assert sanitize_filename("<>:*?") == "untitled"

    def test_custom_max_length(self):
        """Respects custom max_length."""
        result = sanitize_filename("this is a long filename", max_length=10)
        assert len(result) <= 10


class TestNormalizeSubject:
    """Tests for normalizing email subjects for thread grouping."""

    def test_simple_subject(self):
        """Keeps simple subject unchanged."""
        assert normalize_subject("Meeting tomorrow") == "Meeting tomorrow"

    def test_removes_re_prefix(self):
        """Removes RE: prefix."""
        assert normalize_subject("RE: Meeting tomorrow") == "Meeting tomorrow"

    def test_removes_fw_prefix(self):
        """Removes FW: prefix."""
        assert normalize_subject("FW: Meeting tomorrow") == "Meeting tomorrow"

    def test_removes_fwd_prefix(self):
        """Removes Fwd: prefix."""
        assert normalize_subject("Fwd: Meeting tomorrow") == "Meeting tomorrow"

    def test_removes_multiple_prefixes(self):
        """Removes multiple RE/FW prefixes."""
        assert normalize_subject("RE: FW: Meeting tomorrow") == "Meeting tomorrow"

    def test_case_insensitive(self):
        """Handles various cases of RE/FW."""
        assert normalize_subject("re: meeting") == "meeting"
        assert normalize_subject("Re: meeting") == "meeting"
        assert normalize_subject("RE: meeting") == "meeting"

    def test_empty_subject(self):
        """Empty subject returns 'No Subject'."""
        assert normalize_subject("") == "No Subject"

    def test_none_subject(self):
        """None subject returns 'No Subject'."""
        assert normalize_subject(None) == "No Subject"

    def test_only_prefix_returns_no_subject(self):
        """Subject with only RE: returns 'No Subject'."""
        assert normalize_subject("RE: ") == "No Subject"


class TestParseDate:
    """Tests for parsing date strings."""

    def test_iso_format(self):
        """Parses YYYY-MM-DD format."""
        result = parse_date("2024-03-15")
        assert result == datetime(2024, 3, 15)

    def test_iso_slash_format(self):
        """Parses YYYY/MM/DD format."""
        result = parse_date("2024/03/15")
        assert result == datetime(2024, 3, 15)

    def test_us_format(self):
        """Parses MM/DD/YYYY format."""
        result = parse_date("03/15/2024")
        assert result == datetime(2024, 3, 15)

    def test_eu_format(self):
        """Parses DD-MM-YYYY format."""
        result = parse_date("15-03-2024")
        assert result == datetime(2024, 3, 15)

    def test_invalid_format_raises(self):
        """Invalid format raises ValueError."""
        with pytest.raises(ValueError) as exc_info:
            parse_date("March 15, 2024")
        assert "Cannot parse date" in str(exc_info.value)

    def test_error_message_includes_format_hint(self):
        """Error message includes YYYY-MM-DD hint."""
        with pytest.raises(ValueError) as exc_info:
            parse_date("invalid")
        assert "YYYY-MM-DD" in str(exc_info.value)
