"""Tests for PeopleManager — CRUD operations, lookup, bulk import.

These tests use temporary files to avoid touching real people directory.
"""

import json
import pytest
from pathlib import Path

from outlook_cli.core.people import PeopleManager


class TestPeopleManagerList:
    """Tests for listing all people."""

    def test_list_returns_empty_when_no_file(self, tmp_path):
        """When file doesn't exist, returns empty list."""
        people_path = tmp_path / "people.json"
        mgr = PeopleManager(path=people_path)

        result = mgr.list()

        assert result == []

    def test_list_returns_all_people(self, tmp_path):
        """Returns all people from file."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"},
            {"name": "Bob Jones", "email": "bob@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.list()

        assert len(result) == 2
        assert result[0]["name"] == "Alice Smith"
        assert result[1]["name"] == "Bob Jones"

    def test_list_handles_invalid_json(self, tmp_path):
        """Invalid JSON returns empty list."""
        people_path = tmp_path / "people.json"
        people_path.write_text("not valid json")
        mgr = PeopleManager(path=people_path)

        result = mgr.list()

        assert result == []

    def test_list_handles_empty_file(self, tmp_path):
        """Empty file returns empty list."""
        people_path = tmp_path / "people.json"
        people_path.write_text("")
        mgr = PeopleManager(path=people_path)

        result = mgr.list()

        assert result == []


class TestPeopleManagerLookup:
    """Tests for looking up people by name or email."""

    def test_lookup_by_name_exact(self, tmp_path):
        """Lookup finds person by exact name."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"},
            {"name": "Bob Jones", "email": "bob@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("Alice Smith")

        assert len(result) == 1
        assert result[0]["email"] == "alice@example.com"

    def test_lookup_by_name_partial(self, tmp_path):
        """Lookup finds person by partial name."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"},
            {"name": "Bob Jones", "email": "bob@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("Alice")

        assert len(result) == 1
        assert result[0]["name"] == "Alice Smith"

    def test_lookup_by_email(self, tmp_path):
        """Lookup finds person by email."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"},
            {"name": "Bob Jones", "email": "bob@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("alice@example.com")

        assert len(result) == 1
        assert result[0]["name"] == "Alice Smith"

    def test_lookup_by_email_domain(self, tmp_path):
        """Lookup finds multiple people by domain."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"},
            {"name": "Bob Jones", "email": "bob@example.com"},
            {"name": "Charlie Brown", "email": "charlie@other.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("example.com")

        assert len(result) == 2

    def test_lookup_case_insensitive(self, tmp_path):
        """Lookup is case-insensitive."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("ALICE")

        assert len(result) == 1
        assert result[0]["name"] == "Alice Smith"

    def test_lookup_no_match(self, tmp_path):
        """Lookup returns empty list when no match."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("Charlie")

        assert result == []

    def test_lookup_empty_query(self, tmp_path):
        """Lookup with empty string matches all."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"},
            {"name": "Bob Jones", "email": "bob@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.lookup("")

        assert len(result) == 2


class TestPeopleManagerAdd:
    """Tests for adding a single person."""

    def test_add_creates_file_and_dir(self, tmp_path):
        """Add creates parent directories and file."""
        people_path = tmp_path / "subdir" / "people.json"
        mgr = PeopleManager(path=people_path)

        result = mgr.add("Alice Smith", "alice@example.com")

        assert result is True
        assert people_path.exists()
        data = json.loads(people_path.read_text())
        assert len(data) == 1
        assert data[0]["name"] == "Alice Smith"

    def test_add_appends_to_existing(self, tmp_path):
        """Add appends to existing people list."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.add("Bob Jones", "bob@example.com")

        assert result is True
        data = json.loads(people_path.read_text())
        assert len(data) == 2

    def test_add_sorts_alphabetically(self, tmp_path):
        """Add maintains alphabetical order by name."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Charlie Brown", "email": "charlie@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        mgr.add("Alice Smith", "alice@example.com")

        data = json.loads(people_path.read_text())
        assert data[0]["name"] == "Alice Smith"
        assert data[1]["name"] == "Charlie Brown"

    def test_add_rejects_duplicate(self, tmp_path):
        """Add returns False for duplicate name+email pair."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.add("Alice Smith", "alice@example.com")

        assert result is False
        data = json.loads(people_path.read_text())
        assert len(data) == 1

    def test_add_allows_same_name_different_email(self, tmp_path):
        """Add allows same name with different email."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@work.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.add("Alice Smith", "alice@personal.com")

        assert result is True
        data = json.loads(people_path.read_text())
        assert len(data) == 2

    def test_add_allows_same_email_different_name(self, tmp_path):
        """Add allows same email with different name (alias)."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.add("Alice S.", "alice@example.com")

        assert result is True
        data = json.loads(people_path.read_text())
        assert len(data) == 2


class TestPeopleManagerExtractAndAdd:
    """Tests for bulk import from email interactions."""

    def test_extract_and_add_from_empty(self, tmp_path):
        """Extract and add to empty directory."""
        people_path = tmp_path / "people.json"
        mgr = PeopleManager(path=people_path)

        result = mgr.extract_and_add({
            "Alice Smith": "alice@example.com",
            "Bob Jones": "bob@example.com"
        })

        assert len(result) == 2
        assert "Alice Smith <alice@example.com>" in result
        assert "Bob Jones <bob@example.com>" in result

    def test_extract_and_add_skips_existing(self, tmp_path):
        """Extract and add skips already known people."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.extract_and_add({
            "Alice Smith": "alice@example.com",
            "Bob Jones": "bob@example.com"
        })

        assert len(result) == 1
        assert "Bob Jones <bob@example.com>" in result
        assert "Alice Smith" not in " ".join(result)

    def test_extract_and_add_skips_empty_values(self, tmp_path):
        """Extract and add skips entries with empty name or email."""
        people_path = tmp_path / "people.json"
        mgr = PeopleManager(path=people_path)

        result = mgr.extract_and_add({
            "Alice Smith": "alice@example.com",
            "": "empty@example.com",
            "No Email": ""
        })

        assert len(result) == 1
        assert "Alice Smith <alice@example.com>" in result

    def test_extract_and_add_maintains_sort(self, tmp_path):
        """Extract and add maintains alphabetical order."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Bob Jones", "email": "bob@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        mgr.extract_and_add({
            "Alice Smith": "alice@example.com",
            "Charlie Brown": "charlie@example.com"
        })

        data = json.loads(people_path.read_text())
        assert data[0]["name"] == "Alice Smith"
        assert data[1]["name"] == "Bob Jones"
        assert data[2]["name"] == "Charlie Brown"

    def test_extract_and_add_returns_empty_when_all_known(self, tmp_path):
        """Extract and add returns empty list when all people known."""
        people_path = tmp_path / "people.json"
        people_path.write_text(json.dumps([
            {"name": "Alice Smith", "email": "alice@example.com"}
        ]))
        mgr = PeopleManager(path=people_path)

        result = mgr.extract_and_add({
            "Alice Smith": "alice@example.com"
        })

        assert result == []

    def test_extract_and_add_empty_map(self, tmp_path):
        """Extract and add with empty map returns empty list."""
        people_path = tmp_path / "people.json"
        mgr = PeopleManager(path=people_path)

        result = mgr.extract_and_add({})

        assert result == []
