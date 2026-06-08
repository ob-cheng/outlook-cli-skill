"""People directory manager for outlook-cli.

Stores a growing list of known name+email pairs in ~/.outlook-cli/people.json.
Auto-discovered contacts are added silently; the user is informed but not asked.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

PEOPLE_FILE = Path.home() / ".outlook-cli" / "people.json"


class PeopleManager:
    """Read/write the people directory."""

    def __init__(self, path: Path | None = None):
        self.path = path or PEOPLE_FILE

    def _ensure_dir(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def _load(self) -> list[dict[str, str]]:
        if not self.path.exists():
            return []
        try:
            with open(self.path) as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []

    def _save(self, people: list[dict[str, str]]) -> None:
        self._ensure_dir()
        with open(self.path, "w") as f:
            json.dump(people, f, indent=2)
            f.write("\n")

    def list(self) -> list[dict[str, str]]:
        """Return all known people."""
        return self._load()

    def lookup(self, query: str) -> list[dict[str, str]]:
        """Find people matching name or email (case-insensitive substring)."""
        q = query.lower()
        people = self._load()
        return [
            p for p in people
            if q in p.get("name", "").lower()
            or q in p.get("email", "").lower()
        ]

    def add(self, name: str, email: str) -> bool:
        """Add a person. Returns False if duplicate (name+email pair already exists)."""
        people = self._load()
        for p in people:
            if p["name"] == name and p["email"] == email:
                return False
        people.append({"name": name, "email": email})
        people.sort(key=lambda p: p["name"].lower())
        self._save(people)
        return True

    def extract_and_add(self, people_map: dict[str, str]) -> list[str]:
        """Add multiple people at once from a {name: email} map.

        Returns a list of formatted strings for people that were actually added.
        Skips any name+email pairs already in the directory.
        """
        added = []
        known = self._load()
        known_set = {(p["name"], p["email"]) for p in known}

        for name, email in sorted(people_map.items()):
            if not name or not email:
                continue
            if (name, email) not in known_set:
                known.append({"name": name, "email": email})
                known_set.add((name, email))
                added.append(f"{name} <{email}>")

        if added:
            known.sort(key=lambda p: p["name"].lower())
            self._save(known)

        return added
