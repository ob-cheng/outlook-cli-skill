"""Tests for argument parsing — all commands and flags.

These test argparse directly without importing the CLI module
(which would trigger win32com import and fail on non-Windows).
"""

import argparse
import pytest


# ---------------------------------------------------------------------------
# Reconstruct the parser setup from cli.py (argparse portion only)
# ---------------------------------------------------------------------------

def build_parser():
    """Build the same argparse parser as cli.py main()."""
    parser = argparse.ArgumentParser("outlook")

    subparsers = parser.add_subparsers(dest="command")

    # --- search ---
    sp = subparsers.add_parser("search")
    sp.add_argument("--folder", "-F", action="append")
    sp.add_argument("--days", "-d", type=int, default=7)
    sp.add_argument("--from-date", type=str)
    sp.add_argument("--to-date", type=str)
    sp.add_argument("--unread", "-u", action="store_true")
    sp.add_argument("--filter-email", "-f", action="append")
    sp.add_argument("--filter-domain", "-D", action="append")
    sp.add_argument("--keyword", "-k", type=str)
    sp.add_argument("--limit", "-N", type=int, default=None)
    sp.add_argument("--export", "-e", type=str)
    sp.add_argument("--no-view", action="store_true")
    sp.add_argument("--json", action="store_true")

    # --- reply ---
    rp = subparsers.add_parser("reply")
    rp.add_argument("message_id")
    rp.add_argument("--body", "-b", type=str, required=True)
    rp.add_argument("--all", action="store_true")
    rp.add_argument("--attach", "-a", action="append")
    rp.add_argument("--html", action="store_true")
    rp.add_argument("--cc", type=str, default="")
    rp.add_argument("--bcc", type=str, default="")
    rp.add_argument("--send", action="store_true")
    rp.add_argument("--json", action="store_true")

    # --- forward ---
    fp = subparsers.add_parser("forward")
    fp.add_argument("message_id")
    fp.add_argument("--to", type=str, required=True)
    fp.add_argument("--body", "-b", type=str)
    fp.add_argument("--cc", type=str, default="")
    fp.add_argument("--bcc", type=str, default="")
    fp.add_argument("--attach", "-a", action="append")
    fp.add_argument("--html", action="store_true")
    fp.add_argument("--send", action="store_true")
    fp.add_argument("--json", action="store_true")

    # --- send ---
    sp2 = subparsers.add_parser("send")
    sp2.add_argument("--to", type=str, required=True)
    sp2.add_argument("--subject", type=str, required=True)
    sp2.add_argument("--body", type=str, required=True)
    sp2.add_argument("--cc", type=str, default="")
    sp2.add_argument("--bcc", type=str, default="")
    sp2.add_argument("--attach", "-a", action="append")
    sp2.add_argument("--html", action="store_true")
    sp2.add_argument("--send", action="store_true")
    sp2.add_argument("--json", action="store_true")

    # --- export ---
    ep = subparsers.add_parser("export")
    ep.add_argument("--output", "-o", type=str, required=True)
    ep.add_argument("--days", "-d", type=int, default=7)
    ep.add_argument("--folder", "-F", action="append")
    ep.add_argument("--filter-email", "-f", action="append")
    ep.add_argument("--filter-domain", "-D", action="append")
    ep.add_argument("--keyword", "-k", type=str)
    ep.add_argument("--limit", "-N", type=int, default=None)
    ep.add_argument("--format", type=str, choices=["markdown", "json"], default="markdown")
    ep.add_argument("--batch", action="store_true")
    ep.add_argument("--stdout", action="store_true")
    ep.add_argument("--no-threads", action="store_true")
    ep.add_argument("--no-overwrite", action="store_true")
    ep.add_argument("--incremental", action="store_true")
    ep.add_argument("--json", action="store_true")

    # --- folders ---
    fop = subparsers.add_parser("folders")
    fop.add_argument("--json", action="store_true")

    return parser


# ---------------------------------------------------------------------------
# CC/BCC parsing helper (identical to the one in cli.py cmd_reply)
# ---------------------------------------------------------------------------

def parse_cc_bcc(args):
    """Replicate the CC/BCC parsing logic from cmd_reply."""
    cc = [e.strip() for e in args.cc.split(",") if e.strip()] if args.cc else None
    bcc = [e.strip() for e in args.bcc.split(",") if e.strip()] if args.bcc else None
    return cc, bcc


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestReplyArgs:
    def test_cc_parses_comma_separated(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test", "--cc", "alice@x.com,bob@x.com"])
        cc, bcc = parse_cc_bcc(args)
        assert cc == ["alice@x.com", "bob@x.com"]
        assert bcc is None

    def test_cc_defaults_to_none_when_omitted(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test"])
        cc, bcc = parse_cc_bcc(args)
        assert cc is None
        assert bcc is None

    def test_bcc_parses_correctly(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test", "--bcc", "hidden@x.com"])
        cc, bcc = parse_cc_bcc(args)
        assert cc is None
        assert bcc == ["hidden@x.com"]

    def test_cc_and_bcc_together(self):
        parser = build_parser()
        args = parser.parse_args(
            ["reply", "abc123", "--body", "test", "--cc", "a@x.com", "--bcc", "b@x.com,c@x.com"]
        )
        cc, bcc = parse_cc_bcc(args)
        assert cc == ["a@x.com"]
        assert bcc == ["b@x.com", "c@x.com"]

    def test_cc_empty_string_after_strip_is_filtered(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test", "--cc", "alice@x.com, ,bob@x.com"])
        cc, _ = parse_cc_bcc(args)
        assert cc == ["alice@x.com", "bob@x.com"]

    def test_cc_single_value_no_comma(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test", "--cc", "onlyone@x.com"])
        cc, _ = parse_cc_bcc(args)
        assert cc == ["onlyone@x.com"]

    def test_reply_all_flag(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test", "--all"])
        assert args.all is True

    def test_reply_send_flag(self):
        parser = build_parser()
        args = parser.parse_args(["reply", "abc123", "--body", "test", "--send", "--json"])
        assert args.send is True
        assert args.json is True


class TestSearchArgs:
    def test_limit_default_none(self):
        parser = build_parser()
        args = parser.parse_args(["search"])
        assert args.limit is None

    def test_limit_parsed_as_int(self):
        parser = build_parser()
        args = parser.parse_args(["search", "--limit", "10"])
        assert args.limit == 10
        assert isinstance(args.limit, int)

    def test_limit_short_flag(self):
        parser = build_parser()
        args = parser.parse_args(["search", "-N", "5"])
        assert args.limit == 5

    def test_limit_0_edge_case(self):
        parser = build_parser()
        args = parser.parse_args(["search", "--limit", "0"])
        assert args.limit == 0

    def test_days_default(self):
        parser = build_parser()
        args = parser.parse_args(["search"])
        assert args.days == 7

    def test_multiple_folders(self):
        parser = build_parser()
        args = parser.parse_args(["search", "--folder", "Inbox", "-F", "Archive"])
        assert args.folder == ["Inbox", "Archive"]

    def test_filter_email_multiple(self):
        parser = build_parser()
        args = parser.parse_args(["search", "-f", "a@x.com", "--filter-email", "b@x.com"])
        assert args.filter_email == ["a@x.com", "b@x.com"]

    def test_keyword(self):
        parser = build_parser()
        args = parser.parse_args(["search", "--keyword", "invoice"])
        assert args.keyword == "invoice"

    def test_json_flag(self):
        parser = build_parser()
        args = parser.parse_args(["search", "--json"])
        assert args.json is True


class TestExportArgs:
    def test_limit_on_export(self):
        parser = build_parser()
        args = parser.parse_args(["export", "--output", "/tmp/out", "--limit", "25"])
        assert args.limit == 25

    def test_export_limit_default_none(self):
        parser = build_parser()
        args = parser.parse_args(["export", "--output", "/tmp/out"])
        assert args.limit is None


class TestForwardArgs:
    def test_forward_cc_bcc(self):
        parser = build_parser()
        args = parser.parse_args(
            ["forward", "abc123", "--to", "bob@x.com", "--cc", "cc@x.com", "--bcc", "bcc@x.com"]
        )
        assert args.to == "bob@x.com"
        assert args.cc == "cc@x.com"
        assert args.bcc == "bcc@x.com"
