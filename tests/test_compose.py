"""Tests for compose service — CC/BCC parsing from args."""

import pytest
from unittest.mock import MagicMock

from outlook_cli.services.compose import ComposeService, _as_html_fragment, _prepend_html


class TestCCBCCParsing:
    """Test the CC/BCC parsing logic used in cmd_reply and cmd_send."""

    def parse_cc_bcc(self, cc_raw, bcc_raw):
        """Replicate the exact parsing logic from cli.py cmd_reply."""
        cc = [e.strip() for e in cc_raw.split(",") if e.strip()] if cc_raw else None
        bcc = [e.strip() for e in bcc_raw.split(",") if e.strip()] if bcc_raw else None
        return cc, bcc

    def test_both_empty_returns_none(self):
        cc, bcc = self.parse_cc_bcc("", "")
        assert cc is None
        assert bcc is None

    def test_whitespace_only_returns_empty_list(self):
        """Whitespace-only input is truthy string, yields empty list after strip."""
        cc, _ = self.parse_cc_bcc("  ,  ", "")
        # "  ,  " is truthy → split gives ["  ", "  "] → strip → ["", ""] → filtered → []
        assert cc == []

    def test_mixed_valid_and_whitespace(self):
        cc, _ = self.parse_cc_bcc("alice@x.com, , bob@x.com", "")
        assert cc == ["alice@x.com", "bob@x.com"]

    def test_no_duplicate_normalization(self):
        """CC list preserves duplicates — dedup is caller's responsibility."""
        cc, _ = self.parse_cc_bcc("a@x.com,a@x.com", "")
        assert cc == ["a@x.com", "a@x.com"]

    def test_case_preserved(self):
        cc, _ = self.parse_cc_bcc("Alice@Company.COM", "")
        assert cc == ["Alice@Company.COM"]

    def test_send_cc_parsing(self):
        """The same parsing logic applies to send command (identical parser setup)."""
        cc, bcc = self.parse_cc_bcc("cc1@x.com,cc2@x.com", "bcc1@x.com")
        assert cc == ["cc1@x.com", "cc2@x.com"]
        assert bcc == ["bcc1@x.com"]


class TestSendEmailSignature:
    """send_email must preserve Outlook's auto-inserted signature (issue #12).

    GetInspector is what makes Outlook populate Body/HTMLBody with the
    default signature; the mock pre-seeds Body/HTMLBody to stand in for
    that side effect and asserts the user's text is prepended, not it
    overwriting the signature.
    """

    def _service_with_mail(self, mail):
        namespace = MagicMock()
        namespace.Application.CreateItem.return_value = mail
        return ComposeService(namespace)

    def test_plain_text_prepends_before_existing_signature(self):
        mail = MagicMock()
        mail.Body = "Best regards,\nSomeone"
        service = self._service_with_mail(mail)

        service.send_email(to=["a@x.com"], subject="Hi", body="Hello there", html=False)

        assert mail.Body == "Hello there\n\nBest regards,\nSomeone"

    def test_html_prepends_before_existing_signature(self):
        mail = MagicMock()
        mail.HTMLBody = "<p>Best regards</p>"
        service = self._service_with_mail(mail)

        service.send_email(
            to=["a@x.com"], subject="Hi", body="<p>Hello</p>", html=True
        )

        assert mail.HTMLBody == "<p>Hello</p><br><br><p>Best regards</p>"


class TestPrependHtml:
    """new content must land inside <body>, not before <html> (font/style bug)."""

    def test_inserts_after_body_tag_not_before_html(self):
        doc = '<html><head><style>p{font-family:"Open Sans"}</style></head><body>sig</body></html>'
        result = _prepend_html(doc, "hello<br><br>")
        assert result.startswith("<html>")
        assert "<body>hello<br><br>sig</body>" in result

    def test_no_body_tag_falls_back_to_prepend(self):
        assert _prepend_html("just text", "hello") == "hellojust text"

    def test_reuses_signature_font_style(self):
        doc = (
            "<html><body><p class=MsoNormal>"
            "<span style='font-size:10.0pt;font-family:\"Open Sans\",sans-serif'>Best regards</span>"
            "</p></body></html>"
        )
        result = _prepend_html(doc, "hello<br><br>")

        assert (
            "<span style='font-size:10.0pt;font-family:\"Open Sans\",sans-serif'>"
            "hello<br><br></span>" in result
        )
        assert result.index("hello") < result.index("Best regards")

    def test_extracted_style_keeps_original_quote_char(self):
        """A double-quoted font-family inside the style must not break the attribute."""
        doc = "<body><span style='font-family:\"Open Sans\"'>sig</span></body>"
        result = _prepend_html(doc, "x")
        assert 'style=\'font-family:"Open Sans"\'' in result


class TestAsHtmlFragment:
    """plain-text bodies need <br> so multi-line content doesn't collapse."""

    def test_plain_text_newlines_become_br(self):
        assert _as_html_fragment("line one\nline two") == "line one<br>\nline two"

    def test_existing_markup_left_untouched(self):
        body = "<p>already html</p>\nmore"
        assert _as_html_fragment(body) == body
