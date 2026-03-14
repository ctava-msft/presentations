"""Tests for src/enrichment.py."""

from __future__ import annotations

import os
from unittest.mock import patch, MagicMock

import pytest

from src.enrichment import (
    _HTMLTextExtractor,
    _extract_ai_bullets,
    _fetch_url_text,
    _get_openai_endpoint,
    enrich_content_from_urls,
    enrich_notes_from_urls,
)


# ---------------------------------------------------------------------------
# _get_openai_endpoint
# ---------------------------------------------------------------------------


class TestGetOpenaiEndpoint:
    def test_from_ai_project_name(self):
        with patch.dict(os.environ, {"AI_PROJECT_NAME": "myacct"}, clear=False):
            assert _get_openai_endpoint() == "https://myacct.openai.azure.com"

    def test_from_project_endpoint(self):
        env = {
            "AI_PROJECT_NAME": "",
            "AZURE_AI_PROJECT_ENDPOINT": "https://myproj.services.ai.azure.com/foo",
        }
        with patch.dict(os.environ, env, clear=False):
            assert _get_openai_endpoint() == "https://myproj.openai.azure.com"

    def test_returns_none_when_unset(self):
        env = {"AI_PROJECT_NAME": "", "AZURE_AI_PROJECT_ENDPOINT": ""}
        with patch.dict(os.environ, env, clear=False):
            assert _get_openai_endpoint() is None


# ---------------------------------------------------------------------------
# _HTMLTextExtractor
# ---------------------------------------------------------------------------


class TestHTMLTextExtractor:
    def test_extracts_plain_text(self):
        ext = _HTMLTextExtractor()
        ext.feed("<p>Hello <b>World</b></p>")
        assert "Hello" in ext.get_text()
        assert "World" in ext.get_text()

    def test_strips_script_tags(self):
        ext = _HTMLTextExtractor()
        ext.feed("<p>Visible</p><script>hidden();</script><p>Also visible</p>")
        text = ext.get_text()
        assert "Visible" in text
        assert "Also visible" in text
        assert "hidden" not in text

    def test_strips_style_tags(self):
        ext = _HTMLTextExtractor()
        ext.feed("<style>.x{color:red}</style><p>Content</p>")
        text = ext.get_text()
        assert "Content" in text
        assert "color" not in text

    def test_empty_html(self):
        ext = _HTMLTextExtractor()
        ext.feed("")
        assert ext.get_text().strip() == ""


# ---------------------------------------------------------------------------
# _extract_ai_bullets
# ---------------------------------------------------------------------------


class TestExtractAiBullets:
    def test_dash_bullets(self):
        text = "- Bullet one\n- Bullet two\n- Bullet three"
        result = _extract_ai_bullets(text, max_bullets=2)
        assert result == ["Bullet one", "Bullet two"]

    def test_asterisk_bullets(self):
        text = "* Alpha\n* Beta"
        result = _extract_ai_bullets(text, max_bullets=3)
        assert result == ["Alpha", "Beta"]

    def test_ignores_non_bullet_lines(self):
        text = "Some intro text\n- Actual bullet\nMore text"
        result = _extract_ai_bullets(text)
        assert result == ["Actual bullet"]

    def test_empty_input(self):
        assert _extract_ai_bullets("") == []

    def test_max_bullets_respected(self):
        text = "- A\n- B\n- C\n- D"
        assert len(_extract_ai_bullets(text, max_bullets=2)) == 2


# ---------------------------------------------------------------------------
# enrich_notes_from_urls
# ---------------------------------------------------------------------------


class TestEnrichNotesFromUrls:
    def test_no_urls_is_noop(self):
        slide = {"content_urls": [], "notes": "Original"}
        enrich_notes_from_urls(slide)
        assert slide["notes"] == "Original"

    def test_no_endpoint_skips(self):
        slide = {"content_urls": ["https://example.com"], "notes": "N", "title": "T"}
        env = {"AI_PROJECT_NAME": "", "AZURE_AI_PROJECT_ENDPOINT": ""}
        with patch.dict(os.environ, env, clear=False):
            enrich_notes_from_urls(slide)
        assert slide["notes"] == "N"

    def test_enrichment_appends_notes(self):
        slide = {
            "content_urls": ["https://example.com"],
            "notes": "Original notes",
            "title": "Test Slide",
        }
        mock_response = MagicMock()
        mock_response.choices = [MagicMock()]
        mock_response.choices[0].message.content = "- Supplemental point 1"

        mock_cred_cls = MagicMock()
        mock_cred_cls.return_value.get_token.return_value.token = "fake_token"
        mock_client_cls = MagicMock()
        mock_client_cls.return_value.chat.completions.create.return_value = mock_response

        with patch.dict(os.environ, {"AI_PROJECT_NAME": "acct"}, clear=False), \
             patch("src.enrichment._fetch_url_text", return_value="Fetched content"), \
             patch.dict("sys.modules", {
                 "azure.identity": MagicMock(DefaultAzureCredential=mock_cred_cls),
                 "openai": MagicMock(AzureOpenAI=mock_client_cls),
             }):
            enrich_notes_from_urls(slide, text_model="gpt-4o-mini")

        assert "Original notes" in slide["notes"]
        assert "Supplemental" in slide["notes"]


# ---------------------------------------------------------------------------
# enrich_content_from_urls
# ---------------------------------------------------------------------------


class TestEnrichContentFromUrls:
    def test_non_content_type_skipped(self):
        slide = {"type": "title", "content_urls": ["https://x.com"]}
        enrich_content_from_urls(slide)
        # Should not crash or modify

    def test_no_urls_is_noop(self):
        slide = {"type": "content", "content_urls": [], "bullets": ["A"]}
        enrich_content_from_urls(slide)
        assert slide["bullets"] == ["A"]

    def test_no_endpoint_skips(self):
        slide = {
            "type": "content",
            "content_urls": ["https://example.com"],
            "bullets": ["A"],
            "title": "T",
        }
        env = {"AI_PROJECT_NAME": "", "AZURE_AI_PROJECT_ENDPOINT": ""}
        with patch.dict(os.environ, env, clear=False):
            enrich_content_from_urls(slide)
        assert slide["bullets"] == ["A"]
