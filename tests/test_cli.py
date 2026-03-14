"""Tests for src/cli.py."""

from __future__ import annotations

import os
import tempfile
from unittest.mock import patch

import pytest

from src.cli import main


@pytest.fixture()
def spec_file(tmp_path):
    """Create a minimal valid .spec.md file and return its path."""
    content = (
        "---\n"
        "title: Test\n"
        "output: test.pptx\n"
        "---\n"
        "\n"
        "## [title] Hello World\n"
    )
    p = tmp_path / "test.spec.md"
    p.write_text(content, encoding="utf-8")
    return str(p)


# ---------------------------------------------------------------------------
# Argument parsing
# ---------------------------------------------------------------------------


def test_missing_spec_file_exits():
    with pytest.raises(SystemExit):
        main(["nonexistent_file.spec.md"])


def test_default_output_dir(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file])
        _, kwargs = mock_render.call_args
        assert kwargs.get("output_dir") or mock_render.call_args[0][1] == "output"


def test_custom_output_dir(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file, "-o", "my_output"])
        args, kwargs = mock_render.call_args
        assert "my_output" in args or kwargs.get("output_dir") == "my_output"


def test_image_model_flag(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file, "--image-model", "dall-e-3"])
        _, kwargs = mock_render.call_args
        assert kwargs["image_model"] == "dall-e-3"


def test_image_model_default_none(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file])
        _, kwargs = mock_render.call_args
        assert kwargs["image_model"] is None


def test_refetch_flag(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file, "--refetch"])
        _, kwargs = mock_render.call_args
        assert kwargs["refetch"] is True


def test_refetch_default_false(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file])
        _, kwargs = mock_render.call_args
        assert kwargs["refetch"] is False


def test_slides_flag(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file, "--slides", "1,3,5-8"])
        _, kwargs = mock_render.call_args
        assert kwargs["slide_selection"] == "1,3,5-8"


def test_slides_default_none(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file])
        _, kwargs = mock_render.call_args
        assert kwargs["slide_selection"] is None


def test_spec_path_forwarded(spec_file):
    with patch("src.cli.render") as mock_render:
        main([spec_file])
        _, kwargs = mock_render.call_args
        assert kwargs["spec_path"] == spec_file
