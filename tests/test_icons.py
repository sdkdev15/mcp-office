"""Tests for icon library, icon renderer, and auto-icon detection."""

import os
import pytest
from pathlib import Path

from src.utils.icon_library import (
    ICONS,
    get_icon_names,
    get_icon,
    find_icon_by_keyword,
)
from src.utils.icon_renderer import (
    IconRenderer,
    _cache_key,
    _get_icon_commands,
)


class TestIconLibrary:
    """Test icon registry and keyword search."""

    def test_icons_registered(self):
        """All icons should be registered."""
        names = get_icon_names()
        assert len(names) >= 50

    def test_get_icon_exists(self):
        """get_icon should return data for known icons."""
        icon = get_icon("lock")
        assert icon is not None
        assert "path" in icon
        assert "categories" in icon

    def test_get_icon_missing(self):
        """get_icon should return None for unknown icons."""
        assert get_icon("nonexistent_icon") is None

    def test_find_icon_by_keyword_exact(self):
        """find_icon_by_keyword should match exact keywords."""
        assert find_icon_by_keyword("security") in ("lock", "shield", "key")
        assert find_icon_by_keyword("database") == "database"

    def test_find_icon_by_keyword_content(self):
        """find_icon_by_keyword should match keywords inside text."""
        assert find_icon_by_keyword("Server Infrastructure Report") == "server"
        # "Revenue" matches dollar icon; "Growth" matches trend-up
        assert find_icon_by_keyword("Revenue Growth Trend") in ("trend-up", "chart-bar", "dollar")

    def test_find_icon_by_keyword_empty(self):
        """find_icon_by_keyword should return None for empty text."""
        assert find_icon_by_keyword("") is None
        assert find_icon_by_keyword(None) is None

    def test_find_icon_by_keyword_no_match(self):
        """find_icon_by_keyword should return None when no match."""
        assert find_icon_by_keyword("xyzabc123") is None

    def test_find_icon_variety(self):
        """Test auto-detection across multiple icon categories."""
        # Security
        assert find_icon_by_keyword("authentication session") == "lock"
        # Infrastructure
        # "hosting" matches both server and cloud; order-dependent
        assert find_icon_by_keyword("AWS cloud hosting") in ("server", "cloud")
        assert find_icon_by_keyword("kubernetes docker") == "docker"
        # Business
        assert find_icon_by_keyword("target goal objective") == "target"
        assert find_icon_by_keyword("calendar schedule date") == "calendar"
        # Data
        assert find_icon_by_keyword("chart bar statistics") == "chart-bar"
        assert find_icon_by_keyword("spreadsheet table grid") == "table"
        # Communication
        assert find_icon_by_keyword("email mail message") == "mail"
        assert find_icon_by_keyword("globe internet web") == "globe"
        # DevOps
        assert find_icon_by_keyword("terminal console bash") == "terminal"
        assert find_icon_by_keyword("CI/CD pipeline") == "pipeline"
        # Actions
        assert find_icon_by_keyword("refresh reload sync") == "refresh"
        assert find_icon_by_keyword("alert warning notice") == "alert"
        # Extra
        assert find_icon_by_keyword("notification bell") == "bell"
        assert find_icon_by_keyword("incident emergency") == "incident"


class TestIconRenderer:
    """Test icon rendering to PNG."""

    @pytest.fixture
    def renderer(self, tmp_path):
        return IconRenderer(cache_dir=str(tmp_path))

    def test_render_returns_path(self, renderer):
        """render() should return a file path."""
        path = renderer.render(icon_name="lock", size=48, color="#FFFFFF", bg_color="#1E40AF", bg_shape="circle")
        assert isinstance(path, str)
        assert path.endswith(".png")

    def test_render_file_exists(self, renderer):
        """The rendered file should exist on disk."""
        path = renderer.render(icon_name="shield", size=64, color="#FFFFFF", bg_color="#059669", bg_shape="circle")
        assert os.path.exists(path)

    def test_render_caching(self, renderer):
        """Same parameters should return the cached path."""
        path1 = renderer.render(icon_name="chart-bar", size=32, color="#FFF", bg_color="#1E40AF", bg_shape="circle")
        path2 = renderer.render(icon_name="chart-bar", size=32, color="#FFF", bg_color="#1E40AF", bg_shape="circle")
        assert path1 == path2

    def test_render_different_sizes(self, renderer):
        """Different sizes should produce different files."""
        path1 = renderer.render(icon_name="lock", size=32, color="#FFF", bg_color="#1E40AF", bg_shape="circle")
        path2 = renderer.render(icon_name="lock", size=64, color="#FFF", bg_color="#1E40AF", bg_shape="circle")
        assert path1 != path2
        assert os.path.exists(path1)
        assert os.path.exists(path2)

    def test_render_no_bg(self, renderer):
        """render() with bg_shape=none should work."""
        path = renderer.render(icon_name="check", size=48, color="#00FF00", bg_shape="none")
        assert os.path.exists(path)

    def test_render_unknown_icon_fallback(self, renderer):
        """Unknown icon should fall back to a circle."""
        path = renderer.render(icon_name="unknown_xyz", size=32, color="#FFF", bg_color="#1E40AF", bg_shape="circle")
        assert os.path.exists(path)

    def test_render_numbered_circle(self, renderer):
        """render_numbered_circle should produce valid PNG."""
        path = renderer.render_numbered_circle(number=1, size=40, bg_color="#1E40AF", text_color="#FFFFFF")
        assert os.path.exists(path)
        assert path.endswith(".png")

    def test_render_batch(self, renderer):
        """render_batch should render multiple icons."""
        items = [
            {"icon_name": "lock", "size": 32, "color": "#FFF", "bg_color": "#1E40AF", "bg_shape": "circle"},
            {"icon_name": "server", "size": 32, "color": "#FFF", "bg_color": "#059669", "bg_shape": "circle"},
            {"icon_name": "chart-bar", "size": 32, "color": "#FFF", "bg_color": "#DC2626", "bg_shape": "circle"},
        ]
        paths = renderer.render_batch(items)
        assert len(paths) == 3
        for p in paths:
            assert os.path.exists(p)

    def test_cache_key_uniqueness(self):
        """_cache_key should produce unique keys for different inputs."""
        k1 = _cache_key("lock", 32, "#FFF", "#1E40AF", "circle")
        k2 = _cache_key("lock", 64, "#FFF", "#1E40AF", "circle")
        k3 = _cache_key("shield", 32, "#FFF", "#1E40AF", "circle")
        assert k1 != k2
        assert k1 != k3

    def test_icon_commands(self):
        """_get_icon_commands should return commands for known icons."""
        cmds = _get_icon_commands("lock")
        assert len(cmds) > 0
        unknown = _get_icon_commands("nonexistent")
        assert unknown == []