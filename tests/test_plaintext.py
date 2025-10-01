"""Regression tests for the LaTeX to plain-text helpers."""

from __future__ import annotations

import sys
import types
from importlib import util
from pathlib import Path

import pytest


def _install_test_stubs(monkeypatch: pytest.MonkeyPatch) -> None:
    """Provide lightweight shims for optional runtime dependencies."""

    # ttkthemes is only required for the GUI, so a simple placeholder works.
    themed_module = types.ModuleType("ttkthemes")
    themed_module.ThemedTk = type("DummyThemedTk", (), {})
    monkeypatch.setitem(sys.modules, "ttkthemes", themed_module)

    # Pillow stubs that satisfy the module-level imports without shipping the
    # heavyweight dependency in the test environment.
    pil_module = types.ModuleType("PIL")
    pil_image = types.ModuleType("PIL.Image")
    pil_image.MAX_IMAGE_PIXELS = None
    pil_image.LANCZOS = 1

    def _unavailable(*_args, **_kwargs):  # pragma: no cover - defensive helper
        raise RuntimeError("Image operations are not available in tests.")

    pil_image.open = _unavailable
    pil_imagetk = types.ModuleType("PIL.ImageTk")
    pil_module.Image = pil_image
    pil_module.ImageTk = pil_imagetk
    monkeypatch.setitem(sys.modules, "PIL", pil_module)
    monkeypatch.setitem(sys.modules, "PIL.Image", pil_image)
    monkeypatch.setitem(sys.modules, "PIL.ImageTk", pil_imagetk)

    # Matplotlib shims that keep rcParams available and provide inert pyplot
    # helpers so module import succeeds.
    mpl_module = types.ModuleType("matplotlib")
    mpl_module.rcParams = {}
    mpl_module.use = lambda _backend: None

    pyplot = types.ModuleType("matplotlib.pyplot")

    class _DummyFigure:
        def __init__(self) -> None:
            self.patch = types.SimpleNamespace(set_alpha=lambda *_a, **_k: None)

        def add_axes(self, _rect):
            return types.SimpleNamespace(
                axis=lambda *_a, **_k: None,
                text=lambda *_a, **_k: None,
            )

    pyplot.figure = lambda *_a, **_k: _DummyFigure()
    pyplot.savefig = lambda *_a, **_k: None
    pyplot.close = lambda *_a, **_k: None
    mpl_module.pyplot = pyplot
    monkeypatch.setitem(sys.modules, "matplotlib", mpl_module)
    monkeypatch.setitem(sys.modules, "matplotlib.pyplot", pyplot)


@pytest.fixture
def latexclip(monkeypatch: pytest.MonkeyPatch):
    """Import the module under test with lightweight dependency shims."""

    # Ensure a clean import each time the fixture initialises.
    sys.modules.pop("latexclip", None)
    _install_test_stubs(monkeypatch)

    spec = util.spec_from_file_location("latexclip", Path(__file__).resolve().parent.parent / "latexclip.py")
    if spec is None or spec.loader is None:  # pragma: no cover - defensive guard
        raise RuntimeError("Unable to load latexclip module for testing.")

    module = util.module_from_spec(spec)
    sys.modules["latexclip"] = module
    spec.loader.exec_module(module)
    return module


def test_plaintext_preserves_nested_fractions(latexclip):
    result = latexclip.latex_to_plaintext(r"\frac{1+\frac{1}{x}}{y}")
    assert result == r"\frac(1+(1)/(x))(y)"


def test_plaintext_handles_nested_sqrt(latexclip):
    result = latexclip.latex_to_plaintext(r"\sqrt{\frac{a}{b}}")
    assert result == r"sqrt(\frac(a)(b))"


def test_plaintext_expands_text_blocks(latexclip):
    result = latexclip.latex_to_plaintext(r"\text{Area} = \frac{1}{2} b h")
    assert result == "Area = (1)/(2) b h"


def test_sanitizer_escapes_plain_text_specials(latexclip):
    result = latexclip.sanitize_for_mathtext("Save 50% & more #1")
    assert result == r"$Save 50\% \& more \#1$"


def test_sanitizer_retains_existing_escapes(latexclip):
    result = latexclip.sanitize_for_mathtext(r"Already escaped \% value")
    assert result == r"$Already escaped \% value$"
