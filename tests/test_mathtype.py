import pytest

from mathtype_mcp.backend import WordBackend
from mathtype_mcp.validation import latex_for_layout, validate_layout_options


@pytest.mark.parametrize("source", ["x+1", "$x+1$", "$$x+1$$", r"\(x+1\)", r"\[x+1\]"])
def test_mathtype_inline_delimiters(source):
    assert latex_for_layout(source, "inline") == "$x+1$"
    assert latex_for_layout(source, "display") == r"\[x+1\]"


@pytest.mark.parametrize(
    "source", [r"\frac{x}{y", "x}", r"\input{a}", "$x$ $y$", r"x% comment", ""]
)
def test_reject_bad_input(source):
    with pytest.raises(ValueError):
        latex_for_layout(source, "inline")


def test_braces_escaped_and_gcd_allowed():
    assert latex_for_layout(r"\gcd(a,b)+\{x\}", "inline") == r"$\gcd(a,b)+\{x\}$"


def test_native_numbered_supported_without_manual_label():
    assert validate_layout_options("numbered").layout == "numbered"
    assert latex_for_layout("x+1", "numbered") == r"\[x+1\]"
    with pytest.raises(ValueError):
        validate_layout_options("numbered", number_label="(2)")


@pytest.mark.parametrize(
    "progid,expected", [("Equation.DSMT4", True), ("Equation.AxMath", False), ("Equation.3", False)]
)
def test_object_isolation(progid, expected):
    from types import SimpleNamespace

    shape = SimpleNamespace(Type=1, OLEFormat=SimpleNamespace(ProgID=progid))
    assert WordBackend._is_equation(None, shape) is expected
