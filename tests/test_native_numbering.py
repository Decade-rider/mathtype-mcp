"""Native numbering boundaries and temporary document setting restoration."""

from types import SimpleNamespace

import pytest

from mathtype_mcp.native_numbering import preflight_native_numbering, temporary_property


class Properties(list):
    def __call__(self, name):
        return next(p for p in self if p.Name == name)

    def Add(self, name, _linked, _kind, value):
        prop = SimpleNamespace(Name=name, Value=value)
        prop.Delete = lambda: self.remove(prop)
        self.append(prop)


def document():
    return SimpleNamespace(
        CustomDocumentProperties=Properties(),
        Application=SimpleNamespace(
            Templates=[SimpleNamespace(Name="MathType Commands 2016.dotm")]
        ),
    )


@pytest.mark.parametrize("existing", [None, "0", "1"])
def test_temporary_vendor_settings_restored_after_macro_failure(existing):
    doc = document()
    if existing is not None:
        doc.CustomDocumentProperties.Add("MTDeferFieldUpdate", False, 4, existing)
    with pytest.raises(RuntimeError), temporary_property(doc, "MTDeferFieldUpdate", "1"):
        assert doc.CustomDocumentProperties("MTDeferFieldUpdate").Value == "1"
        raise RuntimeError("macro failed")
    assert [p.Value for p in doc.CustomDocumentProperties] == (
        [] if existing is None else [existing]
    )


def test_custom_chapter_settings_are_never_overwritten():
    doc = document()
    doc.CustomDocumentProperties.Add("MTEquationNumber2", False, 4, "(#S1.#E1)")
    with pytest.raises(ValueError, match="custom/chapter"):
        preflight_native_numbering(doc)
    assert doc.CustomDocumentProperties("MTEquationNumber2").Value == "(#S1.#E1)"


def test_target_requires_its_own_loaded_axmath_template():
    doc = document()
    doc.Application.Templates = []
    with pytest.raises(RuntimeError, match="target Word"):
        preflight_native_numbering(doc)
