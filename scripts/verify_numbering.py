import json
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client

from mathtype_mcp.backend import WordBackend
from mathtype_mcp.layout import is_visible_number
from mathtype_mcp.word_workspace import configure_private_printer

root = Path(__file__).resolve().parents[1]
out = root / "artifacts" / ("numbering-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
out.mkdir(parents=True)
pythoncom.CoInitialize()
b = WordBackend()
app = doc = None
report = {}
try:
    app = win32com.client.DispatchEx("Word.Application")
    assert app.Documents.Count == 0
    configure_private_printer(app)
    doc = app.Documents.Add(Visible=False)
    b._remember_app(app)
    did = b._remember_doc(doc)

    def numbers():
        return [str(f.Result.Text) for f in doc.Fields if is_visible_number(f)]

    for latex in [r"\frac{a+b}{c+d}", r"x^2+y^2=z^2"]:
        item = b.insert_equation(did, latex, layout="numbered")
        print("insert", json.dumps(item), repr(doc.Content.Text), flush=True)
        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
    report["initial"] = numbers()
    assert numbers() == ["1", "2"], numbers()
    eqs = b.list_equations(did)["equations"]
    b.update_equation(did, eqs[0]["equation_id"], r"\sum_{i=1}^n x_i")
    assert numbers() == ["1", "2"]
    doc.Bookmarks.Add("before_first", doc.Range(0, 0))
    b.insert_equation(did, "a=b", layout="numbered", position="bookmark", bookmark="before_first")
    assert numbers() == ["1", "2", "3"], numbers()
    report["prepend"] = numbers()
    doc.InlineShapes(2).Range.Paragraphs(1).Range.Delete()
    report["refresh"] = b.refresh_equation_numbers(did)
    assert numbers() == ["1", "2"], numbers()
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
    item = b.insert_equation(did, "c=d", layout="display")
    numbered = b.number_equation(did, item["equation_id"])
    assert numbers() == ["1", "2", "3"]
    assert b.number_equation(did, item["equation_id"])["changed"] is False
    report["number_existing"] = numbered
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
    unrelated = doc.Fields.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 31)
    unrelated.Result.Text = "KEEP_THIS_RESULT"
    b.refresh_equation_numbers(did)
    assert unrelated.Result.Text == "KEEP_THIS_RESULT"
    report["unrelated_field_preserved"] = True
    number_field = next(f for f in doc.Fields if is_visible_number(f))
    number_field.Locked = True
    try:
        b.refresh_equation_numbers(did)
        raise AssertionError("Locked number accepted")
    except ValueError:
        report["locked_number_rejected"] = True
    finally:
        number_field.Locked = False
    doc.CustomDocumentProperties("MTEquationNumber2").Value = "(#S1.#E1)"
    old_count = doc.InlineShapes.Count
    try:
        b.insert_equation(did, "a=b", layout="numbered")
        raise AssertionError("Custom format was overwritten")
    except ValueError:
        assert doc.InlineShapes.Count == old_count
        assert doc.CustomDocumentProperties("MTEquationNumber2").Value == "(#S1.#E1)"
        report["custom_format_preserved"] = True
    finally:
        doc.CustomDocumentProperties("MTEquationNumber2").Value = "(#E1)"

    path = out / "native-numbering.docx"
    b.save_document(did, str(path))
    doc.Close(0)
    doc = app.Documents.Open(str(path), ReadOnly=True, Visible=False)
    assert numbers() == ["1", "2", "3"]
    report["reopened"] = numbers()
    report["passed"] = True
    print("PASS", out, flush=True)
finally:
    b.shutdown()
    if doc is not None:
        doc.Close(0)
    if app is not None and app.Documents.Count == 0:
        app.Quit(0)
    pythoncom.CoUninitialize()
    (out / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
