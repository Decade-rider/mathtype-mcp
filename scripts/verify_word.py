"""Native integration acceptance; operates only on documents created by this run."""

import json
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile

import pythoncom
import win32com.client

from mathtype_mcp.backend import WordBackend
from mathtype_mcp.word_workspace import configure_private_printer

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "artifacts" / ("word-acceptance-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
OUT.mkdir(parents=True)
report = {"cases": [], "directory": str(OUT)}
pythoncom.CoInitialize()
backend = WordBackend()
app = doc = None
try:
    app = win32com.client.DispatchEx("Word.Application")
    assert app.Documents.Count == 0
    configure_private_printer(app)
    doc = app.Documents.Add(Visible=False)
    doc.Content.Text = "中文正文 Before PLACEHOLDER After 正文结尾。\r"
    backend._remember_app(app)
    document_id = backend._remember_doc(doc)
    start = str(doc.Content.Text).index("PLACEHOLDER")
    doc.Bookmarks.Add("formula_here", doc.Range(start, start + 11))
    first = backend.insert_equation(
        document_id, r"\frac{a+b}{c+d}", position="bookmark", bookmark="formula_here"
    )
    eqid = first["equation_id"]
    assert first["progid"] == "Equation.DSMT4"
    original_text = str(doc.Content.Text)
    read = backend.get_equation_latex(document_id, eqid)
    assert r"\frac{a+b}{c+d}" in read["latex"]
    assert str(doc.Content.Text) == original_text
    print("insert/read passed", read, flush=True)
    updated = backend.update_equation(document_id, eqid, r"x^2+y^2=z^2")
    assert str(doc.Content.Text) == original_text
    read_updated = backend.get_equation_latex(document_id, eqid)
    assert "z" in read_updated["latex"] and "frac" not in read_updated["latex"]
    assert doc.InlineShapes.Count == 1
    report["cases"].append(
        {
            "name": "inline_insert_read_update_preserve_text",
            "insert": first,
            "read": read,
            "update": updated,
            "read_updated": read_updated,
        }
    )
    for name, latex in [
        ("matrix", r"\begin{bmatrix}a&b\\c&d\end{bmatrix}"),
        ("sum", r"\sum_{i=1}^{n}x_i"),
        ("integral", r"\int_0^1 x^2\,dx"),
        ("greek", r"\alpha+\beta=\gamma"),
        ("simple", "x"),
    ]:
        before = int(doc.InlineShapes.Count)
        try:
            item = backend.insert_equation(document_id, latex, layout="display")
            value = backend.get_equation_latex(document_id, item["equation_id"])
            assert doc.InlineShapes.Count == before + 1
            assert doc.Range(item["start"], item["end"]).Paragraphs(1).Format.Alignment == 1
            report["cases"].append({"name": name, "success": True, "result": item, "read": value})
            print(name, value["latex"], flush=True)
        except Exception as exc:  # noqa: BLE001 - report native acceptance failures
            assert doc.InlineShapes.Count == before
            report["cases"].append({"name": name, "success": False, "error": repr(exc)})
            print(name, "FAILED", repr(exc), flush=True)
    before_xml = (str(doc.Content.Text), [backend._fingerprint(s) for s in doc.InlineShapes])
    for invalid in [r"\frac{a}{b", r"\input{file}", "$x$ $y$"]:
        try:
            backend.update_equation(document_id, eqid, invalid)
            raise AssertionError("Invalid input was accepted")
        except ValueError:
            assert (
                str(doc.Content.Text),
                [backend._fingerprint(s) for s in doc.InlineShapes],
            ) == before_xml
    report["invalid_input_unchanged"] = True
    expected_count = doc.InlineShapes.Count
    saved = OUT / "mathtype-word-demo.docx"
    backend.save_document(document_id, str(saved))
    # Close only this run's document, then reopen the saved result read-only.
    doc.Close(0)
    doc = app.Documents.Open(str(saved), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    reopened_id = backend._remember_doc(doc)
    equations = backend.list_equations(reopened_id)["equations"]
    assert len(equations) == expected_count
    for equation in equations:
        value = backend.get_equation_latex(reopened_id, equation["equation_id"])
        assert value["latex"]
    report["reopened_equation_count"] = len(equations)
    with ZipFile(saved) as archive:
        binaries = [n for n in archive.namelist() if n.startswith("word/embeddings/")]
        assert len(binaries) == expected_count
        assert all(archive.read(n).startswith(bytes.fromhex("d0cf11e0a1b11ae1")) for n in binaries)
        assert b"Equation.DSMT4" in archive.read("word/document.xml")
    report["native_ole_storage_verified"] = True
    assert all(case.get("success", True) for case in report["cases"])
    report["completed"] = True
    print("saved/reopened", expected_count, flush=True)
finally:
    backend.shutdown()
    report["workspace_cleanup"] = backend.conversion_status()
    if doc is not None:
        doc.Close(0)
    if app is not None and app.Documents.Count == 0:
        app.Quit(0)
    pythoncom.CoUninitialize()
    (OUT / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("REPORT", OUT / "report.json", flush=True)
