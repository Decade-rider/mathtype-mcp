import sys
from pathlib import Path

import pythoncom
import win32com.client
from PIL import Image

from mathtype_mcp.word_workspace import configure_private_printer

path = Path(sys.argv[1]).resolve()
pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Word.Application")
doc = None
try:
    assert app.Documents.Count == 0
    configure_private_printer(app)
    doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    doc.Repaginate()
    metafile = path.with_suffix(".emf")
    metafile.write_bytes(bytes(doc.ActiveWindow.Panes(1).Pages(1).EnhMetaFileBits))
    with Image.open(metafile) as image:
        image._size = (
            round(float(doc.PageSetup.PageWidth) * 2),
            round(float(doc.PageSetup.PageHeight) * 2),
        )
        image.load()
        image.save(path.with_suffix(".png"))
    print(path.with_suffix(".png"))
finally:
    if doc is not None:
        doc.Close(0)
    if app.Documents.Count == 0:
        app.Quit(0)
    pythoncom.CoUninitialize()
