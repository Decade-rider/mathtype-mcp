"""Resolve the installed desktop MathType Word template without changing security."""

import os
import struct
from pathlib import Path


def find_template(bits: int = 64) -> Path | None:
    override = os.environ.get("MATHTYPE_MCP_TEMPLATE")
    if override:
        path = Path(override)
        return path if path.is_file() else None
    for base in [os.environ.get("ProgramFiles(x86)", ""), os.environ.get("ProgramFiles", "")]:
        path = (
            Path(base) / "MathType" / "Office Support" / str(bits) / "MathType Commands 2016.dotm"
        )
        if path.is_file():
            return path
    return None


def ensure_template(app):
    if any(str(t.Name).lower().startswith("mathtype commands") for t in app.Templates):
        return
    executable = Path(str(app.Path)) / "WINWORD.EXE"
    with executable.open("rb") as stream:
        stream.seek(0x3C)
        offset = struct.unpack("<I", stream.read(4))[0]
        stream.seek(offset + 4)
        machine = struct.unpack("<H", stream.read(2))[0]
    if machine not in (0x8664, 0x14C):
        raise RuntimeError("Unsupported Word architecture.")
    template = find_template(64 if machine == 0x8664 else 32)
    if template is None:
        raise RuntimeError(
            "MathType Word template was not found. Repair MathType Office integration."
        )
    app.AddIns.Add(str(template), True)
