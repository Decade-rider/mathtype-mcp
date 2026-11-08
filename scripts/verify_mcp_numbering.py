"""Exercise native numbering through the actual stdio MCP tool interface."""

import asyncio
import json
import sys
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "artifacts" / ("mcp-numbering-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
OUT.mkdir(parents=True)


async def main():
    report = {}
    path = OUT / "mcp-numbering.docx"
    async with (
        stdio_client(
            StdioServerParameters(
                command=sys.executable, args=["-m", "mathtype_mcp.server"], cwd=str(ROOT)
            )
        ) as (read, write),
        ClientSession(read, write) as client,
    ):
        await client.initialize()

        async def call(name, **kwargs):
            response = await client.call_tool(name, kwargs)
            assert not response.isError, (name, response)
            value = response.structuredContent or json.loads(response.content[0].text)
            return value

        health = await call("mathtype_health")
        assert health["version"] == "0.2.0"
        tools = await client.list_tools()
        assert len(tools.tools) == 12
        created = await call("word_create_document", visible=False)
        did = created["document_id"]
        await call("word_save_document", document_id=did, path=str(path))
        try:
            first = await call(
                "word_insert_mathtype", document_id=did, latex=r"\frac{a+b}{c+d}", layout="numbered"
            )
            second = await call(
                "word_insert_mathtype", document_id=did, latex="x+y=z", layout="display"
            )
            second = await call(
                "word_number_mathtype", document_id=did, equation_id=second["equation_id"]
            )
            assert first["number_text"] == "(1)" and second["number_text"] == "(2)"
            changed = await call(
                "word_update_mathtype",
                document_id=did,
                equation_id=first["equation_id"],
                latex=r"\sum_{i=1}^{n}x_i",
            )
            assert changed["number_text"] == "(1)"
            refreshed = await call("word_refresh_equation_numbers", document_id=did)
            assert refreshed["numbers"] == ["1", "2"]
            await call("word_save_document", document_id=did)
            report = {
                "version": health["version"],
                "tool_count": len(tools.tools),
                "numbered_insert": first,
                "number_existing": second,
                "update_keeps_number": changed,
                "refresh": refreshed,
                "passed": True,
            }
            print(
                json.dumps(
                    {"version": health["version"], "numbers": refreshed["numbers"], "passed": True}
                ),
                flush=True,
            )
        finally:
            # Only bind to the unique test document saved above; never enumerate/close user files.
            pythoncom.CoInitialize()
            try:
                document = win32com.client.GetObject(str(path))
                app = document.Application
                document.Close(0)
                if app.Documents.Count == 0:
                    app.Quit(0)
            finally:
                pythoncom.CoUninitialize()
            (OUT / "report.json").write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
            )
    print(OUT, flush=True)


if __name__ == "__main__":
    asyncio.run(main())
