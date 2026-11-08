import asyncio
import json
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parents[1]


async def main():
    (ROOT / "artifacts").mkdir(exist_ok=True)
    parameters = StdioServerParameters(
        command=sys.executable, args=["-m", "mathtype_mcp.server"], cwd=str(ROOT)
    )
    async with stdio_client(parameters) as (read, write), ClientSession(read, write) as client:
        await client.initialize()
        result = await client.list_tools()
        names = [t.name for t in result.tools]
        assert len(names) == 12
        assert "word_insert_mathtype" in names and "word_update_mathtype" in names
        assert not any("axmath" in n or n.startswith("ppt_") for n in names)
        health = await client.call_tool("mathtype_health", {})
        assert not health.isError
        failure = await client.call_tool(
            "word_update_mathtype",
            {"document_id": "invalid", "equation_id": "invalid", "latex": "x+1"},
        )
        assert failure.isError
        report = {
            "tools": names,
            "health": health.model_dump(mode="json"),
            "invalid_id_rejected": True,
        }
        (ROOT / "artifacts/mcp-smoke.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(
            json.dumps({"tool_count": len(names), "health_ok": True, "invalid_id_rejected": True})
        )


if __name__ == "__main__":
    asyncio.run(main())
