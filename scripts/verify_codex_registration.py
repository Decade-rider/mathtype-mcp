"""Read-only connection test using the registered Codex MathType entry."""

import argparse
import asyncio
import json
import os
import tomllib
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

from mathtype_mcp import __version__

ROOT = Path(__file__).resolve().parents[1]


async def verify(config: Path, backup: Path | None = None):
    current = tomllib.loads(config.read_text(encoding="utf-8-sig"))
    entry = current.get("mcp_servers", {}).get("mathtype")
    if not entry:
        raise ValueError("MathType is not registered. Run register_codex.py --install first.")
    if entry.get("enabled") is False:
        raise ValueError("The registered MathType server is disabled.")
    preserved = None
    if backup:
        before = tomllib.loads(backup.read_text(encoding="utf-8-sig"))
        other = tomllib.loads(config.read_text(encoding="utf-8-sig"))
        for data in (before, other):
            servers = data.get("mcp_servers", {})
            servers.pop("mathtype", None)
            if not servers:
                data.pop("mcp_servers", None)
        if before != other:
            raise ValueError("Other settings differ from the supplied backup; no changes made.")
        preserved = True
    parameters = StdioServerParameters(
        command=entry["command"],
        args=entry.get("args", []),
        env=entry.get("env"),
        cwd=entry.get("cwd"),
    )
    async with stdio_client(parameters) as (read, write), ClientSession(read, write) as client:
        await client.initialize()
        tools = await client.list_tools()
        names = {t.name for t in tools.tools}
        expected = {"mathtype_health", "word_number_mathtype", "word_refresh_equation_numbers"}
        if len(names) != 12 or not expected <= names:
            raise RuntimeError("Unexpected tool set. Check registered checkout and restart server.")
        health = await client.call_tool("mathtype_health", {})
        if health.isError:
            raise RuntimeError("Health call failed: " + str(health.content))
        data = health.structuredContent or json.loads(health.content[0].text)
        if data.get("version") != __version__:
            raise RuntimeError("Registered server version differs from this checkout.")
        documents = await client.call_tool("word_list_documents", {})
        if documents.isError:
            raise RuntimeError("Word document-list call failed: " + str(documents.content))
        report = {
            "version": data["version"],
            "tool_count": len(names),
            "document_list_call_ok": True,
            "other_settings_preserved": preserved,
            "note": "Connection test only; no Word conversion was exercised.",
        }
        (ROOT / "artifacts").mkdir(exist_ok=True)
        (ROOT / "artifacts/codex-registration-check.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(json.dumps(report, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(os.environ.get("CODEX_HOME", str(Path.home() / ".codex"))) / "config.toml",
    )
    parser.add_argument(
        "--backup", type=Path, help="Optional backup to compare unrelated settings."
    )
    args = parser.parse_args()
    asyncio.run(verify(args.config, args.backup))


if __name__ == "__main__":
    main()
