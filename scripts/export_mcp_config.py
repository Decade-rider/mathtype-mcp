import json
import shutil
from pathlib import Path

root = Path(__file__).resolve().parents[1]
uv = shutil.which("uv")
if not uv:
    raise SystemExit("uv was not found on PATH. Install uv before configuring the client.")
config = {
    "mcpServers": {
        "mathtype": {
            "command": uv,
            "args": ["--directory", str(root), "run", "--no-sync", "mathtype-mcp"],
        }
    }
}
print(json.dumps(config, ensure_ascii=False, indent=2))
