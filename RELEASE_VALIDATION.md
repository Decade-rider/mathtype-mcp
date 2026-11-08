# Release validation — v0.2.0

Date: 2026-11-08 (Asia/Shanghai)

This release directory was assembled independently from the installed checkout.
It contained no pre-existing virtual environment, caches, artifacts or real client configuration.
A new Python 3.13.6 environment was created with `uv sync --locked --python 3.13`.

Results:
- Locked dependency synchronization: passed.
- Ruff lint and formatting: passed.
- Unit tests: 121 passed.
- MCP smoke: 12 tools; health call and invalid-ID rejection passed.
- Generic registration verification: successfully connected with the locally registered 0.2.0 entry; read-only document-list call passed. No client configuration was modified.
- Wheel and source distribution build: passed.
- MathType backend code is identical to the previously tested 0.2.0 checkout. No new native Word integration run was performed solely for packaging; historical native verification scope is documented in docs/TESTING.md.
- Runtime code and tests contain no dependency on the original checkout path.
- Private local configurations, environment directories, vendor software and raw logs are excluded from the ZIP.
- GitHub-hosted CI has not run because the repository has not been published.

Release-only changes: public documentation, license-status notice, CI and issue template,
portable registration verification, executable-presence check in the generic configuration exporter,
package metadata and source-distribution allowlist. The active installed MCP and Codex configuration
were not replaced by this release preparation.

License: intentionally unspecified pending maintainer confirmation. See LICENSE_STATUS.md.
