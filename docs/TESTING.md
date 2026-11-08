# 测试与兼容性

## 单元检查（不启动 Office）

```powershell
uv sync --locked
uv run --no-sync ruff check src tests scripts
uv run --no-sync ruff format --check src tests scripts
uv run --no-sync pytest -q
uv build
```

Windows CI 仅运行静态检查、单元测试和构建。CI 不代表已安装 Word/MathType，也不能验证原生公式显示。项目开发组依赖通过 uv 管理，uv.lock 随源码发布。

## MCP 协议检查

`uv run --no-sync python scripts/smoke_mcp.py` 在独立服务进程中验证握手、工具发现、健康接口及非法 ID 拒绝。它不插入公式。

`uv run --no-sync python scripts/verify_codex_registration.py` 使用实际注册参数测试连接和文档列表；配置只读，不要求已安装 AxMath。

## 原生 Office 集成检查（本机桌面）

确认没有其他客户端同时操作 Office 后执行：

```powershell
uv run --no-sync python scripts/verify_word.py
uv run --no-sync python scripts/verify_numbering.py
uv run --no-sync python scripts/verify_mcp_numbering.py
```

仅操作测试中新建的文档。结果写入 artifacts/，不提交原始运行记录。编号验收覆盖前插、删除刷新、修改保号、重复补号、保存重开、锁定字段与自定义格式保护及无关字段保留。协议编号验收通过 stdio 接口实际调用工具。

可选 Word 原生页面显示检查：

```powershell
uv run --with pillow python scripts/render_word.py "测试文档的绝对路径.docx"
```

该命令在测试文件旁生成 EMF/PNG；Pillow 不是服务运行依赖。

## 已有验收证据

2026-11-08 在 Windows、Word 16.0 和本机 MathType 安装上通过 121 项单元测试、原生公式闭环、编号增删重排与真实 MCP 编号调用。开发环境为 Python 3.13.6。MathType 可执行文件报告版本 7.11.1.462；这只记录该安装的文件版本，不宣称所有 MathType 7 版本均兼容。

32 位 Word、其他 Office 构建和其他 MathType 安装尚未形成兼容性认证。README 截图来自本次原生 Word 测试文档，只包含示例公式。

发布包的干净目录验证记录见 RELEASE_VALIDATION.md。GitHub 云端 CI 必须在实际推送后才会运行，不把本地检查称为云端已通过。
