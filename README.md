# MathType MCP for Word

让支持 MCP 的助手在 Windows 桌面版 Word 中插入、读取和修改 **原生可编辑 MathType 公式**，并调用 MathType 原生连续自动编号。

**版本：0.2.0 · Alpha · Windows desktop only**

这是独立第三方项目，不是 MathType / WIRIS / Microsoft 官方产品。不包含 Word 或 MathType，也不提供其授权。代码许可证目前待维护者确认，详见 [LICENSE_STATUS.md](LICENSE_STATUS.md)。

## 支持什么

- 以 TeX 输入插入原生 `Equation.DSMT4` OLE 公式，可继续使用 MathType 编辑。
- 从当前公式副本读取 TeX；按公式 ID 修改内容。
- 行内、独立居中、公式居中且编号靠右三种布局。
- 原生 `(1)、(2)…` 连续编号，已有独立公式补编号，增删移动后刷新编号。
- Word 文档列表、新建、打开、选区检查与显式保存。

有一份deepseek harness调用该mcp的示例[DEA数据包络分析原理简介](docs/DEA数据包络分析原理简介.docx)

**暂不支持** PowerPoint、章节/自定义编号格式配置、交叉引用刷新、公式内部字号定制，以及页眉、脚注、文本框和浮动公式。更新公式不保证保留原公式的自定义内部字体、颜色和字号。编号布局要求单栏、非表格区域。

## 安装前提

1. Windows 桌面环境；不能在 WSL、Linux、容器或网页 Word 中运行此服务。
2. 桌面版 Microsoft Word 和可正常使用的桌面版 MathType 7；先确认手工插入和编辑公式正常。
3. [uv](https://docs.astral.sh/uv/getting-started/installation/)；项目要求 Python 3.12 或以上，uv 可管理 Python 和依赖。

MathType Word 模板必须可加载。项目不降低 Office 宏安全设置，也不调整软件授权。部分组织设备可能需要管理员按组织策略配置插件。

## 快速开始

将仓库克隆或 ZIP 解压到任意本地目录，在该目录打开 PowerShell：

```powershell
uv sync --locked --default-index https://pypi.org/simple
uv run --no-sync python scripts/smoke_mcp.py
```

烟雾测试通过时输出 `tool_count: 12` 和 `health_ok: true`。它验证 MCP 连接与接口，不代表 Word 转换已通过；[测试说明](docs/TESTING.md)区分了各层检查。

### 注册到 Codex

```powershell
uv run --no-sync python scripts/register_codex.py
uv run --no-sync python scripts/register_codex.py --install
```

第一条预览配置，第二条向用户 Codex 配置追加独立的 `mathtype` 条目。安装会备份已有配置并保留其他服务，包括 AxMath；相同条目重复执行不写入，同名冲突会报错。重启或重新加载客户端 MCP 后使用。

验证实际注册参数：

```powershell
uv run --no-sync python scripts/verify_codex_registration.py
```

它不会注册服务、修改配置或保存 Word 文档。

### 其他 MCP 客户端

```powershell
uv run --no-sync python scripts/export_mcp_config.py
```

把生成的 `mcpServers.mathtype` 条目加入客户端配置。不同客户端可能使用不同外层结构，按客户端配置界面填写 command、args。建议工具超时至少 300 秒。通用占位模板见 [config/mcp.json](config/mcp.json) 和 [config/codex.toml](config/codex.toml)。

直接启动入口为 `uv run --no-sync mathtype-mcp`；这是 stdio 服务，会等待 MCP 客户端消息，不是交互式命令行。

## 如何使用

可以告诉助手：

> 使用 MathType MCP 检查环境，列出当前 Word 文档。
>
> 在指定文档末尾插入贝叶斯公式，居中并使用自动编号，先不保存。
>
> 读取这条公式，再修改其内容并保留编号。
>
> 刷新公式编号，然后另存为我指定的完整路径。

| 工具                              | 用途                                   |
| --------------------------------- | -------------------------------------- |
| `mathtype_health`               | 检查注册信息、模板路径及转换工作区状态 |
| `word_list_documents`           | 列出服务可访问的 Word 文档             |
| `word_create_document`          | 新建未保存文档                         |
| `word_open_document`            | 打开绝对本地路径                       |
| `word_document_info`            | 检查选区、书签和文档状态               |
| `word_list_equations`           | 获取原生 MathType 公式 ID              |
| `word_insert_mathtype`          | 插入行内、居中或带编号公式             |
| `word_get_equation_latex`       | 从当前对象副本读取 TeX                 |
| `word_update_mathtype`          | 修改指定公式，保留已有编号             |
| `word_number_mathtype`          | 为独占段落的已有公式添加编号           |
| `word_refresh_equation_numbers` | 刷新正文 MathType 编号                 |
| `word_save_document`            | 明确保存或另存为                       |

插入示例（document_id 必须来自工具返回）：

```json
{
  "document_id": "从文档列表获取",
  "latex": "\\frac{a+b}{c+d}",
  "position": "end",
  "layout": "numbered"
}
```

`layout` 为 `inline`、`display` 或 `numbered`。`position` 为 `end`、`bookmark` 或 `selection`；书签插入在书签文字之前，选区替换要求传入检查后得到的完全一致的 `expected_selection`。公式和文档 ID 在服务重启后应重新获取。

修改不会自动保存。原生自动编号是可更新的字段，不是静态文本；手动删除或移动公式后请调用刷新工具。已有自定义/章节编号格式不会被新插入操作覆盖。刷新限定于正文 MathType 编号域，不刷新引用或其他域。

## 注意事项与已知限制

- 操作时暂勿在 Office/MathType 中编辑，也不要同时调用另一套公式 MCP；Office 和剪贴板是共享桌面资源。
- 客户端取消或超时不保证 Office 调用已停止。先检查实际文档，再决定是否重试。
- `mathtype_health` 的检查不包含临时目录写权限、Word 启动或公式转换完整验收。模板存在不等于 Word 工作环境完全可用。
- TeX 往返成功不等于数学语义得到验证，需检查显示结果。
- 自定义段落样式 `MathTypeMCP_Display` / `MathTypeMCP_Numbered` 用于排版与识别，部分排版通过直接格式设置；并非自动适配所有论文模板。
- 已观察到外部客户端出现 Word“无法创建工作文件”弹窗，根因尚未确认；不要把当前版本视为已修复该问题。

更多内容：[排错与卸载](docs/TROUBLESHOOTING.md) · [测试与兼容性](docs/TESTING.md) · [更新日志](CHANGELOG.md) · [贡献说明](CONTRIBUTING.md)。
