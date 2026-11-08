# 上传 GitHub 前

本目录是仓库根目录，可直接解压后作为新仓库内容。不要上传外层 ZIP 本身来替代源码仓库。

- 已排除虚拟环境、缓存、真实客户端配置、配置备份、原始验收日志与用户文档。
- 许可证按维护者要求保留待确认状态，见 LICENSE_STATUS.md；尚未指定上游项目 URL，见 NOTICE.md。
- 通用配置含明确的路径占位符，需要用户按实际目录生成或填写。
- 尚未发布到 PyPI，也未创建 GitHub 仓库或 Release。
- 可使用 docs/RELEASE_NOTES_0.2.0.md 作为首次 Release 描述。
- CI 已配置为 Windows 单元检查，不在 CI 上尝试启动 Office。

将来选择正式许可证、补充真实仓库 URL 后，再更新 pyproject.toml 元数据和来源说明；不要在未确认前填写虚构作者、仓库地址或许可信息。

构建产物仅保留发行所需文件，不提交 dist/、artifacts/ 或本机 Codex 配置备份。首次提交前检查暂存文件列表。后续发布先运行 docs/TESTING.md 的适用检查。
