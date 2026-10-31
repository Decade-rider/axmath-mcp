# 开发与发布

日常安装和使用见 [README](../README.md)，测试命令与实机证据见 [测试文档](testing.md)。Word 和 PowerPoint 功能随同一个 `axmath-mcp` 项目分发，客户端只需配置一个 `axmath` 服务。

## 开发环境

```powershell
uv sync --locked --default-index https://pypi.org/simple
uv run --no-sync python -m axmath_mcp.server
```

服务使用 stdio，stdout 保留给 MCP 协议，诊断日志应写到 stderr。手动启动后等待客户端请求是正常状态；通常由 MCP 客户端启动服务。

如果默认缓存目录受限，可以使用项目内缓存：

```powershell
uv sync --locked --cache-dir .uv-cache --default-index https://pypi.org/simple
uv run --no-sync --cache-dir .uv-cache python -m axmath_mcp.server
```

环境依赖通过 `pyproject.toml` 和 `uv.lock` 管理。项目显式使用 PyPI；不要让个人的 `UV_DEFAULT_INDEX` 设置意外改写公共锁文件。升级 Office 集成逻辑后，应同时运行无 Office 的自动检查和具备授权软件的本机验证。

## Office 调用约定

所有 Office COM 对象由同一个专用 STA 线程持有，工具请求串行执行。客户端取消等待不等于 Office 宏已停止；正在执行的调用不能被误当成“没有发生”，也不能立即盲目重试。

Word 转换在临时文档中完成：调用用户已安装的 AxMath 插件，检查原生 OLE 类型和对象数量，在另一副本中导出 LaTeX，再将通过检查的对象写入目标范围。失败时不自动重试，保存由单独工具显式执行。

PowerPoint 功能复用 Word 中的 AxMath 转换，从临时范围的 WordOpenXML 提取单一原生 OLE 存储，通过 `Shapes.AddOLEObject` 嵌入幻灯片，并再次检查 `Equation.AxMath` ProgID。读取时将 PowerPoint 对象复制到临时 Word 文档，必要时将浮动 OLE 转为行内对象，再调用 AxMath 导出。它需要本机 Word 和 `AxMath.dotm`，不依赖 `AxMath.ppam` 的 LaTeX 接口。公式定位以幻灯片左上角为原点，单位为磅；编号是关联文本框中的固定标签，不使用 Word 的 SEQ 域。PowerPoint Tags 只保存字号、尺寸和关联对象 ID，不保存可能被自动大写的 LaTeX。

Word 独立公式使用 `AxMathMCP_Display`、`AxMathMCP_Numbered` 段落样式；自动编号由已安装插件的 `AMCEqnNum.InsertEqnNum` 生成，包含原生 `AMMPlaceRM` 和嵌套 `AMEqn` 域；不再生成或维护旧版 `AxMathEquation` 序列。默认使用文档级 `(#E1)` 格式，拒绝覆盖已有章节格式。调用期间临时延迟厂商的全域更新，并跳过普通连续编号不需要的章节初始化，之后恢复文档属性。更新公式内容时保留编号，刷新编号时仅更新 AxMath 的 `AMEqn`、`AMChap`、`AMSec` 序列，避免刷新外部链接等无关域。设置段落样式应早于复制公式；原生编号宏再次改变样式后，恢复公式字体、字号、基线和段落格式。

AxMath 集成使用已安装插件中的转换入口。厂商宏不是公开承诺稳定的 SDK；升级 AxMath 后必须重新验证。项目不分发 AxMath 程序、插件或其源码，不修改软件授权状态或 Office 宏安全策略。

## 构建

```powershell
uv run --no-sync ruff check src tests scripts
uv run --no-sync ruff format --check src tests scripts
uv run --no-sync pytest -q
uv build --default-index https://pypi.org/simple
uv run --no-sync python scripts/check_dist.py
```

`dist/` 中生成 wheel 和源码包。内容检查会拒绝本机调试目录、运行环境、可执行程序及厂商文件。CI 仅进行单元测试和包检查，不声称在 GitHub 托管 runner 中完成 Office 实机验证。

## 首次发布到 GitHub

下列仓库 URL 是占位符，不表示已经创建或上传远程仓库。将 `YOUR_GITHUB_USER` 替换为实际用户名或组织名。

1. 选择项目许可证，并在根目录添加完整的 `LICENSE` 文件。目前项目没有替维护者决定许可证；可按使用目标考虑 MIT 等方案。公开仓库本身不等于授予完整开源许可，参见 [GitHub 许可证说明](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/licensing-a-repository)。
2. 在 GitHub 创建空仓库 `axmath-mcp`。本地已有 README 和 `.gitignore`，远程创建时不要重复初始化这些文件。
3. 检查暂存文件，再提交和推送。

```powershell
git init -b main
git add .github .gitignore .gitattributes README.md pyproject.toml uv.lock src tests scripts config docs LICENSE
git diff --cached --stat
git diff --cached --name-only
git commit -m "Prepare axmath-mcp release"
git remote add origin https://github.com/YOUR_GITHUB_USER/axmath-mcp.git
git push -u origin main
```

若 Git 提示缺少提交身份，可先用 `git config --local user.name "你的名字"` 和 `git config --local user.email "你的 GitHub 提交邮箱"` 设置此仓库身份；推送时按 Git 客户端提示登录。流程依据 [GitHub 导入本地代码说明](https://docs.github.com/en/migrations/importing-source-code/using-the-command-line-to-import-source-code/adding-locally-hosted-code-to-github)。

`.gitignore` 排除了 `.venv/`、缓存、`.research/`、`artifacts/` 和本地配置。不要强制加入这些目录，也不要把整个开发目录压缩上传。`.research/` 可能包含提取的厂商 VBA 和本机调试材料；测试日志可能包含私人文件路径。精选示例图片放在 `docs/images/` 后再纳入仓库。

## 发布版本

发布前应确认：

- 版本号、安装示例和变更说明一致。
- 新环境中的安装与 MCP 握手通过。
- Word 与 PowerPoint 的受影响功能分别完成实机验证。
- wheel 和源码包通过内容检查。
- 许可证、仓库地址和文档链接已填写正确。

完成后创建与实际版本对应的标签。下面的 `vX.Y.Z` 需要替换为版本号，不可直接照抄：

```powershell
git tag -a vX.Y.Z -m "AxMath MCP X.Y.Z"
git push origin vX.Y.Z
```

在 GitHub Releases 页面选择该标签，说明 Windows/Office/AxMath 依赖、安装方法、变更和已知限制，再附上检查通过的分发文件。仓库工作流不会自动推送或发布任何内容。

如果仓库与版本标签已经存在，支持 stdio 的客户端也可以使用 uv 从 Git 启动，例如：

```powershell
uvx --from git+https://github.com/YOUR_GITHUB_USER/axmath-mcp.git@vX.Y.Z axmath-mcp
```

这种方式不会自动运行仓库中的 Codex 注册脚本，首次使用优先采用克隆仓库的教程。GitHub 分发无需先发布到 PyPI；语法见 [uv 工具安装文档](https://docs.astral.sh/uv/guides/tools/)。

## 参考资料

- [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk)
- [AxMath 官方站](https://www.amyxun.com/)
- [AxMath 输出与 Word 插件](https://axmath.gitbooks.io/axmath-docs-en/content/6._equation_output_and_word_plugin.html)
- [Word Application.Run](https://learn.microsoft.com/en-us/office/vba/api/word.application.run)
