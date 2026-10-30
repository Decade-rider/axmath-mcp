# 测试与验证

本页面向维护者，记录可重复执行的检查方法与已完成的实机验证。日常安装和使用请先看 [README](../README.md)。

## 检查层次

| 检查 | 需要 Office/AxMath | 能说明什么 |
| --- | --- | --- |
| 单元测试、Ruff、包内容检查 | 否 | 输入校验、对象管理、配置写入及分发边界是否符合预期 |
| MCP 握手与环境诊断 | Windows；诊断需要已安装的软件 | 服务能否启动、暴露工具并识别本机环境 |
| Office 实际调用 | 是 | 公式能否生成、读取、修改、定位和保存 |
| 保存重开与视觉检查 | 是 | 原生对象是否保留，以及基线、行距、裁切和对齐是否符合预期 |

单元测试通过不能代替真实 Office 调用；公式导出 LaTeX 成功，也不能代替页面或幻灯片的视觉检查。

## 单元测试和分发检查

从项目根目录执行：

```powershell
uv sync --locked --default-index https://pypi.org/simple
uv run --no-sync ruff check src tests scripts
uv run --no-sync ruff format --check src tests scripts
uv run --no-sync pytest -q
uv build --default-index https://pypi.org/simple
uv run --no-sync python scripts/check_dist.py
```

GitHub Actions 使用 Windows 和 Python 3.12 执行这些检查，不会在托管 runner 上运行需要 Word 或 AxMath 授权的实机测试，也不会自动发布包。

## Word 实机测试

先确认本机 Word 可以手动插入并编辑 AxMath 公式。测试创建独立文档，不修改已有文档；测试期间请避免同时操作 Office 或 AxMath。

```powershell
# 只检查 MCP 握手、工具列表和环境状态
uv run --no-sync python scripts/smoke_mcp.py

# 新建文档，验证插入、读取、更新与保存
uv run --no-sync python scripts/smoke_mcp.py --exercise-word

# 验证自动编号、14 磅字号保持、独立公式和固定编号切换
uv run --no-sync python scripts/smoke_layout_mcp.py --exercise-word
```

输出文件和 JSON 记录位于 `artifacts/`。该目录不会进入 Git 仓库或分发包；日志中可能包含本机文件路径，分享之前应检查内容。

## PowerPoint 实机测试

PowerPoint 功能也需要 Word 与 AxMath 的 Word 插件。测试创建两页独立演示文稿，不修改已有演示文稿：

```powershell
# 仅检查统一 MCP 的 25 个工具与环境
uv run --no-sync python scripts/smoke_powerpoint_mcp.py

# 原生公式插入、读取、更新、移动、编号及保存
uv run --no-sync python scripts/smoke_powerpoint_mcp.py --exercise-powerpoint

# 将文件参数替换为上一命令输出的真实路径
uv run --no-sync python scripts/verify_powerpoint_file.py C:\path\to\saved-test.pptx --read-latex --expected-slides 2 --expected-equations 2
```

最后一个命令使用保存文件的独立只读副本，逐页导出 PNG，检查对象边界、固定编号的双向关联、字号和坐标，并可从原生 OLE 中重新读取 LaTeX。原文件的 SHA256 必须保持不变。输出目录默认唯一；指定的已有输出目录不会被覆盖。完成后应人工查看每张 PNG，不把结构检查当作视觉检查。

### 0.3.0 本机验证

2026-10-30，Windows、Microsoft 365 Word / PowerPoint 16.0、AxMath 2.7061、Python 3.12。通过 343 项自动化测试、Ruff、统一服务握手和以下实机检查：

| 检查 | 结果 |
| --- | --- |
| PowerPoint 插入 24 磅分式并读取 LaTeX | 原生 `Equation.AxMath`，源码正确 |
| 修改为积分公式 | 公式 ID、位置、字号和固定编号保留 |
| 移动公式 | 关联编号一起移动，保存后的 XML 坐标一致 |
| 越界位置及错误 SlideID | 请求被拒绝，原公式保持不变 |
| 第二页矩阵 | 独立原生对象，两个幻灯片分别可列出 |
| 保存 PPTX | 两个原生 OLE、两个嵌入存储、无 Word 文档包装对象 |
| 重开只读副本 | 两条公式均可重新导出 LaTeX，编号关联保留 |
| 原生 PowerPoint 图片导出 | 中文和英文标题可读，积分、矩阵及编号未观察到裁切或重叠 |
| PPTX 结构 | 通过 PPTX 技能的 OOXML 校验器；Windows 下以 UTF-8 模式运行 |
| Word 回归 | 在统一服务中完成插入、读取、修改、选区校验和保存 |

PowerPoint 会为一个 OLE 对象保存 Choice/Fallback 两种兼容 XML 表示，因此结构检查按逻辑图形和嵌入关系计数，不能简单用 `p:oleObj` 元素个数代替公式数量。精选图片：[固定编号](images/powerpoint-numbered.png)、[矩阵](images/powerpoint-matrix.png)。

验证限于顶层、未旋转、未分组的原生公式对象。文本框内真正行内公式、自动 SEQ 编号、动画／动作及连接器绑定对象的自动替换不属于已支持范围。交互式打开 AxMath 编辑器的入口未纳入无人值守回归；用户需自行完成并关闭编辑器。历史 Word 混排验证见下文。

## Word 混排与编号的视觉验证

```powershell
uv run --no-sync python scripts/build_layout_assets.py --output artifacts/layout-assets.docx
uv run --no-sync python scripts/verify_layouts.py --assets artifacts/layout-assets.docx --output artifacts/layout-validation
```

第一步通过真实 AxMath 转换生成七个原生公式对象，以及记录导出 LaTeX、字号和宽高的同名 JSON。仅验证用例复用这些对象；生产工具仍逐次执行转换。已有素材文件不会被覆盖，重新生成时请指定新的输出路径。

第二步生成四页 Word 文档、PDF 与 `verification.json`，覆盖中文混排、英文混排、独立公式和带编号公式。程序检查对象、段落、序号字段和保存后的 XML；仍需逐页查看 Word 导出的页面图像。

检查时应确认：

- 公式仍是 `Equation.AxMath` OLE 对象，能够继续用 AxMath 编辑。
- 分式、根号、积分上下限、矩阵括号和分段函数没有裁切。
- 中英文正文、公式与标点的基线自然，换行后没有重叠。
- 独立公式位于正文宽度中点，编号靠右，并与公式纵向对齐。
- 新增、删除或移动公式后，自动序号按文档顺序更新；固定编号保持原样。
- 保存重开后，公式对象和可更新的编号字段仍然存在。

Word 的 OLE 范围包含隐藏字段，`Range.Font.Size` 或 `Range.Font.Position` 可能返回未定义值 `9999999`。验证使用实际对象所在 XML run 的字体信息、样式继承及渲染结果，不把这一混合范围值当作测量结果。

## 已完成的 Word 基线验证

以下为 **0.2.0 版本于 2026-10-30 完成的历史验证**，不能自动扩展为后续版本或其他 Office 应用的测试结果。

环境：AxMath 2.7061、Microsoft 365 Word 16.0、Python 3.12、MCP Python SDK 1.30.0。该轮通过 202 项自动化测试（包括安装注册检查）、Ruff、真实 MCP 调用及独立环境中的 wheel 安装检查。

| 场景 | 覆盖内容 | 已观察结果 |
| --- | --- | --- |
| 中文行内 3 例 | 宋体 12 磅正文与上下标、分式、根式 | 基线保留，无可见裁切或重叠；固定行距可改为最小行距 |
| 英文行内 3 例 | Times New Roman 12 磅正文与同三类公式 | 基线、标点和换行后的行距检查通过 |
| 独立居中 4 例 | 求和、高斯积分、矩阵、分段函数 | 独占段落，公式上下内容完整 |
| 自动编号 4 例 | 同段落公式居中、编号右对齐 | `(1)` 至 `(4)` 连续，四个编号右边界一致 |
| 编号维护 | 前插、删除、替换、保存重开 | 重新编号正确，替换保留编号，重开后 SEQ 域仍可更新 |
| 真实 MCP 排版切换 | 14 磅编号公式 → 修改 → 独立公式 → `(A.1)` | 字号及公式 ID 保留，无重复对象或遗留自动编号 |

四页文档已由本机 Word 导出并逐页查看。14 个对象的原生半磅基线偏移在保存后全部保留，含公式段落的有效网格对齐均已关闭。精选预览：[中文混排](images/chinese-inline.png)、[英文混排](images/english-inline.png)、[独立公式](images/display.png)、[自动编号](images/numbered.png)。

这些结果针对上述环境、七种表达式和单栏正文。不同 AxMath/Office 版本、任意 LaTeX 命令及跨页超高公式需要单独验证。
