## Office MCP Assistant

一个开箱即用的 Office 文件智能助手项目，整合了 Excel、PowerPoint、Word 三个 MCP 服务，支持在 Cursor 中通过自然语言对话处理/生成办公文档。所有读写默认统一到项目根目录下的 `workspace` 目录，保证可控、可复现。

### 依赖mcp服务仓库地址
- word: https://github.com/GongRzhe/Office-Word-MCP-Server
- excel: https://github.com/haris-musa/excel-mcp-server
- ppt: https://github.com/GongRzhe/Office-PowerPoint-MCP-Server

### 能力概览
- Excel：读取/写入单元格与区间、格式化、公式、图表、数据透视表、表格与验证、批量导入导出
- PowerPoint：新建/编辑演示文稿、添加文本/图片/形状、图表、模板与专业设计、过渡与母版管理
- Word：新建/编辑文档、段落与样式、表格、批注、脚注/尾注、导出 PDF（需本地 Word 支持）
- 统一工作目录：一切文件读写均在 `workspace` 中完成；如文件不在此目录会明确提示

### 环境准备
1) 安装 Python 3.10+ 与 pip
2) 安装cursor
详见https://yb.tencent.com/s/6glutjUYCmw7

### 快速开始
1) 进入 `mcp_server` 目录并执行依赖安装与初始化
打开终端（ctrl+`）,输入以下命令后回车

```bash
cd mcp_server
python setup_cursor_mcp.py
```

该脚本会：
- 创建目录：`workspace`（以及必要的辅助目录）
- 安装并校验依赖（`requirements.txt`）
- 生成 `cursor_mcp_config.json`：包含 Excel/PowerPoint/Word MCP 服务配置，统一将读写路径指向 `workspace`

2）在 Cursor 中配置 MCP
将生成的 `cursor_mcp_config.json` 的内容合并到你的 Cursor MCP 配置中（通常位于 `~/.cursor/mcp.json`）。
也可以在cursor中 点击右上角设置 -> MCP & integrations -> New MCP Service,覆盖整个json。

3）（可选）配置你自己的apikey
cursor中 点击右上角设置 -> Models -> API Keys,配置你自己的apikey。

4）（可选）添加 Workspace Rules（项目规则）
项目根目录下创建：
.cursor\rules\office-role.mdc
填入根目录下pompt.txt的内容

或者cursor中 点击右上角设置 -> Rules & Memories -> ProjectRules -> Add Role -> 输入office-role并回车,填入根目录下pompt.txt的内容

### 使用说明
1）将你要操作的文件放入项目根目录的workspase文件夹下
2）在右侧窗口与cursor对话，说明你的需求（哪个文件，要做什么；新建一个什么文件，加入什么内容;总结一下哪个文件的主要内容等）
3）输入框右侧的圆形进度条是此次对话的上下文，快满了就代表需要新开一个聊天了（点击聊天框上侧的加号），不然ai会输出质量较低的内容。
4）注意不要在让ai操作某个文件的同时使用wps等软件打开它，会导致文件被占用，ai无法查看。推荐使用cursor的插件来在ai操作的同时查看这些文件，如Office Viewer
5）推荐打开cursor的自动保存功能 左上角File -> Auto Save，重要文件一定要备份

### 故障排查
- 无法写入 Excel：确保未在本地 Excel 中打开该文件；必要时脚本会生成一个 `*_updated` 文件供替代
- 部分 Word 功能（如导出 PDF）依赖本机 Office：请确保已安装并可调用
- 端口/传输问题：本项目默认使用 stdio；如你切换到 SSE/HTTP，请按各服务 README 配置环境变量
- 依赖安装失败：手动执行 `python -m pip install -r requirements.txt`，并检查 Python 版本与网络

### 目录说明
- `workspace/`：统一的输入/输出目录（你需要处理或新建的文件都在这里）
- `excel-mcp-server-main/`、`Office-PowerPoint-MCP-Server-main/`、`Office-Word-MCP-Server-main/`：三类 MCP 服务实现
- `setup_cursor_mcp.py`：初始化脚本（创建目录、安装依赖、生成 MCP 配置）
- `cursor_mcp_config.json`：生成的 Cursor MCP 配置（可合并到本机 Cursor 配置）
- `requirements.txt`：运行所需依赖


