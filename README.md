# OfficeHelper

OfficeHelper 是一个 Windows 桌面 AI 助手，通过自然语言对话操控 Microsoft Word。它由两个独立进程组成——Tauri 前端（Rust + React）和 Python 后端（FastAPI + Word COM）——运行在同一台机器上，通过 HTTP（`http://127.0.0.1:8765`）通信。

---

## 目录结构

```
OfficeHelper/
├── config.json                      # 全局配置（LLM、UI、Word 设置）
├── requirements.txt                  # Python 运行时依赖（核心 + Web 框架）
├── OfficeHelperBackend.spec         # PyInstaller 打包规格
├── BUILD_STEPS.txt                  # 三步构建详细说明
├── build.bat                        # 一键自动化构建脚本
│
├── pservice/                        # Python FastAPI 后端
│   ├── main.py                      # 入口：分配控制台 → 启动 uvicorn
│   ├── requirements.txt             # Web 专用依赖（fastapi, uvicorn, pydantic）
│   └── api/
│       ├── __init__.py             # FastAPI app 工厂（CORS 已启用）
│       ├── routes.py                # 所有 HTTP 路由（/api/chat、/api/word/* 等）
│       ├── service.py               # 全局单例：word_service + llm_service
│       └── action_registry.py       # action 名称 → Word COM 方法调度器
│
├── core/                            # Python 业务逻辑（前后端共用）
│   ├── config.py / settings.py      # 配置管理（读取/写入 config.json）
│   ├── word_connector.py            # WordConnector 单例（win32com COM 封装）
│   ├── llm_service.py               # LLMService（多provider：OpenAI/Anthropic/Ollama/Qwen）
│   ├── agent.py                     # WordAgent（意图解析 + action 执行）
│   ├── screenshot_manager.py
│   ├── format_analyzer.py
│   └── format_fixer.py
│
├── skills/                          # AI Skill 定义（Markdown 格式，含元数据 + action 列表）
│   ├── word-text-operator/          # 82 个 Range 级操作（字体、查找替换、书签、选区等）
│   │   ├── SKILL.md
│   │   └── scripts/
│   │       ├── word_text_operator.py   # facade 类
│   │       ├── word_base.py            # COM 连接基类
│   │       ├── word_range_navigation.py
│   │       ├── word_text_operations.py
│   │       ├── word_find_replace.py
│   │       ├── word_format.py
│   │       ├── word_bookmark.py
│   │       └── word_selection.py
│   ├── word-paragraph-operator/     # 57 个段落级操作（对齐、缩进、列表、批量等）
│   │   ├── SKILL.md
│   │   └── scripts/
│   │       └── word_paragraph_operator.py
│   └── word-page-operator/           # 49 个页面/节级操作（页边距、纸张、分栏、页眉页脚等）
│       ├── SKILL.md
│       └── scripts/
│           ├── word_page_operator.py
│           ├── word_page_operator_base.py
│           └── word_section_operator.py
│
└── front/                           # Tauri + React 桌面前端
    ├── package.json                  # npm 依赖（React 18、Fluent UI、Tauri API）
    ├── vite.config.ts               # Vite 开发服务器（端口 1420）
    └── src-tauri/
        ├── Cargo.toml               # Tauri 2、serde、tauri-plugin-shell
        ├── tauri.conf.json          # 打包配置：NSIS/MSI、资源路径、窗口尺寸
        ├── capabilities/default.json  # Shell 插件权限
        └── src/
            ├── main.rs              # 入口 → setup() → run()
            └── lib.rs               # BackendProcess 管理器；Tauri IPC 命令（start/stop/restart/health）
```

---

## 系统架构

### 双进程模型

OfficeHelper 由两个独立运行的进程组成，通过 HTTP 通信：

```
┌─────────────────────────────────────────┐
│  Process 1: Tauri Desktop Shell (Rust) │  office-helper-front.exe
│                                         │
│  - 窗口管理（420×680，原生边框）           │
│  - React WebView（接收用户输入、显示结果）  │
│  - 子进程管理（启动/停止 Python 后端）       │
│  - TCP 心跳检测（端口 8765）               │
└──────────────────┬──────────────────────┘
                   │ spawn child process
                   │ HTTP 127.0.0.1:8765
                   ▼
┌─────────────────────────────────────────┐
│  Process 2: Python FastAPI 后端          │  OfficeHelperBackend.exe
│                                         │
│  - uvicorn HTTP 服务器                   │
│  - AI 意图理解（LLM 调用）                │
│  - Word COM 操作（win32com）             │
│  - Action 执行（action_registry）        │
└─────────────────────────────────────────┘
```

**前端（Tauri）职责：**
- 窗口管理：原生 420×680 可调整窗口
- 子进程生命周期：启动时启动后端，关闭时关闭后端
- TCP 心跳：`TcpStream::connect("127.0.0.1:8765")` 检测后端可用性
- Tauri IPC 命令：`start_backend`、`stop_backend`、`restart_backend`、`backend_health`

**后端（Python）职责：**
- HTTP API（`http://127.0.0.1:8765/api/*`）：聊天、Word 状态、文档列表
- Word COM 封装：`WordConnector` 单例，复用已有 Word 实例
- AI 推理：`LLMService` 路由到 OpenAI / Anthropic / Qwen 等 provider
- Action 执行：`action_registry` 将 LLM 生成的 action 调度到 Word COM

**API 路由：**

| 方法 | 路径 | 说明 |
|------|------|------|
| `POST` | `/api/chat` | 发送聊天消息（包含多轮会话历史） |
| `GET` | `/api/word/status` | Word 连接状态 |
| `POST` | `/api/word/connect` | 主动连接 Word |
| `POST` | `/api/word/disconnect` | 断开 Word |
| `GET` | `/api/word/documents` | 所有打开的 Word 文档 |
| `GET` | `/api/sessions` | 会话列表 |
| `GET` | `/api/chat/history` | 指定会话历史 |
| `DELETE` | `/api/chat/clear` | 清除会话 |

**一次对话的处理流程（`/api/chat`）：**

```
1. 接收 ChatRequest（message、selection_text、document_name、session_id）
2. 自动重连 Word（若未连接）
3. 初始化 WordTextOperator，捕获当前格式状态（字体 + 段落）
4. 收集文档统计（字数、段落数、节数、页数）
5. Plan 阶段：LLM 决定使用哪个 skill、如何拆解为 steps
6. Execute 阶段（每 step）：
   - LLM 根据 skill 内容生成 action JSON
   - action_registry.execute_action() 调度到 Word COM
   - 结果追加到 session 历史
7. 返回 ChatResponse（摘要 + 每步执行结果）
```

---

## 运行时架构

### 开发模式

```
终端 A: python pservice/main.py        → FastAPI 监听 http://127.0.0.1:8765
终端 B: cd front && npm run dev       → Vite 开发服务器 http://localhost:1420
```

前端开发服务器（`localhost:1420`）与后端（`127.0.0.1:8765`）同源策略兼容，Tauri WebView 通过 CSP 允许连接 `http://127.0.0.1:*` 和 `http://localhost:*`。

### 打包后运行模式

```
用户双击 OfficeHelper.exe
  → Rust lib.rs setup() 调用 BackendProcess.start()
  → find_backend_path() 按顺序查找后端：
      1. <前端exe>/OfficeHelperBackend/OfficeHelperBackend.exe  (生产路径)
      2. dist/OfficeHelperBackend/OfficeHelperBackend.exe      (开发路径)
      3. <前端exe>/resources/OfficeHelperBackend/...          (Tauri bundle)
      4. 4层 .. 回退路径                                     (开发 fallback)
  → Command::spawn() 启动后端子进程
  → 后端 main.py：
      - AllocConsole() 分配隐藏控制台（console=False 需要）
      - 直接传 app 实例给 uvicorn.Server().run()
      - uvicorn 监听 0.0.0.0:8765
  → 前端 React（在 WebView 中）调用 http://127.0.0.1:8765/api/*
```

### WordConnector 连接策略

`core/word_connector.py` 使用 `pythoncom.CoInitialize()` + `win32com.client.GetObject("Word.Application")`：
- 优先复用已有 Word 进程（避免弹出多个 Word 窗口）
- 备用 `Dispatch()` 创建新实例
- 文档操作前自动备份到 `~/.office_helper/backups/`（保留最近 10 份）
- `undo_manager` 支持批量撤销（最多 20 步）

---

## 打包架构

### 三步构建流程

```
Step 1: PyInstaller      →  dist/OfficeHelperBackend/
Step 2: Vite             →  front/dist/
Step 3: Tauri/Cargo      →  front/src-tauri/target/release/bundle/
```

**Step 1 — PyInstaller 打包 Python 后端**

```bash
D:\soft\python3.14\python.exe -m PyInstaller OfficeHelperBackend.spec --clean
```

输出：`dist/OfficeHelperBackend/OfficeHelperBackend.exe`

关键机制：
- `console=False`：GUI 模式，不显示控制台窗口
- 手动包含 Python 核心 DLL（`python314.dll`、`python3.dll`、`vcruntime140.dll`、`vcruntime140_1.dll`）到打包根目录
- `hiddenimports`：显式声明 PyInstaller 无法自动检测的模块（win32com 全套、uvicorn/starlette/fastapi/pydantic、openai/anthropic/requests/httpx/psutil/jinja2/markdown/PIL）
- 打包后 post-build 脚本将 DLL 从 `_internal/` 复制到根目录（Tauri 只复制 exe 本身，DLL 必须在根目录才可被发现）
- `datas`：将 `skills/`、`core/`、`config.json` 打入 `_internal/`

**Step 2 — Vite 构建 React 前端**

```bash
cd front && npm run build
```

输出：`front/dist/index.html` + `assets/`

**Step 3 — Tauri 打包桌面应用**

```bash
cd front && npm run tauri build
```

关键机制：
- `tauri.conf.json` 中 `"resources": ["../../dist/OfficeHelperBackend/*"]` 指令 Tauri 将 PyInstaller 产物复制到 bundle 目录
- NSIS / MSI 安装器将 `office-helper-front.exe` + `OfficeHelperBackend/` 打包在一起

### 打包后目录结构

```
front/src-tauri/target/release/
├── office-helper-front.exe          # Rust 主程序（Tauri）
│
├── OfficeHelperBackend/             # ← Tauri bundle resources 复制自 dist/
│   ├── OfficeHelperBackend.exe     # Python 后端
│   ├── python314.dll               # ← 根目录 DLL（Tauri 可发现）
│   ├── python3.dll
│   ├── vcruntime140.dll
│   ├── vcruntime140_1.dll
│   └── _internal/                  # Python 运行时 + 依赖
│       ├── base_library.zip
│       ├── core/
│       ├── skills/
│       ├── config.json
│       └── *.pyd
│
└── bundle/
    ├── nsis/
    │   └── OfficeHelper_1.0.0_x64-setup.exe   # Windows 安装包
    └── msi/
        └── OfficeHelper_1.0.0_x64_en-US.msi    # MSI 安装包
```

---

## 构建环境需求

### 软件要求

| 组件 | 版本要求 | 说明 |
|------|----------|------|
| **Python** | 3.14+ | 位于 `D:\soft\python3.14\python.exe` |
| **Node.js** | 18+ | 位于 `D:\soft\nodeJS\` |
| **Rust** | 1.70+ | 位于 `C:\Users\gaotianyu.35\.cargo\` |
| **MSVC 工具链** | Visual Studio 2022 (14.50.35717) | 位于 `D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\bin\Hostx64\x64` |
| **PyInstaller** | 6.x | `pip install pyinstaller` |

> **注意**：Rust 编译器（`rustc`）和 MSVC 是编译 Tauri（Rust）所必需的。仅有 Python 无法构建前端。

### MSVC 配置说明

Tauri 底层使用 Rust + Windows API，需要 MSVC（Microsoft Visual C++）工具链。构建时必须将 MSVC bin 目录加入 `PATH`，否则 Cargo 找不到 C++ 编译器：

```batch
set "PATH=%CARGO_HOME%\bin;D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\bin\Hostx64\x64;%PATH%"
```

`build.bat` 已内置此配置，无需手动设置。

### Python 依赖安装

```bash
# 核心依赖（包含 Word COM + LLM 支持）
pip install -r requirements.txt

# Web 服务专用依赖（独立于 requirements.txt，供后端 spec 使用）
pip install -r pservice/requirements.txt

# PyInstaller
pip install pyinstaller
```

---

## 快速构建

### 方式一：自动化脚本（推荐）

```batch
cd D:\soft\0project\test\OfficeHelper
build.bat
```

`build.bat` 按顺序执行三步构建，记录详细日志到 `build.log`，错误时自动生成 `build_error.log`。

### 方式二：手动分步构建

**Step 1：** 打包 Python 后端

```batch
cd D:\soft\0project\test\OfficeHelper
D:\soft\python3.14\python.exe -m PyInstaller OfficeHelperBackend.spec --clean
```

**Step 2：** 构建 React 前端

```batch
cd D:\soft\0project\test\OfficeHelper\front
D:\soft\nodeJS\npm.cmd run build
```

**Step 3：** 打包 Tauri 应用

```batch
set "CARGO_HOME=C:\Users\gaotianyu.35\.cargo"
set "RUSTUP_HOME=C:\Users\gaotianyu.35\.cargo"
set "PATH=%CARGO_HOME%\bin;D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\bin\Hostx64\x64;%PATH%"

cd D:\soft\0project\test\OfficeHelper\front
D:\soft\nodeJS\npm.cmd run tauri build
```

### 开发调试

```batch
# 终端 1：启动 Python 后端
python pservice/main.py

# 终端 2：启动前端开发服务器
cd front && npm run dev
```

Tauri 开发模式（`cargo run`）会自动从 `dist/OfficeHelperBackend/` 加载后端。

---

## 配置

所有配置通过 `config.json` 管理（支持热重载）：

```json
{
  "llm": {
    "provider": "qwen",
    "api_key": "your-key",
    "model": "qwen3.5-plus",
    "base_url": "https://coding.dashscope.aliyuncs.com/v1",
    "temperature": 0.7,
    "max_tokens": 8000
  },
  "ui": {
    "theme": "light",
    "window_width": 1200,
    "window_height": 800,
    "language": "zh-CN"
  },
  "word": {
    "auto_open": true,
    "backup_enabled": true,
    "backup_dir": ""
  },
  "debug": false
}
```

支持 LLM provider：`openai`（含兼容端点）、`anthropic`、`qwen`（阿里云 DashScope）、`ollama`（本地模型）。

---

## 调试

### 后端启动日志

exe 模式启动时写入 `%TEMP%\oh_debug.log`，包含：
- `sys._MEIPASS`（打包根目录）
- 模块导入状态
- uvicorn 启动过程

### PyInstaller 打包后 uvicorn 特殊处理

`main.py` 中的 uvicorn 启动做了以下适配：

```python
# 1. console=False 时手动分配控制台
if sys.stderr is None:
    kernel32.AllocConsole()
    sys.stdout = sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)

# 2. 清除 uvicorn logger handler，防止 sys.stderr=None 触发 AttributeError
for _name in ("uvicorn", "uvicorn.error", "uvicorn.access", "uvicorn.asgi"):
    logging.getLogger(_name).handlers.clear()
    logging.getLogger(_name).propagate = False

# 3. 直接传 app 实例 + lifespan="off"
config = uvicorn.Config(app, host="0.0.0.0", port=8765,
                        lifespan="off", log_config=None)
```

### 端口占用检查

```batch
netstat -ano | findstr :8765
```
