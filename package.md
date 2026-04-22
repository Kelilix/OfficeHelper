# OfficeHelper 打包与部署方案

> 文档版本：2026-04-22
> 目标读者：开发者 / 部署工程师

---

## 一、背景与目标

### 1.1 当前架构

OfficeHelper 采用**双进程独立部署**模式：

- **前端**：Tauri 桌面外壳（Tauri 2 + Rust），负责窗口管理和子进程生命周期
- **后端**：Python FastAPI 服务，通过 PyInstaller 打包为独立 exe
- **通信**：HTTP `http://127.0.0.1:8765`，同机器进程间通信

这种方式的优势：
- 前端与后端完全解耦，各自独立构建
- Tauri 自动处理后端的子进程管理（启动/停止/重启）
- PyInstaller 打包的 Python exe 零依赖（无需用户安装 Python）
- 通过 NSIS / MSI 安装器分发，用户体验与原生桌面应用一致

### 1.2 目标格式

| 项目 | 要求 |
|------|------|
| 分发格式 | Windows 安装包（NSIS `.exe` 或 MSI `.msi`） |
| 用户体验 | 双击安装，自动完成所有配置，无需手动安装 Python |
| 权限需求 | 普通用户权限（`HKEY_CURRENT_USER`，无需管理员） |
| 外网依赖 | 无（纯内网运行，LLM API 除外） |

---

## 二、系统架构

### 2.1 双进程架构

```
┌──────────────────────────────────────────────┐
│  Process 1: Tauri Desktop Shell (Rust)       │
│  exe: office-helper-front.exe                │
│                                              │
│  - 420×680 窗口，原生边框，可调整大小          │
│  - React WebView（渲染前端 UI）               │
│  - 子进程管理：启动/停止 Python 后端           │
│  - TCP 心跳检测：http://127.0.0.1:8765       │
│  - Tauri IPC：start/stop/restart/health      │
└──────────────────────┬───────────────────────┘
                       │ spawn child
                       │ HTTP 127.0.0.1:8765
                       ▼
┌──────────────────────────────────────────────┐
│  Process 2: Python FastAPI 后端               │
│  exe: OfficeHelperBackend.exe                │
│                                              │
│  - uvicorn HTTP 服务器                        │
│  - Word COM 操作（win32com）                 │
│  - LLM 调用（OpenAI/Anthropic/Qwen 等）      │
│  - AI 意图解析 + action 执行                  │
└──────────────────────────────────────────────┘
```

### 2.2 前端（Tauri）职责

| 功能 | 实现 |
|------|------|
| 窗口管理 | 420×680，最小 360×500，居中，原生边框 |
| 子进程启动 | `Command::new()` + `spawn()`，工作目录为 exe 所在目录 |
| 后端路径查找 | 按优先级：bundle 路径 → dev 路径 → dev fallback |
| 生命周期绑定 | `on_window_event(CloseRequested)` 时调用 `stop()` 杀死后端 |
| 心跳检测 | `TcpStream::connect("127.0.0.1:8765")`，1 秒超时 |

### 2.3 后端（Python）职责

| 功能 | 实现 |
|------|------|
| HTTP 服务 | uvicorn，`http://0.0.0.0:8765` |
| Word 连接 | `win32com.client.GetObject("Word.Application")`，复用已有实例 |
| AI 推理 | `LLMService`，路由到 OpenAI / Anthropic / Qwen 等 |
| 资源定位 | `sys._MEIPASS`（PyInstaller）→ `_internal/` |
| 调试日志 | 写入 `%TEMP%\oh_debug.log` |

### 2.4 API 路由

| 方法 | 路径 | 说明 |
|------|------|------|
| `POST` | `/api/chat` | 发送聊天消息（带多轮会话历史） |
| `GET` | `/api/word/status` | Word 连接状态 |
| `POST` | `/api/word/connect` | 主动连接 Word |
| `POST` | `/api/word/disconnect` | 断开 Word |
| `GET` | `/api/word/documents` | 所有打开的 Word 文档 |
| `GET` | `/api/sessions` | 会话列表 |
| `GET` | `/api/chat/history` | 指定会话历史 |
| `DELETE` | `/api/chat/clear` | 清除会话 |

---

## 三、运行时路径解析

### 3.1 Tauri 查找后端的顺序

`lib.rs` 中 `find_backend_path(exe_dir)` 按以下顺序搜索：

```
第 1 优先级：CARGO_MANIFEST_DIR 环境变量（仅 cargo run 开发模式）
  front/src-tauri/target/release/
    → .. → .. → .. → dist/OfficeHelperBackend/OfficeHelperBackend.exe

第 2 优先级：Tauri bundle 布局
  <前端exe>/resources/OfficeHelperBackend/OfficeHelperBackend.exe

第 3 优先级：扁平布局
  <前端exe>/OfficeHelperBackend/OfficeHelperBackend.exe

第 4 优先级：开发 fallback
  front/src-tauri/target/release/
    → .. → .. → .. → .. → dist/OfficeHelperBackend/OfficeHelperBackend.exe
```

### 3.2 后端定位资源文件的路径

PyInstaller 打包后，运行时关键路径：

| 变量 | 值（exe 模式） |
|------|----------------|
| `sys._MEIPASS` | `dist/OfficeHelperBackend/_internal/` |
| `sys.executable` | `dist/OfficeHelperBackend/OfficeHelperBackend.exe` |
| `os.getcwd()` | `dist/OfficeHelperBackend/` |
| `__file__` | `_internal/main.py`（编译进 exe） |

```python
def _get_bundle_root():
    if getattr(sys, "_MEIPASS", None):
        return Path(sys._MEIPASS)      # exe 模式 → _internal/
    return Path(__file__).parent.parent  # dev 模式 → 项目根
```

---

## 四、打包架构

### 4.1 三步构建流程

```
Step 1: PyInstaller  ──→  dist/OfficeHelperBackend/
Step 2: Vite          ──→  front/dist/
Step 3: Tauri/Cargo   ──→  front/src-tauri/target/release/bundle/
```

**Step 1 — PyInstaller 打包 Python 后端**

```batch
cd D:\soft\0project\test\OfficeHelper
D:\soft\python3.14\python.exe -m PyInstaller OfficeHelperBackend.spec --clean
```

- 输出：`dist/OfficeHelperBackend/OfficeHelperBackend.exe`
- `console=False`：GUI 模式，不显示控制台窗口
- 手动包含 Python 核心 DLL 到打包根目录（Tauri 只复制 exe 本身）
- `hiddenimports`：显式声明 PyInstaller 无法自动检测的模块
- `datas`：将 `skills/`、`core/`、`config.json` 打入 `_internal/`

**Step 2 — Vite 构建 React 前端**

```batch
cd front && npm run build
```

- 输出：`front/dist/index.html` + `assets/`

**Step 3 — Tauri 打包桌面应用**

```batch
cd front && npm run tauri build
```

- 读取 `tauri.conf.json` 中 `"resources": ["../../dist/OfficeHelperBackend/*"]`
- 将 `dist/OfficeHelperBackend/` 复制到 `target/release/`
- 输出 NSIS / MSI 安装包到 `bundle/`

### 4.2 打包后目录结构

```
front/src-tauri/target/release/
├── office-helper-front.exe       # Rust 主程序
│
├── OfficeHelperBackend/          # ← Tauri bundle 资源（来自 dist/）
│   ├── OfficeHelperBackend.exe  # Python 后端
│   ├── python314.dll            # ← 根目录 DLL（Tauri 可见）
│   ├── python3.dll
│   ├── vcruntime140.dll
│   ├── vcruntime140_1.dll
│   └── _internal/               # Python 运行时
│       ├── base_library.zip
│       ├── core/
│       ├── skills/
│       ├── config.json
│       └── *.pyd
│
└── bundle/
    ├── nsis/
    │   └── OfficeHelper_1.0.0_x64-setup.exe
    └── msi/
        └── OfficeHelper_1.0.0_x64_en-US.msi
```

### 4.3 DLL 定位机制

Python.exe 搜索 DLL 的顺序：

```
1. exe 同级目录            ← DLL 现在在这里（spec post-build 复制）
2. exe/_internal/          ← PyInstaller 默认位置
3. PATH 环境变量
4. Windows/System32
```

spec 文件底部 post-build 脚本负责将 DLL 从 `_internal/` 复制到根目录：

```python
for dll in ("python314.dll", "python3.dll", "vcruntime140.dll", "vcruntime140_1.dll"):
    src = internal_dir / dll
    dst = dist_root / dll
    if src.exists():
        shutil.copy2(src, dst)
```

### 4.4 Tauri 资源绑定机制

`tauri.conf.json` 中：

```json
"bundle": {
  "resources": [
    "../../dist/OfficeHelperBackend/*"
  ]
}
```

Tauri 在 `npm run tauri build` 时执行：
1. 编译 Rust 代码，生成 `office-helper-front.exe`
2. 将 `dist/OfficeHelperBackend/` 完整复制到 `target/release/OfficeHelperBackend/`
3. 将 `bundle/resources` 中的文件路径硬编码进 exe 资源表
4. NSIS / MSI 安装器将所有内容打包为安装包

运行时，Tauri 读取资源表，将资源解压到 exe 同级目录（或临时目录），确保相对路径正确。

---

## 五、后端 uvicorn 启动特殊处理

PyInstaller 打包后，后端 `main.py` 中 uvicorn 启动做了以下适配：

### 5.1 控制台分配（console=False 修复）

```python
if sys.stderr is None:
    kernel32.AllocConsole()
    sys.stdout = sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)
```

`console=False` 时 Python 没有标准流，`sys.stderr = None`，`StreamHandler` 写日志时触发 `AttributeError: 'NoneType' object has no attribute 'write'`。

### 5.2 日志 handler 清除

```python
for _name in ("uvicorn", "uvicorn.error", "uvicorn.access", "uvicorn.asgi"):
    logging.getLogger(_name).handlers.clear()
    logging.getLogger(_name).propagate = False
```

清除所有 uvicorn logger 的 handler，防止残留的 `StreamHandler` 尝试写入 `sys.stderr = None`。

### 5.3 直接传 app 实例 + lifespan=off

```python
config = uvicorn.Config(
    app,                    # 直接传实例，跳过 import_from_string
    host="0.0.0.0",
    port=8765,
    lifespan="off",         # 禁用 lifespan 框架
    log_config=None,
)
```

- **传实例而非字符串**：避免 `import_from_string("pservice.main:app")` 在打包后路径解析失败
- **lifespan="off"**：避免 uvicorn 的 `BaseException` handler 将 `SystemExit(1)` 当作"lifespan 不支持"静默吞掉

### 5.4 skills 提取到临时目录

```python
_extracted_root = _extract_resources()  # 提取到 %TEMP%\OfficeHelperBackend\
if _is_frozen and str(_extracted_root) not in sys.path:
    sys.path.insert(0, str(_extracted_root))
```

`_internal/` 中的 skills 目录被提取到 `%TEMP%\OfficeHelperBackend\skills\`，并加入 `sys.path`，供 `importlib.util.spec_from_file_location()` 动态加载 skill 脚本。

---

## 六、构建环境需求

### 6.1 软件要求

| 组件 | 版本 | 路径 | 说明 |
|------|------|------|------|
| Python | 3.14+ | `D:\soft\python3.14\python.exe` | PyInstaller 运行环境 |
| Node.js | 18+ | `D:\soft\nodeJS\` | npm/pnpm 包管理器 |
| Rust | 1.70+ | `C:\Users\...\ .cargo\` | Rust 编译器 |
| MSVC | VS 2022 (14.50.35717) | `D:\soft\MSVC\...` | Tauri/Rust 编译必需 |
| PyInstaller | 6.x | pip 安装 | `pip install pyinstaller` |

### 6.2 MSVC 配置说明

Tauri 基于 Rust + Windows API，编译时必须找到 MSVC 工具链。`build.bat` 中内置了 PATH 配置：

```batch
set "PATH=%CARGO_HOME%\bin;D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\bin\Hostx64\x64;%PATH%"
```

若 Cargo 报错找不到编译器，首先检查 MSVC bin 目录是否在 PATH 中。

### 6.3 Python 依赖

```batch
# 核心依赖（Word COM + LLM 支持）
pip install -r requirements.txt

# Web 服务依赖（独立于 requirements.txt）
pip install -r pservice/requirements.txt

# 打包工具
pip install pyinstaller
```

### 6.4 构建工具路径（当前配置）

以下路径硬编码在 `build.bat` 和 `OfficeHelperBackend.spec` 中，如工具链位置变更需同步更新：

| 路径 | 用途 |
|------|------|
| `D:\soft\python3.14\python.exe` | PyInstaller 执行 |
| `D:\soft\python3.14\python314.dll` | 打包包含 |
| `D:\soft\python3.14\vcruntime140.dll` | 打包包含 |
| `D:\soft\nodeJS\npm.cmd` | npm/pnpm 执行 |
| `C:\Users\gaotianyu.35\.cargo\` | Rust 工具链 |
| `D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\` | MSVC 编译器 |

---

## 七、快速构建

### 7.1 自动化脚本（推荐）

```batch
cd D:\soft\0project\test\OfficeHelper
build.bat
```

日志输出到 `build.log`，出错时生成 `build_error.log`。

### 7.2 手动分步构建

**Step 1：** PyInstaller 打包 Python 后端

```batch
cd D:\soft\0project\test\OfficeHelper
D:\soft\python3.14\python.exe -m PyInstaller OfficeHelperBackend.spec --clean
```

**Step 2：** Vite 构建前端

```batch
cd D:\soft\0project\test\OfficeHelper\front
D:\soft\nodeJS\npm.cmd run build
```

**Step 3：** Tauri 打包

```batch
set "CARGO_HOME=C:\Users\gaotianyu.35\.cargo"
set "RUSTUP_HOME=C:\Users\gaotianyu.35\.cargo"
set "PATH=%CARGO_HOME%\bin;D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\bin\Hostx64\x64;%PATH%"

cd D:\soft\0project\test\OfficeHelper\front
D:\soft\nodeJS\npm.cmd run tauri build
```

### 7.3 开发调试

```batch
# 终端 1：启动 Python 后端
python pservice/main.py

# 终端 2：启动前端开发服务器
cd front && npm run dev
```

Tauri 开发模式（`cargo run`）通过 `CARGO_MANIFEST_DIR` 环境变量自动找到 `dist/OfficeHelperBackend/OfficeHelperBackend.exe`。

---

## 八、部署与分发

### 8.1 安装包

构建完成后，分发以下文件之一：

```
front/src-tauri/target/release/bundle/nsis/OfficeHelper_1.0.0_x64-setup.exe
front/src-tauri/target/release/bundle/msi/OfficeHelper_1.0.0_x64_en-US.msi
```

用户双击安装包，自动完成所有配置。

### 8.2 免安装运行（开发测试）

直接运行 `front/src-tauri/target/release/office-helper-front.exe`，后端位于同级的 `OfficeHelperBackend/` 子目录中。

### 8.3 安装后目录结构

```
用户机器（安装到 Program Files 或用户目录）
├── OfficeHelper/
│   ├── OfficeHelper.exe          # Tauri 主程序
│   ├── resources/                # Tauri 资源目录（可选）
│   │   └── OfficeHelperBackend/
│   │       ├── OfficeHelperBackend.exe
│   │       ├── python314.dll
│   │       └── _internal/
│   └── 其他资源...
```

---

## 九、调试指南

### 9.1 后端启动日志

exe 模式运行时，日志写入 `%TEMP%\oh_debug.log`：

```
sys._MEIPASS     = D:\...\dist\OfficeHelperBackend\_internal
sys.frozen       = True
sys.executable  = D:\...\dist\OfficeHelperBackend\OfficeHelperBackend.exe
[OK] imports succeeded
[run] sys.path[0] = C:\Users\...\Temp\OfficeHelperBackend\skills\word-page-operator\scripts
```

### 9.2 端口占用检查

```batch
netstat -ano | findstr :8765
```

### 9.3 常见错误

| 错误 | 原因 | 解决 |
|------|------|------|
| `SystemExit: 1` | uvicorn 内部异常被吞 | 确认已使用 `lifespan="off"` + 传 app 实例 |
| `AttributeError: 'NoneType' object has no attribute 'write'` | `sys.stderr = None` | 确认已清除 uvicorn logger handler |
| 后端找不到 | Step 3 未完成 | 检查 `target/release/OfficeHelperBackend/` 是否存在 |
| DLL 加载失败 | DLL 不在 exe 根目录 | 重新运行 PyInstaller，spec post-build 会复制 DLL |
| 前端无法连接后端 | 端口 8765 未监听 | 检查 `%TEMP%\oh_debug.log` 中 uvicorn 启动是否成功 |

### 9.4 重建清理

修改代码后重新打包前，清理构建缓存：

```batch
# 清理 PyInstaller 构建缓存
rd /s /q dist\OfficeHelperBackend
rd /s /q build

# 清理 Tauri 构建缓存
rd /s /q front\src-tauri\target\release\OfficeHelperBackend
rd /s /q front\src-tauri\target\release\bundle

# 重新打包
build.bat
```

---

## 十、架构变更历史

| 版本 | 日期 | 主要变更 |
|------|------|----------|
| v1 | 2026-04-16 | 初始方案：Python 单入口 + Word Web Add-in（manifest.xml） |
| v2 | 2026-04-22 | 重大重构：改用 Tauri + 双进程独立部署，移除 Word Web Add-in |
