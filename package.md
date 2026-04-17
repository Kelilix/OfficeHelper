# OfficeHelper 打包与部署方案

> 文档版本：2026-04-16
> 目标读者：开发者 / 部署工程师

---

## 一、背景与约束

### 1.1 问题描述

Office LTSC 专业增强版 2024 是微软最新的批量授权版 Office（永久授权，非订阅）。与 Microsoft 365 订阅版不同，LTSC 版**默认禁用了用户侧的「上载我的加载项」UI 入口**，即使用户从「开始 → 加载项 → 我的加载项」进入，看到的也只有「应用商店」跳转，完全没有「上载我的加载项」按钮。

这导致：
- 手动引导用户「点几下按钮就能加载」的方式**完全失效**
- 必须使用**绕过 UI 的自动化方式**来完成 Word Web Add-in 的注册（sideload）

### 1.2 已知可行路径

用户已通过 `npm start`（`office-addin-dev-settings` 底层工具链）成功 sideload 过 wordassistant，说明注册表写入路径**完全通畅**。因此，打包方案必须基于**直接写注册表**的方式，而非依赖 UI 对话框。

### 1.3 约束条件

| 约束 | 说明 |
|---|---|
| 目标用户 | 普通用户（无 Node.js / 无管理员权限） |
| Word 版本 | Office LTSC 专业增强版 2024（可能涉及其他批量授权版） |
| Excel/PowerPoint | 暂不覆盖，仅针对 Word |
| 目标格式 | 单文件 EXE（无需安装 Python 环境） |
| 外网依赖 | 无（纯内网运行） |
| 注册表路径 | `HKEY_CURRENT_USER`（不需要管理员权限） |

---

## 二、系统架构

### 2.1 整体架构

```
用户运行 OfficeHelper.exe
    │
    ├─ 解压到临时目录（PyInstaller 自解压行为）
    │
    ├─ [1] 检查依赖环境
    │       ├─ 检测是否已安装 Office
    │       ├─ 检测 Python 环境（内部打包的 runtime）
    │       └─ 检测 Word 是否正在运行
    │
    ├─ [2] 启动后端服务（内嵌 Python 进程）
    │       ├─ 启动 FastAPI (uvicorn) 监听 127.0.0.1:8765
    │       ├─ 预加载 skill 模块
    │       └─ 保持进程运行
    │
    ├─ [3] 注册 Word Add-in
    │       ├─ 将 manifest.xml 复制到用户本地目录
    │       ├─ 从 manifest.xml 提取 addon_id
    │       ├─ 将 addon_id + manifest路径 写入注册表
    │       └─ 通知 Word 刷新加载项
    │
    ├─ [4] 启动前端（静默）
    │       └─ Word Add-in 页面通过 manifest 指向 localhost:3000
    │
    └─ [5] 启动 React 开发服务器（监听 localhost:3000）
            ├─ 启动 word-addin dev server (webpack dev server)
            └─ 保持进程运行
```

### 2.2 打包产物结构（运行态）

```
临时解压目录/
├── OfficeHelper.exe          # 主程序入口（PyInstaller bundle）
├── _internal/               # PyInstaller 内部文件
│   ├── python310.dll
│   ├── VCRUNTIME140.dll
│   └── ...
├── word-addin/              # React Web Add-in 源码
│   ├── package.json
│   ├── src/
│   └── dist/               # 构建产物（npm run build 输出）
├── pservice/               # FastAPI 后端源码（启动时用）
│   ├── main.py
│   └── api/
└── core/                   # 核心模块
```

### 2.3 关键路径说明

| 路径 | 用途 |
|---|---|
| `word-addin/wordassistant/manifest.xml` | Word Add-in 清单文件 |
| `pservice/` | FastAPI 后端服务源码 |
| `core/` | WordConnector、LLM 服务等核心模块 |
| `skills/` | AI Skill 定义文件（Markdown） |

---

## 三、Word Add-in 注册方案

### 3.1 注册表机制详解

Word Web Add-in 的 sideload 本质上是将 `{addon_id} → {manifest_absolute_path}` 写入注册表，Word 启动时读取该表并加载对应 Add-in。

**注册表路径：**
```
HKEY_CURRENT_USER\Software\Microsoft\Office\{OfficeVersion}\WEF\Developer
```

**值项：**
- 名称（Name）：`{addon_id}`（从 manifest.xml 的 `<Id>` 节点提取）
- 数据（Value）：manifest.xml 的**绝对路径**（必须是本地路径，Word 不支持远程 URL）

**Word 版本号映射：**

| Office 版本 | 主版本号 |
|---|---|
| Office 2016 / 2019 / 2021 / LTSC 2024 | `16.0` |
| Office 365 (订阅版) | `16.0`（相同） |

> 注：`16.0` 对应 Office 2016+ 的共同内核版本号，并非 Word 版本号本身。

### 3.2 两种注册方式对比

| 方式 | 优点 | 缺点 |
|---|---|---|
| **A. office-addin-dev-settings CLI** | 微软官方工具，逻辑完整 | 需要 Node.js 环境，普通用户不满足 |
| **B. 直接写注册表（winreg）** | 零依赖，纯 Python | 需要自己实现完整逻辑 |

**本方案选择方式 B**（`winreg`），因为 PyInstaller 打包后无法调用外部 Node.js 命令，且 `winreg` 是 Python 内置模块，无额外依赖。

### 3.3 manifest.xml 路径策略

Word 要求 manifest 路径为**本地绝对路径**，不允许使用远程 URL。因此 EXE 启动时必须将 manifest.xml 复制到用户本地的固定目录。

**目标目录优先级：**
1. `~/.office_helper/addins/wordassistant/manifest.xml`（用户目录，推荐）
2. `~/.office_helper/` 作为基准目录（与 `config.json` 共用）

**关键要求：**
- manifest 中的 `<SourceLocation>` 必须是 HTTPS URL（`https://localhost:3000/taskpane.html`）
- 注册表中的路径必须是 manifest.xml 的**本地绝对路径**

### 3.4 注册函数设计

```python
# ============================================================
# 文件：pservice/api/addin_manager.py
# 用途：Word Add-in 的自动化注册/卸载（绕过 UI）
# 依赖：仅 Python 内置库（winreg, xml.etree, pathlib）
# ============================================================

import winreg
import uuid
import os
import shutil
import hashlib
from pathlib import Path
from typing import Optional
import xml.etree.ElementTree as ET

# 注册表路径（对应 office-addin-dev-settings 底层行为）
_WEF_REG_KEY = r"Software\Microsoft\Office\{version}\WEF\Developer"

# manifest.xml 默认 ID（fallback）
_DEFAULT_ADDIN_ID = "45e162e8-0cf4-4b83-ac50-e6a0463719b1"


def get_office_version() -> str:
    """
    探测 Office 实际主版本号。
    优先从 ClickToRun 配置读取（Office 365 / LTSC 2024 使用此路径），
    失败则 fallback 到 16.0。
    """
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            0, winreg.KEY_READ
        )
        version, _ = winreg.QueryValueEx(key, "VersionToReport")
        winreg.CloseKey(key)
        # 返回主版本号，如 "16.0"
        return ".".join(version.split(".")[:2])
    except WindowsError:
        pass

    # 备选：HKCU
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
            0, winreg.KEY_READ
        )
        winreg.CloseKey(key)
        return "16.0"
    except WindowsError:
        return "16.0"


def get_wef_key_path() -> str:
    """构造完整的 WEF 注册表路径"""
    version = get_office_version()
    return _WEF_REG_KEY.format(version=version)


def extract_addin_id(manifest_path: str) -> str:
    """
    从 manifest.xml 提取 Add-in ID。
    支持统一清单格式（OfficeApp/Id）和旧格式。
    """
    tree = ET.parse(manifest_path)
    root = tree.getroot()

    # 统一清单格式（taskpaneapp / mailapp / contentapp）
    namespaces = [
        "http://schemas.microsoft.com/office/matml/manifest",
        "http://schemas.microsoft.com/office/taskpaneapp",
        "http://schemas.microsoft.com/office/mailappversionoverrides",
    ]

    for ns in namespaces:
        elem = root.find(f".//{{{ns}}}Id")
        if elem is not None and elem.text:
            return elem.text.strip()

    # 旧格式 / OfficeApp 根节点 id 属性
    for attr in ["{http://schemas.microsoft.com/office/matml/manifest}id",
                 "id"]:
        val = root.get(attr) or root.get(attr.replace("{http://schemas.microsoft.com/office/matml/manifest}", ""))
        if val:
            return val.strip()

    # fallback
    return _DEFAULT_ADDIN_ID


def get_addins_dir() -> Path:
    """
    获取 Add-in 存储目录（与 config.json 同目录）。
    返回 ~/.office_helper/addins/
    """
    base = Path.home() / ".office_helper" / "addins"
    base.mkdir(parents=True, exist_ok=True)
    return base


def copy_manifest_to_local(manifest_src: Path, addon_name: str = "wordassistant") -> Path:
    """
    将 manifest.xml 复制到用户本地目录。
    保持原文件名（manifest.xml），按 addon_name 组织子目录。
    """
    dest_dir = get_addins_dir() / addon_name
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / "manifest.xml"

    # 始终重新复制（确保是最新的）
    shutil.copy2(manifest_src, dest_path)
    return dest_path.resolve()


def register_addin(manifest_path: str) -> dict:
    """
    将 Word Add-in 注册到系统（写入注册表）。
    等价于 office-addin-dev-settings sideload <manifest.xml>

    Returns:
        dict with keys: success(bool), addon_id(str), manifest_path(str),
                        message(str), error(str)
    """
    result = {
        "success": False,
        "addon_id": "",
        "manifest_path": "",
        "message": "",
        "error": ""
    }

    manifest_path = str(Path(manifest_path).resolve())

    # 1. 解析 manifest，提取 addin_id
    try:
        addon_id = extract_addin_id(manifest_path)
    except Exception as e:
        result["error"] = f"manifest.xml 解析失败: {e}"
        return result
    result["addon_id"] = addon_id

    # 2. 复制 manifest 到用户本地目录
    try:
        local_manifest = copy_manifest_to_local(Path(manifest_path), "wordassistant")
        result["manifest_path"] = str(local_manifest)
    except Exception as e:
        result["error"] = f"manifest 复制失败: {e}"
        return result

    # 3. 写入注册表
    key_path = get_wef_key_path()
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValueEx(key, addon_id, 0, winreg.REG_SZ, str(local_manifest))
        winreg.CloseKey(key)
    except PermissionError:
        result["error"] = "权限不足，请确保以管理员身份运行（注册表写入需要 HKCU 写入权限）"
        return result
    except Exception as e:
        result["error"] = f"注册表写入失败: {e}"
        return result

    result["success"] = True
    result["message"] = (
        f"Add-in 注册成功！\n"
        f"  addon_id: {addon_id}\n"
        f"  manifest: {local_manifest}\n"
        f"请关闭所有 Word 窗口后重新打开 Word，功能区「开始」选项卡会出现「AI 助手」按钮。"
    )
    return result


def unregister_addin(addon_id: Optional[str] = None) -> dict:
    """
    从注册表删除指定 Add-in（卸载）。
    如果不提供 addon_id，从 manifest.xml 自动推断。
    """
    result = {"success": False, "error": ""}

    if addon_id is None:
        addins_dir = get_addins_dir() / "wordassistant" / "manifest.xml"
        if addins_dir.exists():
            addon_id = extract_addin_id(str(addins_dir))
        else:
            addon_id = _DEFAULT_ADDIN_ID

    key_path = get_wef_key_path()
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
        winreg.DeleteValue(key, addon_id)
        winreg.CloseKey(key)
        result["success"] = True
    except FileNotFoundError:
        result["success"] = True  # 原本就不存在，视为成功
    except PermissionError:
        result["error"] = "权限不足"
    except Exception as e:
        result["error"] = str(e)

    return result


def is_addin_registered(addon_id: Optional[str] = None) -> bool:
    """检查 Add-in 是否已注册"""
    if addon_id is None:
        addins_dir = get_addins_dir() / "wordassistant" / "manifest.xml"
        if addins_dir.exists():
            addon_id = extract_addin_id(str(addins_dir))
        else:
            addon_id = _DEFAULT_ADDIN_ID

    key_path = get_wef_key_path()
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, addon_id)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False
    except Exception:
        return False


def get_registered_addins() -> dict:
    """列出所有通过 WEF 注册的 Add-in"""
    key_path = get_wef_key_path()
    addins = {}
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        i = 0
        while True:
            try:
                name, data, _ = winreg.EnumValue(key, i)
                addins[name] = data
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    return addins
```

### 3.5 manifest.xml 路径维护

Word Add-in 注册后，manifest.xml **必须保持在原位置不变**（注册表存储的是路径指针）。如果文件被删除，Word 会报错。

**因此：**
- `~/.office_helper/addins/wordassistant/manifest.xml` 作为**永久存储路径**
- EXE 每次启动时检查文件是否存在，如果不存在则重新注册
- 用户卸载 EXE 时，应调用 `unregister_addin()`

### 3.6 Word 刷新机制

写入注册表后，Word 并不会立即感知。需要以下处理：

**方案 A（推荐）：静默通知**
Word 通过 `HKEY_CURRENT_USER\Software\Microsoft\Office\{ver}\WEF\{addon_id}` 下的 `RefreshAddins` 值来触发刷新。写一个额外的 DWORD 值：

```python
# 注册成功后，设置刷新标记
refresh_key = rf"HKEY_CURRENT_USER\Software\Microsoft\Office\{version}\WEF"
try:
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, refresh_key, 0, winreg.KEY_WRITE)
    winreg.SetValueEx(key, "RefreshAddins", 0, winreg.REG_DWORD, 1)
    winreg.CloseKey(key)
except Exception:
    pass  # 非必须，不影响主流程
```

**方案 B（兜底）：提示用户重启 Word**
显示友好的提示对话框：「请关闭所有 Word 窗口后重新打开 Word」——这是最可靠的方式，也是 `office-addin-dev-settings` 官方工具的默认行为。

---

## 四、EXE 打包方案

### 4.1 技术选型

| 工具 | 选型 | 原因 |
|---|---|---|
| **Python 打包** | **PyInstaller 6.x** | Windows EXE 打包事实标准，支持单文件模式，`--onefile` |
| **Node.js 打包** | **pkg** | 将 Node.js 项目打包为无依赖的可执行文件 |
| **React 构建** | `npm run build` | 构建生产版本（而非 dev server） |
| **包管理** | pnpm（与现有项目一致） | 构建 word-addin |

### 4.2 打包策略：分别打包 + 合并

由于 Python 和 Node.js 是两套独立的 runtime，无法用单一工具同时打包。方案采用**分别打包 + 合并目录**：

```
步骤 1：构建 React 生产版本
  cd word-addin && pnpm install && pnpm run build
  → 输出到 word-addin/dist/

步骤 2：打包 Python 后端为单文件 EXE
  pyinstaller --onefile --name OfficeHelperBackend ...
  → 输出到 dist/OfficeHelperBackend.exe

步骤 3：打包 Node.js dev server 为单文件 EXE
  pkg word-addin/package.json --targets node18-win-x64
  → 输出到 dist/wordaddin-server.exe

步骤 4：合并所有文件
  dist/
  ├── OfficeHelper.exe          # Python 后端（主入口）
  ├── wordaddin-server.exe     # Node.js 前端 server
  ├── word-addin/dist/         # React 构建产物
  ├── word-addin/wordassistant/manifest.xml  # manifest
  └── start.bat                # 一键启动脚本
```

### 4.3 PyInstaller 打包 Python 后端

**命令：**
```bash
pyinstaller `
    --onefile `
    --name OfficeHelperBackend `
    --add-data "core;core" `
    --add-data "pservice;pservice" `
    --add-data "skills;skills" `
    --add-data "word-addin/wordassistant/manifest.xml;wordassistant" `
    --hidden-import pywin32 `
    --hidden-import pywin32.system_info `
    --hidden-import win32api `
    --hidden-import win32con `
    --hidden-import win32timezone `
    --exclude-module tkinter `
    --exclude-module matplotlib `
    --exclude-module numpy `
    --console `
    pservice/main.py
```

**关键点：**
- `--onefile`：单文件输出
- `--add-data`：必须明确包含 `core/`, `pservice/`, `skills/` 三个目录
- `--hidden-import pywin32*`：PyInstaller 无法自动检测 COM 模块的隐式导入
- `--console`：保留控制台窗口（用于显示后端日志，便于调试）；发布时可改为 `--windowed`
- `--name OfficeHelperBackend`：区分主程序和 Node.js server

### 4.4 pkg 打包 Node.js dev server

**前提条件：**
`word-addin/package.json` 需要添加 `bin` 字段指向 dev server 入口脚本。

```json
{
  "name": "wordaddin-server",
  "version": "1.0.0",
  "bin": "scripts/server-entry.js",
  "pkg": {
    "scripts": [
      "scripts/**/*.js",
      "node_modules/webpack-dev-server/**/*",
      "node_modules/webpack/**/*",
      "node_modules/@pnp/**/*",
      "node_modules/office-addin-dev-settings/**/*"
    ],
    "assets": [
      "dist/**/*",
      "wordassistant/**/*",
      "assets/**/*",
      "node_modules/office-ui-fabric-core/**/*"
    ]
  }
}
```

**打包命令：**
```bash
cd word-addin
pkg package.json --targets node18-win-x64 --output ../dist/wordaddin-server.exe
```

> **注意**：`pkg` 打包的 Node.js EXE 会在运行时尝试写入临时文件，需要确保临时目录可写。可以通过 `--public` 标志或配置 `pkg` 的 `public` 字段来指定只读资源。

### 4.5 主入口 EXE（Python）设计

Python 后端 EXE 作为**唯一主入口**，负责：

1. **检测 Word 是否在运行**（提示用户关闭）
2. **启动 Node.js server**（subprocess，静默后台）
3. **启动 FastAPI 后端**（subprocess，静默后台）
4. **注册 Word Add-in**（调用 `addin_manager.py`）
5. **等待用户按 Ctrl+C 退出**
6. **清理：注销 Add-in，终止子进程**

```python
# ============================================================
# 文件：pservice/launcher.py
# 用途：OfficeHelper.exe 主入口（PyInstaller 打包目标）
# ============================================================

import sys
import os
import time
import signal
import subprocess
import threading
import webbrowser
import json
from pathlib import Path

# 将 EXE 所在目录加入 Python 路径
BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent.parent
sys.path.insert(0, str(BASE_DIR))

# 导入 add-in 管理器
from pservice.api.addin_manager import (
    register_addin, unregister_addin,
    is_addin_registered, get_addins_dir,
    _DEFAULT_ADDIN_ID
)


# ---------- 全局进程句柄 ----------
backend_process: subprocess.Popen | None = None
node_process: subprocess.Popen | None = None


def log(msg: str):
    """带时间戳的日志输出"""
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def is_word_running() -> bool:
    """检测 Word 是否正在运行"""
    import win32com.client
    try:
        word = win32com.client.GetActiveObject("Word.Application")
        return True
    except Exception:
        return False


def check_office_installed() -> bool:
    """检测 Office 是否已安装"""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            0, winreg.KEY_READ
        )
        winreg.CloseKey(key)
        return True
    except WindowsError:
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
                0, winreg.KEY_READ
            )
            winreg.CloseKey(key)
            return True
        except WindowsError:
            return False


def wait_for_word_closed():
    """循环检测直到 Word 关闭"""
    log("检测到 Word 正在运行，请先关闭所有 Word 窗口...")
    while True:
        if not is_word_running():
            break
        time.sleep(1)
    log("Word 已关闭，继续启动...")


def ensure_wordaddin_dist():
    """
    确保 word-addin 的生产构建存在。
    如果 dist/ 目录为空或不存在，运行构建。
    """
    dist_dir = BASE_DIR / "word-addin" / "dist"
    if not dist_dir.exists() or not list(dist_dir.glob("*")):
        log("检测到 word-addin 未构建，开始构建...")
        build_wordaddin()
    else:
        log(f"word-addin 构建已就绪: {dist_dir}")


def build_wordaddin():
    """构建 React 生产版本（npm run build）"""
    wordaddin_dir = BASE_DIR / "word-addin"
    try:
        result = subprocess.run(
            ["pnpm", "run", "build"],
            cwd=str(wordaddin_dir),
            capture_output=True, text=True, timeout=300
        )
        if result.returncode != 0:
            log(f"WARNING: word-addin 构建失败: {result.stderr.decode()}")
        else:
            log("word-addin 构建完成")
    except FileNotFoundError:
        log("WARNING: pnpm 未找到，尝试使用 npm...")
        try:
            subprocess.run(
                ["npm", "run", "build"],
                cwd=str(wordaddin_dir),
                capture_output=True, timeout=300
            )
        except Exception as e:
            log(f"WARNING: npm 构建也失败: {e}")


def start_node_server():
    """启动 Node.js dev server（监听 localhost:3000）"""
    global node_process
    log("启动 wordaddin-server.exe（端口 3000）...")

    server_exe = BASE_DIR / "wordaddin-server.exe"
    wordaddin_dir = BASE_DIR / "word-addin"

    if server_exe.exists():
        # 使用打包后的 Node.js server
        node_process = subprocess.Popen(
            [str(server_exe)],
            cwd=str(wordaddin_dir),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    else:
        # 开发模式 fallback：直接用 pnpm/npx
        log("wordaddin-server.exe 未找到，fallback 到 pnpm start...")
        node_process = subprocess.Popen(
            ["pnpm", "start"],
            cwd=str(wordaddin_dir),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

    # 等待端口就绪
    for _ in range(30):
        try:
            import socket
            s = socket.socket()
            s.settimeout(1)
            s.connect(("127.0.0.1", 3000))
            s.close()
            log("wordaddin-server 已就绪（端口 3000）")
            return
        except Exception:
            time.sleep(1)
    log("WARNING: wordaddin-server 启动超时")


def start_backend():
    """启动 FastAPI 后端（端口 8765）"""
    global backend_process
    log("启动 OfficeHelper 后端（端口 8765）...")

    backend_exe = BASE_DIR / "OfficeHelperBackend.exe"
    pservice_dir = BASE_DIR / "pservice"

    if backend_exe.exists():
        backend_process = subprocess.Popen(
            [str(backend_exe)],
            cwd=str(pservice_dir),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
    else:
        # 开发模式 fallback
        backend_process = subprocess.Popen(
            [sys.executable, "-m", "uvicorn", "pservice.api:app",
             "--host", "127.0.0.1", "--port", "8765"],
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

    # 等待端口就绪
    for _ in range(20):
        try:
            import socket
            s = socket.socket()
            s.settimeout(1)
            s.connect(("127.0.0.1", 8765))
            s.close()
            log("OfficeHelper 后端已就绪（端口 8765）")
            return
        except Exception:
            time.sleep(1)
    log("WARNING: 后端启动超时")


def register_word_addin():
    """注册 Word Add-in"""
    manifest_src = BASE_DIR / "word-addin" / "wordassistant" / "manifest.xml"

    # 确保构建已就绪（Add-in 的 SourceLocation 指向 localhost:3000）
    ensure_wordaddin_dist()

    if is_addin_registered():
        log("Word Add-in 已注册，跳过注册步骤")
        return

    if not manifest_src.exists():
        log(f"ERROR: manifest.xml 未找到: {manifest_src}")
        return

    result = register_addin(str(manifest_src))
    if result["success"]:
        log(f"Word Add-in 注册成功: {result['addon_id']}")
        log("请关闭所有 Word 窗口后重新打开 Word")
    else:
        log(f"Word Add-in 注册失败: {result['error']}")


def cleanup():
    """退出时清理子进程"""
    log("正在关闭服务...")
    if backend_process:
        backend_process.terminate()
    if node_process:
        node_process.terminate()
    log("已退出")


def signal_handler(sig, frame):
    cleanup()
    sys.exit(0)


def main():
    signal.signal(signal.SIGINT, signal_handler)

    print("=" * 50)
    print("  OfficeHelper - Word AI 助手")
    print("=" * 50)
    print()

    # 1. 环境检测
    if not check_office_installed():
        log("ERROR: 未检测到 Office 安装，请先安装 Microsoft Word")
        input("按回车键退出...")
        return

    # 2. 如果 Word 在运行，等待关闭
    if is_word_running():
        wait_for_word_closed()

    # 3. 注册 Word Add-in
    register_word_addin()

    # 4. 启动服务
    start_node_server()
    start_backend()

    # 5. 打开 Word（可选，帮助用户直接进入）
    print()
    print("  OfficeHelper 已启动！")
    print()
    print("  后端服务: http://127.0.0.1:8765")
    print("  前端服务: http://localhost:3000")
    print("  Word Add-in: 请在 Word 功能区「开始」选项卡中点击「AI 助手」")
    print()
    print("  按 Ctrl+C 或关闭此窗口停止服务")
    print("=" * 50)

    # 6. 等待用户中断
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        cleanup()


if __name__ == "__main__":
    main()
```

### 4.6 最终打包流程脚本

```bash
# ============================================================
# 文件：package.bat
# 用途：一键执行 OfficeHelper EXE 打包
# 前提：已安装 Python 3.10+、Node.js、pnpm、PyInstaller
# ============================================================

@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ========================================
echo  OfficeHelper EXE 打包脚本
echo ========================================
echo.

REM --- 前置检查 ---
echo [1/4] 检查依赖...
where python >nul 2>&1 || (echo ERROR: Python 未安装 && exit /b 1)
where node >nul 2>&1 || (echo ERROR: Node.js 未安装 && exit /b 1)
where pnpm >nul 2>&1 || (echo ERROR: pnpm 未安装 && exit /b 1)
where pyinstaller >nul 2>&1 || (echo ERROR: PyInstaller 未安装 && pip install pyinstaller && exit /b 1)
echo  依赖检查通过
echo.

REM --- 构建 React 前端 ---
echo [2/4] 构建 word-addin（React 生产版本）...
cd word-addin
call pnpm install
call pnpm run build
if errorlevel 1 (
    echo ERROR: word-addin 构建失败
    exit /b 1
)
cd ..
echo  word-addin 构建完成
echo.

REM --- 打包 Node.js server ---
echo [3/4] 打包 wordaddin-server.exe（pkg）...
cd word-addin
call npx pkg package.json --targets node18-win-x64 --output ../dist/wordaddin-server.exe
if errorlevel 1 (
    echo WARNING: pkg 打包失败，wordaddin-server 将使用 pnpm fallback
) else (
    echo  wordaddin-server.exe 打包完成
)
cd ..
echo.

REM --- 打包 Python 后端 ---
echo [4/4] 打包 OfficeHelperBackend.exe（PyInstaller）...
pyinstaller ^
    --onefile ^
    --name OfficeHelperBackend ^
    --add-data "core;core" ^
    --add-data "pservice;pservice" ^
    --add-data "skills;skills" ^
    --add-data "word-addin\wordassistant\manifest.xml;wordassistant" ^
    --hidden-import pywin32 ^
    --hidden-import pywin32.system_info ^
    --hidden-import win32api ^
    --hidden-import win32con ^
    --hidden-import win32timezone ^
    --exclude-module tkinter ^
    --exclude-module matplotlib ^
    --exclude-module numpy ^
    --console ^
    pservice\main.py

if errorlevel 1 (
    echo ERROR: PyInstaller 打包失败
    exit /b 1
)
echo  OfficeHelperBackend.exe 打包完成
echo.

REM --- 合并到 dist ---
echo [合并] 整理打包产物...
if not exist "dist" mkdir dist

REM 复制 Python EXE，重命名为 OfficeHelper.exe
copy /y "dist\OfficeHelperBackend.exe" "dist\OfficeHelper.exe" >nul

REM 复制 word-addin 构建产物
xcopy /y /e "word-addin\dist" "dist\word-addin\dist\" >nul
xcopy /y /e "word-addin\wordassistant" "dist\word-addin\wordassistant\" >nul
xcopy /y /e "word-addin\assets" "dist\word-addin\assets\" >nul

REM 复制 manifest.xml
xcopy /y "word-addin\wordassistant\manifest.xml" "dist\word-addin\wordassistant\" >nul

echo.
echo ========================================
echo  打包完成！产物目录：dist\
echo.
dir /b "dist\"
echo ========================================
echo.
echo  运行说明：
echo    1. 双击 OfficeHelper.exe 启动
echo    2. 首次启动会自动注册 Word Add-in
echo    3. 关闭所有 Word 窗口后重新打开
echo    4. 在「开始」选项卡点击「AI 助手」
echo ========================================
pause
```

---

## 五、manifest.xml 配置

### 5.1 当前 manifest（需保持不变）

```xml
<Id>45e162e8-0cf4-4b83-ac50-e6a0463719b1</Id>
<DisplayName DefaultValue="wordassistant"/>
<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
<TaskpaneId>ButtonId1</TaskpaneId>
```

**重要**：`SourceLocation` 必须使用 `https://localhost:3000`（而非 `http://localhost:3000`），因为 Office Add-in 要求所有源页面通过 HTTPS 提供。本地开发时，`localhost:3000` 使用自签名证书，Word Edge WebView 信任 `localhost`。

### 5.2 关键约束

| 配置项 | 要求 | 当前值 |
|---|---|---|
| `<Id>` | 全局唯一 GUID | `45e162e8-0cf4-4b83-ac50-e6a0463719b1` |
| `<SourceLocation>` | HTTPS URL，指向 taskpane.html | `https://localhost:3000/taskpane.html` |
| `<Icon>` | 必须指向可访问的 URL（dev 模式下用 localhost） | `https://localhost:3000/assets/icon-16.png` |
| `<Requirements>` | Office.js 最低版本 | 需在 manifest 中声明 |

---

## 六、启动流程时序图

```
用户双击 OfficeHelper.exe
    │
    ▼
[Python Launcher 启动]
    │
    ├─ 检测 Office 是否安装
    │   └─ 未安装 → 报错退出
    │
    ├─ 检测 Word 是否运行
    │   └─ 正在运行 → 阻塞等待
    │
    ├─ 构建/确认 word-addin dist/ 已就绪
    │   └─ dist/ 为空 → 自动运行 pnpm run build
    │
    ├─ 调用 addin_manager.register_addin()
    │   ├─ 解析 manifest.xml → addon_id
    │   ├─ 复制到 ~/.office_helper/addins/wordassistant/
    │   ├─ 写入 HKCU\Software\Microsoft\Office\{ver}\WEF\Developer
    │   └─ 提示重启 Word
    │
    ├─ 启动 wordaddin-server.exe（后台，静默）
    │   └─ webpack-dev-server 监听 localhost:3000
    │
    └─ 启动 OfficeHelperBackend.exe（后台，静默）
        └─ uvicorn 监听 127.0.0.1:8765

    ▼
用户看到启动提示（控制台 / 窗口）
    │
    ▼
用户打开 Word
    │
    ▼
Word 读取 HKCU\...\WEF\Developer 注册表
    │
    ▼
Word 加载 wordassistant Add-in
    │
    ▼
Word 功能区「开始」选项卡 → AI助手 → 对话/打开按钮
    │
    ▼
用户点击按钮
    │
    ▼
Word Edge WebView 加载 https://localhost:3000/taskpane.html
    │
    ▼
React App 初始化，调用 /api/chat 等接口
    │
    ▼
FastAPI 后端处理，调用 Word COM
```

---

## 七、文件清单

### 7.1 需新增的文件

| 文件路径 | 用途 | 说明 |
|---|---|---|
| `pservice/api/addin_manager.py` | Add-in 注册管理器 | winreg 写注册表的核心逻辑 |
| `pservice/launcher.py` | EXE 主入口 | 启动两个子服务，注册 Add-in |
| `package.bat` | 打包脚本 | 一键执行 PyInstaller + pkg |
| `package.md` | 本文档 | 完整方案设计文档 |

### 7.2 需修改的文件

| 文件路径 | 修改内容 | 说明 |
|---|---|---|
| `word-addin/package.json` | 添加 `bin` 字段和 `pkg` 配置 | 支持 pkg 打包 |
| `word-addin/scripts/server-entry.js` | 新建 server 入口脚本 | pkg 打包的入口点 |
| `pservice/main.py` | 移除 uvicorn 命令行参数 | 作为模块被 launcher 调用 |

### 7.3 现有文件（无需修改）

- `word-addin/wordassistant/manifest.xml` — 已就绪，无需修改
- `pservice/api/routes.py` — 无需修改
- `pservice/api/service.py` — 无需修改
- `core/` — 无需修改
- `skills/` — 无需修改
- `requirements.txt` — 无需修改（winreg 是 Python 内置模块）

---

## 八、已知限制与注意事项

1. **Word 版本限制**：本方案仅针对 Office LTSC 2024 批量授权版。其他版本（Office 365 个人版、Office 2019 等）可能有不同的 UI 表现，但注册表路径相同，方案同样有效。

2. **HTTPS 要求**：Word Add-in 要求所有页面通过 HTTPS 提供。`localhost:3000` 的 webpack-devServer 自带自签名 HTTPS 证书，Word Edge WebView 默认信任 `localhost`。

3. **manifest 路径不可移动**：注册表存储的是 manifest.xml 的绝对路径。文件不可删除或移动，否则 Word 启动时报错。

4. **Python COM 依赖**：后端 `word_connector.py` 使用 `pywin32`，仅在 Windows 下有效。不支持 macOS / Linux。

5. **卸载时清理**：用户卸载 EXE 时，应调用 `unregister_addin()` 清理注册表，并将 `~/.office_helper/addins/` 目录删除。

6. **多用户兼容性**：如果同一台机器有多个 Windows 用户账户，每个用户的 `HKCU` 注册表是独立的，无需管理员权限，互不影响。

---

## 九、FAQ

**Q：为什么不能直接用「上载我的加载项」按钮？**
A：Office LTSC 2024 批量授权版默认禁用了该 UI 入口。这是微软对永久授权版的策略限制，非用户操作问题。

**Q：为什么选择注册表而非命令行工具 `office-addin-dev-settings`？**
A：`office-addin-dev-settings` 需要 Node.js 环境。PyInstaller 打包的 EXE 无法在运行时调用外部 npm 命令，因此必须自己实现注册表写入逻辑。

**Q：manifest.xml 中的 `https://localhost:3000` 会不会有 HTTPS 证书问题？**
A：不会。`localhost` 在 Windows 受信任站点范围内，Word Edge WebView 不会拒绝自签名证书。

**Q：用户需要管理员权限吗？**
A：不需要。注册表路径使用 `HKEY_CURRENT_USER`（当前用户专属），不需要管理员权限。

**Q：如果用户换了 WiFi 网络，localhost:3000 还能访问吗？**
A：可以。`localhost` 是本机回环地址，与网络无关。

**Q：能否在 EXE 中直接嵌入 Node.js runtime？**
A：理论上可以（通过 pkg 的 `--compile` 或 `nexe`），但会增加包体积。当前方案使用 pkg 将 Node.js server 打包为独立 EXE，已经解决了 Node.js 依赖问题。
