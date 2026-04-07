# OfficeHelper Word Add-in 架构文档

## 1. 整体架构图

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           用户（Microsoft Word）                               │
│                                                                             │
│   ┌─────────────────────────────┐   ┌─────────────────────────────┐        │
│   │     Word 文档窗口            │   │    Office Add-in 任务窗格     │        │
│   │                             │   │                             │        │
│   │   [文档内容]                 │   │   ┌─────────────────────┐   │        │
│   │                             │   │   │  AI 智能文档助手     │   │        │
│   │                             │   │   │  ─────────────────   │   │        │
│   │                             │   │   │                     │   │        │
│   │                             │   │   │  [对话记录区域]       │   │        │
│   │                             │   │   │                     │   │        │
│   │                             │   │   │  ┌───────────────┐  │   │        │
│   │                             │   │   │  │输入框        ▼│  │   │        │
│   │                             │   │   │  └───────────────┘  │   │        │
│   └─────────────────────────────┘   └─────────────────────────────┘        │
│                    ▲                        │                               │
│                    │ Office.js              │                               │
└────────────────────┼────────────────────────┼───────────────────────────────┘
                     │ (同一进程内)           │ fetch("/api/...")
                     │                        ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                           用户 PC（同一台机器）                               │
│                                                                             │
│   ┌────────────────────────────────────────────────────────────────────┐    │
│   │                    前端：webpack dev server (localhost:3000)        │    │
│   │                                                                     │    │
│   │   React 18 + TypeScript + Fluent UI v9                              │    │
│   │   ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐            │    │
│   │   │ App.tsx  │→│AIChat.tsx│→│ api.ts   │→│Office.js │               │    │
│   │   └──────────┘  └──────────┘  └──────────┘  └──────────┘            │    │
│   │      布局组件       对话组件    HTTP 客户端    Word 选区 API          │    │
│   │                                                                     │    │
│   └────────────────────────────────────────────────────────────────────┘     │
│                                        │                                     │
│                                        │ webpack proxy (/api → :8765)        │
│                                        ▼                                     │
│   ┌────────────────────────────────────────────────────────────────────┐     │
│   │               后端：pservice (FastAPI, localhost:8765)              │    │
│   │                                                                     │    │
│   │   Python 3.10+  FastAPI + Uvicorn                                   │    │
│   │   ┌────────────────┐   ┌────────────────────────────────────┐       │    │
│   │   │ routes.py      │   │          service.py                │       │    │
│   │   │  /api/chat    │→  │  word_service: WordConnector        │       │    │
│   │   │  /api/word/*  │   │  llm_service:  LLMService           │       │    │
│   │   └────────────────┘   └────────────────────────────────────┘       │    │
│   │         ↑                      ↑              ↑                     │    │
│   │         └──────────────────────┴──────────────┘                     │    │
│   │                      core/ 模块                                     │    │
│   │   ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐            │    │
│   │   │word_     │  │llm_      │  │settings  │  │skills    │            │    │
│   │   │connector │  │service   │  │.py       │  │/__init__ │            │    │
│   │   └──────────┘  └──────────┘  └──────────┘  └──────────┘            │    │
│   │      COM接口      多Provider      配置加载    Skill描述文本           │    │
│   └────────────────────────────────────────────────────────────────────┘     │
│                                        │                                     │
│                          ┌─────────────┴─────────────┐                       │
│                          ▼                           ▼                       │
│   ┌───────────────────────────┐    ┌────────────────────────────────────┐    │
│   │  Word.Application (COM)   │    │  LLM Provider (HTTP/REST)          │    │
│   │                           │    │                                    │    │
│   │  通过 pywin32 操作 Word    │    │  OpenAI / Anthropic / Ollama       │    │
│   │  · 读写选区文本            │    │  (用户配置的 API Key)               │    │
│   │  · 设置字体/对齐/间距      │     │                                    │   │
│   │  · 页面设置                │    │                                    │   │
│   └───────────────────────────┘    └────────────────────────────────────┘   │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 2. 项目目录结构

```
OfficeHelper/
│
├── word-addin/                          # ★ Office Add-in 前端
│   └── wordassistant/
│       ├── package.json                  # npm 依赖 & 脚本
│       ├── webpack.config.js             # 构建配置（含 /api → :8765 代理）
│       ├── tsconfig.json                # TypeScript 配置
│       ├── manifest.xml                 # Office Add-in 清单（sideload 安装用）
│       ├── assets/                      # logo 等静态资源
│       │   └── logo-filled.png
│       ├── dist/                        # webpack 产物目录（npm run build 输出）
│       └── src/
│           ├── taskpane/
│           │   ├── index.tsx             # React 根入口，Office.onReady 启动
│           │   ├── taskpane.html         # HTML 壳，引入 office.js
│           │   ├── taskpane.ts           # 插入文本的辅助函数（Office.js）
│           │   ├── api.ts                # ★ 与 pservice 通信的 HTTP 客户端
│           │   └── components/
│           │       ├── App.tsx           # 根组件：布局 + Header
│           │       ├── AIChat.tsx         # ★ 核心：对话 UI（消息列表 + 输入框）
│           │       ├── Header.tsx         # (模板遗留)
│           │       ├── HeroList.tsx       # (模板遗留)
│           │       └── TextInsertion.tsx  # (模板遗留)
│           └── commands/                 # 功能区命令（模板遗留）
│               ├── commands.ts
│               └── commands.html
│
├── pservice/                            # ★ FastAPI 后端服务
│   ├── main.py                           # FastAPI 入口，注册路由，启动 uvicorn
│   └── api/
│       ├── __init__.py                  # FastAPI app 实例 + CORS 中间件
│       ├── routes.py                     # ★ 核心路由：/api/chat、/api/word/*
│       └── service.py                    # 全局单例：WordConnector + LLMService
│
├── core/                                # ★ Python 业务核心库
│   ├── __init__.py
│   ├── config.py                        # 配置导出
│   ├── settings.py                       # 配置管理（.env + JSON）
│   ├── word_connector.py                 # ★ Word COM 操作封装（pywin32）
│   ├── llm_service.py                   # ★ 多 LLM Provider 调用封装
│   ├── format_analyzer.py                # 格式问题分析
│   ├── format_fixer.py                  # 格式修复执行
│   └── screenshot_manager.py             # 截屏管理
│
├── skills/                              # ★ AI Agent Skill 定义
│   ├── __init__.py                       # Skill 加载器，get_skill_descriptions()
│   ├── word-font-operation/
│   │   └── SKILL.md                      # 字体操作规范文档
│   ├── word-paragraph-operation/
│   │   └── SKILL.md                      # 段落操作规范文档
│   └── word-page-operation/
│       └── SKILL.md                      # 页面操作规范文档
│
├── utils/                               # ★ 工具函数
│   ├── __init__.py
│   ├── response_parser.py               # LLM 响应解析（JSON → Action）
│   └── prompt_engineering.py            # 提示词模板
│
└── requirements.txt                     # Python 依赖
```

> **说明**：`core/`、`skills/`、`utils/` 三个目录共同构成 Python 后端的业务逻辑层，由 `pservice` 的 `service.py` 按需调用。这些模块**不依赖**任何 GUI 框架，纯粹是数据处理和服务层。

---

## 3. 技术选型

### 3.1 前端（Office Add-in）

| 层次 | 技术 | 说明 |
|------|------|------|
| 框架 | **React 18** | 业界标准，组件化，生态丰富 |
| 语言 | **TypeScript 5** | 类型安全，IDE 支持好 |
| UI 库 | **Fluent UI v9** (`@fluentui/react-components`) | Microsoft 官方，Office 设计语言一致 |
| 构建 | **Webpack 5** + `ts-loader` + `babel-loader` | 官方模板默认配置 |
| 样式 | **Fluent UI makeStyles** (CSS-in-JS) | 与组件库无缝集成 |
| 通信 | **fetch API**（通过 Webpack proxy） | 访问 pservice FastAPI |
| Office API | **Office.js** (`@types/office-js`) | 读取选区、文档信息 |

> **为何用 Webpack Proxy**：Office Add-in 运行在 Edge WebView 中，存在 CORS 限制。Webpack dev server 在 `webpack.config.js` 中将 `/api` 路径代理到 `http://127.0.0.1:8765`，绕过浏览器 CORS。生产环境可通过同域名部署解决。

### 3.2 后端（Python FastAPI）

| 层次 | 技术 | 说明 |
|------|------|------|
| Web 框架 | **FastAPI 0.109+** | 异步、高性能，自动 OpenAPI 文档 |
| ASGI 服务器 | **Uvicorn** | ASGI 标准实现 |
| 数据校验 | **Pydantic 2** | 请求/响应模型自动校验 |
| CORS | FastAPI `CORSMiddleware` | 允许 localhost:3000 访问 |
| 配置 | `python-dotenv` + 自定义 `Settings` | 环境变量 + JSON 持久化 |
| Word 操作 | **pywin32** (COM) | 仅 Windows，支持所有 Word 操作 |
| LLM 调用 | **openai Python SDK** / **anthropic SDK** / **requests** | 多 Provider |

### 3.3 大模型（LLM）

| Provider | 调用方式 | 说明 |
|----------|----------|------|
| **OpenAI** (GPT-4o / GPT-4) | `openai` Python SDK | 默认，支持 Vision |
| **Anthropic** (Claude 3.5 Sonnet) | `anthropic` Python SDK | 支持 Vision |
| **Ollama** (本地模型) | HTTP REST (`requests`) | 离线、低成本 |

---

## 4. 核心数据流

### 4.1 用户发送消息的完整流程

```
用户输入文字 → AIChat.tsx handleSend()
     │
     ▼
┌─────────────────────────────────────────────────────────┐
│ Office.js: getWordSelection() 读取选中文本               │
│ Office.js: getDocumentName()    读取文档名称              │
└─────────────────────────────────────────────────────────┘
     │
     ▼
fetch POST /api/chat
     │  (webpack proxy → http://127.0.0.1:8765/api/chat)
     ▼
┌─────────────────────────────────────────────────────────┐
│ FastAPI routes.py /api/chat                            │
│   1. 接收 {message, selection_text, document_name}     │
│   2. 构建 prompt（含 skill 描述）                        │
│   3. llm_service.chat_with_context(message, prompt)    │
│   4. 解析 LLM 返回的 JSON action 数组                   │
│   5. 遍历执行 actions → word_service                    │
│   6. 生成自然语言摘要                                    │
│   7. 返回 {response, success}                           │
└─────────────────────────────────────────────────────────┘
     │
     ▼
┌─────────────────────────────────────────────────────────┐
│ word_connector.py (COM 操作)                            │
│   · set_font()          · set_font_color()              │
│   · set_font_size()     · set_alignment()               │
│   · set_bold/italic/    · set_line_spacing()            │
│   · set_indent()        · set_paragraph_spacing()       │
│   · set_page_margins()  · set_paper_size()             │
│   · set_page_orientation()                              │
└─────────────────────────────────────────────────────────┘
     │
     ▼
返回 response → AIChat.tsx setMessages([...prev, assistantMsg])
     │
     ▼
UI 更新：气泡显示 AI 回复
```

### 4.2 Word 选区读取

有两种方式读取 Word 选区文本：

| 方式 | 代码位置 | 说明 |
|------|----------|------|
| **前端 Office.js** | `api.ts` → `getWordSelection()` | 通过 `Word.run()` 同步读取，最新 |
| **后端 COM** | `word_connector.py` → `get_selection_text()` | 通过 `pywin32` 读取 |

> 当前实现：**前端优先**。`AIChat.tsx` 在发送请求前主动读取选区传给后端。后端 `routes.py` 也会 fallback 尝试从 COM 读取。

### 4.3 LLM 响应解析

```
LLM 返回文本（混合自然语言 + JSON 数组）
     │
     ▼
routes.py _parse_actions()
     │  re.search(r'\[.*\]', ...)
     ▼
JSON 数组: [{"action": "set_font_size", "params": {...}}, ...]
     │
     ▼
routes.py _execute_action() 分发到 word_service
     │
     ▼
word_connector.py COM 操作
     │
     ▼
_summarize_execution() 提取自然语言摘要（去掉 JSON 部分）
     │
     ▼
返回给前端展示
```

---

## 5. 关键 API 路由

### 5.1 聊天接口（核心）

```
POST /api/chat
Content-Type: application/json

Request:
{
  "message": "把这段文字设为黑体三号",
  "selection_text": "要修改的文字内容",
  "document_name": "我的文档.docx"
}

Response:
{
  "response": "已将选中文本设置为黑体、三号字。",
  "success": true,
  "error": null
}
```

### 5.2 Word 状态查询

```
GET /api/word/status

Response:
{
  "connected": true,
  "document_name": "我的文档.docx",
  "has_selection": true,
  "selection_text": "要修改的文字内容"
}
```

### 5.3 Word 连接管理

```
POST /api/word/connect    # 启动 Word
POST /api/word/disconnect  # 关闭 Word
```

---

## 6. Word COM 操作能力详解

### 6.1 AI Agent 可调用的操作（已实现，共 12 个）

通过 `/api/chat` 接口，LLM 返回 JSON action 数组后自动分发执行。

#### 字体操作（6 个）

| Action | 参数 | 示例 |
|--------|------|------|
| `set_font` | `font_name` | `"font_name": "微软雅黑"` |
| `set_font_size` | `size`（磅值） | `"size": 12`（小四）、`"size": 14`（四号） |
| `set_bold` | `bold: true/false` | `"bold": true` |
| `set_italic` | `italic: true/false` | `"italic": true` |
| `set_underline` | `underline: true/false` | `"underline": true` |
| `set_font_color` | `color`（十六进制） | `"color": "FF0000"`（红色） |

#### 段落操作（4 个）

| Action | 参数 | 示例 |
|--------|------|------|
| `set_alignment` | `alignment` | `"alignment": "center"` / `"justify"` |
| `set_line_spacing` | `spacing`（倍数） | `"spacing": 1.5`（1.5倍） |
| `set_indent` | `first_line`/`left_indent`/`right_indent`（磅） | `"first_line": 21`（2字符） |
| `set_paragraph_spacing` | `space_before`/`space_after`（磅） | `"space_before": 12, "space_after": 6` |

#### 页面操作（3 个）

| Action | 参数 | 示例 |
|--------|------|------|
| `set_page_margins` | `top`/`bottom`/`left`/`right`（厘米） | 上下 2.54，左右 3.17 |
| `set_paper_size` | `paper_size` | `"paper_size": "A4"` / `"Letter"` |
| `set_page_orientation` | `orientation` | `"orientation": "landscape"` |

### 6.2 WordConnector 层已实现但未暴露给 AI 的操作

`word_connector.py` 中已有代码实现，但尚未在 `routes.py` 中注册为 AI 可调用 action：

#### 文档管理

| 方法 | 说明 |
|------|------|
| `open_document(file_path, read_only)` | 打开指定路径的文档 |
| `create_document()` | 新建空白文档 |
| `save_document(file_path)` | 保存文档（另存或覆盖） |
| `close_document(save_changes)` | 关闭当前文档 |
| `quit(save_changes)` | 退出 Word 进程 |

#### 选区操作

| 方法 | 说明 |
|------|------|
| `has_selection()` | 检查是否有文字被选中 |
| `get_selection_text()` | 读取当前选中的文字内容 |
| `get_selection_range()` | 获取选区的字符位置范围 `(start, end)` |
| `select_paragraph()` | 自动选中光标所在整个段落 |

#### 内容读取

| 方法 | 说明 |
|------|------|
| `get_text()` | 获取文档全部文本 |
| `get_paragraphs()` | 获取所有段落（含格式信息） |
| `get_styles()` | 获取文档中所有可用样式名 |

#### 内容写入

| 方法 | 说明 |
|------|------|
| `insert_text(text, at_position)` | 在指定位置插入文本 |
| `insert_table(rows, cols, at_position)` | 插入指定行列数的表格 |

#### 样式与页码

| 方法 | 说明 |
|------|------|
| `apply_style(style_name)` | 将样式（如"标题1"）应用于选区 |
| `add_page_number(position, format)` | 在页眉/页脚添加页码 |

#### 历史管理

| 方法 | 说明 |
|------|------|
| `undo()` / `redo()` | 撤销/重做（基于 UndoManager） |
| `can_undo()` / `can_redo()` | 检查是否可以撤销/重做 |

### 6.3 缺失操作（Word COM 支持但代码中未实现）

以下操作在 Word COM 中可实现，但当前代码未实现：

| 操作 | 说明 |
|------|------|
| 高亮文字 (`highlight`) | 文字背景色高亮 |
| 删除线 (`strikethrough`) | 删除线格式 |
| 上标/下标 (`superscript`/`subscript`) | 上标或下标文字 |
| 字符间距 (`character_spacing`) | 调整字符之间的间距 |
| 段落边框 (`paragraph_border`) | 段落四周边框线 |
| 项目符号/编号 (`bullet`/`numbering`) | 列表的项目符号或编号 |
| 分页符/分节符 | 手动插入分页或分节 |
| 页眉页脚内容 | 除页码外的页眉页脚文本内容 |
| 目录生成 (`generate_toc`) | 自动分析标题样式并生成目录 |
| 图片插入 (`insert_image`) | 在文档中插入图片 |
| 书签操作 (`bookmark`) | 添加、删除、跳转书签 |

> 如需扩展以上任何操作，可在 `word_connector.py` 中添加对应方法，并在 `routes.py` 的 `_execute_action()` 中注册新 action 类型。

### 6.4 技能文档（skills/）

| 目录 | 说明 |
|------|------|
| `word-font-operation/SKILL.md` | 字体操作规范，含 6 种操作的使用说明和 API 示例 |
| `word-paragraph-operation/SKILL.md` | 段落操作规范，含 5 种操作的使用说明 |
| `word-page-operation/SKILL.md` | 页面操作规范，含 3 种操作的使用说明 |

`skills/__init__.py` 中的 `get_skill_descriptions()` 会自动扫描这些目录，将 SKILL.md 内容拼接后作为上下文传给 LLM，帮助 AI 正确理解操作规范。

---

## 7. 部署与运行

### 7.1 开发环境

**终端 1 — 启动 Python 后端：**
```bash
cd OfficeHelper
pip install -r requirements.txt
python pservice/main.py
# → FastAPI 服务监听 http://127.0.0.1:8765
```

**终端 2 — 启动 Webpack Dev Server：**
```bash
cd word-addin/wordassistant
npm install
npm run dev-server
# → https://localhost:3000
```

**安装 Add-in（一次性）：**
```bash
npx office-addin-dev-certs install
npm run start
# → Word 侧边栏打开，点击"信任此加载项"
```

### 7.2 生产构建

```bash
# 构建前端
cd word-addin/wordassistant
npm run build  # 产物输出到 dist/

# 部署后端
# 将整个项目部署到服务器（Linux 可用，仅 COM 调用需 Windows）
# COM 操作层 pywin32 需在 Windows 环境运行
```

---

## 8. 与其他方案的对比

| 方案 | 状态 | 说明 |
|------|------|------|
| **VSTO Add-in**（.NET COM） | 已移除 | 老方案，需要安装，跨版本兼容性差 |
| **customTkinter 独立窗口** | 已移除 | Python GUI 桌面应用，独立于 Word 体验差 |
| **Office Add-in + FastAPI** | **当前方案** | Web 技术，跨平台（Office 365），UI 现代，部署灵活 |

当前方案的核心优势：
- Word 内嵌任务窗格，用户体验无缝
- React + Fluent UI，UI 与 Office 原生风格一致
- FastAPI 后端，语言无关，可替换为 Node/Go 等
- Office Add-in 支持 Windows/macOS/Web 多平台
