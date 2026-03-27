# OfficeHelper 智能文档助手

基于 Microsoft Office Web Add-in + FastAPI 后端 + 大模型的 Word 文档 AI 辅助工具。

## 项目架构

```
OfficeHelper/
├── word-addin/              # Office Web 加载项 (React + Fluent UI)
│   └── wordassistant/       # 任务窗格前端，调用后端 API
├── pservice/                 # FastAPI 后端服务
│   └── api/
│       ├── main.py           # Uvicorn 入口，FastAPI app 初始化
│       ├── routes.py         # 路由：/api/chat、/api/word/status 等
│       └── service.py        # 全局单例（WordConnector / LLMService）
├── core/                     # 核心模块（供 pservice 调用）
│   ├── word_connector.py     # pywin32 COM Word 操作
│   ├── llm_service.py         # 多 Provider 大模型接口（OpenAI / Anthropic / Ollama / Qwen）
│   ├── agent.py              # Agent 编排（任务分解 + 工具调用）
│   ├── settings.py           # 配置管理（读写 ~/.office_helper/config.json）
│   └── config.py             # 配置工具函数
├── skills/                   # AI Agent 技能定义 (SKILL.md)
└── utils/                    # 工具函数
```

### 核心设计原则

- **所有配置集中于 `config.json`**，不依赖环境变量、不依赖 dataclass 包装
- **配置路径**：`~/.office_helper/config.json`（即 `C:\Users\<用户名>\.office_helper\config.json`）
- **打包部署时**：只需修改 `config.json` 文件即可切换模型 / provider，无需改动代码
- **Debug 模式**：将 `config.json` 中 `debug` 设为 `true`，每次 LLM 调用前会在控制台打印完整 prompt

## 环境要求

- Windows 10/11
- Microsoft Word
- Python 3.10+

## 快速启动

```bash
# 安装依赖
pip install -r requirements.txt

# 首次运行会自动在 ~/.office_helper/ 生成 config.json，按需编辑配置

# 启动后端服务（端口 8765）
cd pservice
python -m uvicorn main:app --host 127.0.0.1 --port 8765

# Sideload Word 加载项
# 在 Word 中：文件 → 选项 → 加载项 → 管理 COM 加载项 → 浏览
# 选择 word-addin/wordassistant/dist/manifest.xml
```

## 配置（config.json）

配置文件路径：`~/.office_helper/config.json`（首次运行自动生成）

### 完整默认配置

```json
{
  "llm": {
    "provider": "openai",
    "api_key": "",
    "model": "gpt-4",
    "base_url": "https://api.openai.com/v1",
    "temperature": 0.7,
    "max_tokens": 2000
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

### 各 Provider 配置模板

#### OpenAI

```json
{
  "llm": {
    "provider": "openai",
    "api_key": "sk-xxxxxxxxxxxxxxxxxxxxxxxx",
    "model": "gpt-4o",
    "base_url": "https://api.openai.com/v1",
    "temperature": 0.7,
    "max_tokens": 2000
  }
}
```

#### Anthropic Claude

```json
{
  "llm": {
    "provider": "anthropic",
    "api_key": "sk-ant-xxxxxxxxxxxxxxxxxxxxxxxx",
    "model": "claude-3-5-sonnet-20241022",
    "temperature": 0.7,
    "max_tokens": 2000
  }
}
```

#### Ollama（本地模型）

```json
{
  "llm": {
    "provider": "ollama",
    "api_key": "",
    "model": "qwen2.5",
    "base_url": "http://localhost:11434",
    "temperature": 0.7,
    "max_tokens": 2000
  }
}
```

> 使用 Ollama 前需先安装并启动 Ollama 服务：`ollama serve`

#### 通义千问（DashScope）

```json
{
  "llm": {
    "provider": "qwen",
    "api_key": "你的 DashScope API Key",
    "model": "qwen-plus",
    "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "temperature": 0.7,
    "max_tokens": 2000
  }
}
```

> DashScope API Key 在[阿里云百炼平台](https://bailian.console.aliyun.com/)申请

### 支持的 Provider 列表

| provider 值 | 说明 | API Key 必需 | 模型示例 |
|---|---|---|---|
| `openai` | OpenAI 官方 API | 是 | gpt-4o, gpt-4, gpt-3.5-turbo |
| `anthropic` | Anthropic Claude | 是 | claude-3-5-sonnet, claude-3-opus |
| `ollama` | 本地 Ollama 服务 | 否 | qwen2.5, llama2, mistral |
| `qwen` | 通义千问 DashScope | 是 | qwen-plus, qwen-turbo, qwen-vl-plus |

### 切换 Provider 步骤

1. 修改 `config.json` 中的 `llm.provider`
2. 填写对应的 `api_key` / `model` / `base_url`
3. 重启后端服务即可生效

## Debug 调试

将 `config.json` 中 `debug` 设为 `true`：

```json
{
  "debug": true
}
```

开启后，每次调用大模型接口前会在控制台打印：

```
────────────────────────────────────────────────────────────
[DEBUG][chat_with_context] Provider: qwen  |  Model: qwen-plus  |  BaseURL: https://dashscope.aliyuncs.com/compatible-mode/v1
────────────────────────────────────────────────────────────
  [0] SYSTEM: 你是一个专业的Word文档格式调整助手...
  [1] USER: 请帮我分析这份文档的格式问题
────────────────────────────────────────────────────────────
```

## 注意事项

1. pywin32 需要正确注册 Windows COM 接口
2. 后端服务必须在 Word 加载项之前启动
3. `config.json` 文件修改后需要重启后端服务才能生效
