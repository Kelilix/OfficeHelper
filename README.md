# OfficeHelper 智能文档助手

基于 Microsoft Office Web Add-in + FastAPI 后端 + 大模型的 Word 文档 AI 辅助工具。

## 项目架构

```
OfficeHelper/
├── word-addin/              # Office Web 加载项 (React + Fluent UI)
│   └── wordassistant/       # 任务窗格前端，调用后端 API
├── pservice/                 # FastAPI 后端服务
│   └── api/                  # 路由：/api/chat、/api/word/status 等
├── core/                     # 核心模块（供 pservice 调用）
│   ├── word_connector.py     # pywin32 COM Word 操作
│   ├── llm_service.py         # OpenAI / Anthropic / Ollama 接口
│   └── settings.py            # 配置管理
├── skills/                   # AI Agent 技能定义 (SKILL.md)
└── utils/                    # 工具函数
```

## 环境要求

- Windows 10/11
- Microsoft Word
- Python 3.10+
- 大模型 API Key（OpenAI / Anthropic）或 Ollama 本地部署

## 快速启动

```bash
# 安装依赖
pip install -r requirements.txt

# 启动后端服务（端口 8765）
cd pservice
python -m uvicorn main:app --host 127.0.0.1 --port 8765

# Sideload Word 加载项
# 在 Word 中：文件 → 选项 → 加载项 → 管理 COM 加载项 → 浏览
# 选择 word-addin/wordassistant/dist/manifest.xml
```

## 配置

配置文件位于 `~/.office_helper/config.json`，或通过环境变量：

```bash
LLM_PROVIDER=openai       # openai / anthropic / ollama
LLM_MODEL=gpt-4
OPENAI_API_KEY=your_key
ANTHROPIC_API_KEY=your_key
DEBUG=true
```

## 注意事项

1. pywin32 需要正确注册 Windows COM 接口
2. 使用 Ollama 本地模型需先安装并启动 Ollama 服务
3. 后端服务必须在 Word 加载项之前启动
