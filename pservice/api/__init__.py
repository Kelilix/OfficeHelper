"""
FastAPI 应用模块
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    title="OfficeHelper API",
    description="AI Agent Word 助手后端服务",
    version="1.0.0",
)

# CORS：允许 Office.js (Edge WebView) 从 localhost:3000 访问
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://localhost:3000",
        "https://127.0.0.1:3000",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 路由由 main.py 中 app.include_router(...) 注册，避免重复挂载
