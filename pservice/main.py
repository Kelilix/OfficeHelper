"""
OfficeHelper Python 服务
FastAPI HTTP API，供 Word Web 加载项调用
"""

import sys
import os

# 将项目根目录加入路径，以便导入 core/
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from api import app
from api.routes import router as api_router
from api.service import word_service, llm_service

# 注册路由
app.include_router(api_router)

# 健康检查
@app.get("/")
def root():
    return {
        "name": "OfficeHelper Python Service",
        "version": "1.0.0",
        "word_connected": word_service.is_connected(),
    }

@app.get("/health")
def health():
    return {"status": "ok", "word": word_service.is_connected()}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=8765,
        reload=False,
        log_level="info",
    )
