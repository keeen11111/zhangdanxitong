"""FastAPI 主入口。

启动：
    uvicorn backend.main:app --reload --port 8000
"""
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from backend.config import settings
from backend.database import init_db
from backend.routers import auth, projects, payroll, modules, entities, experiences, pipeline


settings.validate()


@asynccontextmanager
async def lifespan(app: FastAPI):
    # 启动：建表
    init_db()
    yield


app = FastAPI(
    title="AutoPayroll-Pro API",
    version="1.0.0",
    description="多租户薪资处理 SaaS 后端",
    lifespan=lifespan,
)

# CORS：允许前端（Next.js 默认 3000 端口）跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=list(settings.cors_origins),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(auth.router)
app.include_router(projects.router)
app.include_router(payroll.router)
app.include_router(modules.router)
app.include_router(entities.router)
app.include_router(pipeline.router)
app.include_router(experiences.router)


@app.get("/", tags=["health"])
def health():
    return {"status": "ok", "service": "AutoPayroll-Pro API"}


@app.get("/api/health", tags=["health"])
def api_health():
    return {
        "status": "ok",
        "service": "AutoPayroll-Pro API",
        "environment": settings.environment,
    }
