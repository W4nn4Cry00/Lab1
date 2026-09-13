"""
REST API Service Entrypoint
Clean Architecture, Secure Configuration
"""
from fastapi import FastAPI

app = FastAPI(
    title="REST API Service",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

@app.get("/health")
def health_check() -> dict[str, str]:
    return {"status": "ok", "environment": "production-ready"}
