"""
FastAPI entry point for the Revenue Warriors backend.

Run locally:
    uvicorn backend.main:app --reload

from the repository root (not from inside backend/), so that the
`backend.*` imports resolve.
"""

import os
import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from backend.agents.rahul_domain.agents.revenue_warriors.routes import (
    router as revenue_warriors_router,
)

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)

app = FastAPI(
    title="Revenue Warriors Backend",
    description="Parses the Revenue Warriors workbook and dispatches partner reports.",
    version="1.0.0",
)

# The Next.js frontend calls this API from the browser, so it needs CORS.
# Override in production with a comma-separated CORS_ORIGINS.
origins = [
    o.strip()
    for o in os.environ.get(
        "CORS_ORIGINS", "http://localhost:3000,http://127.0.0.1:3000"
    ).split(",")
    if o.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(revenue_warriors_router)


@app.get("/health")
async def health():
    """Liveness probe - deliberately does not touch Graph."""
    return {"status": "ok"}


@app.get("/")
async def root():
    return {
        "service": "Revenue Warriors Backend",
        "docs": "/docs",
        "health": "/health",
    }
