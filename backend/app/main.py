from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routers import upload
from app.routers import convert
from app.routers import download
from app.routers import analyze

app = FastAPI(
    title="Free Document Converter",
    description="Convert, edit and download documents",
    version="1.0.0"
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Register routers
app.include_router(upload.router)
app.include_router(convert.router)
app.include_router(download.router)
app.include_router(analyze.router)


@app.get("/")
def root():
    return {
        "status": "success",
        "message": "Document Converter API is running"
    }


@app.get("/api/health")
def health():
    return {
        "status": "healthy"
    }