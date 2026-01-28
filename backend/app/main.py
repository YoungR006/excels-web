from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routes.sessions import router as sessions_router
from app.routes.nlp import router as nlp_router


app = FastAPI(title="Excels Web API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
def health():
    return {"ok": True}


app.include_router(sessions_router, prefix="/sessions", tags=["sessions"])
app.include_router(nlp_router, prefix="/nlp", tags=["nlp"])
