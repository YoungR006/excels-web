from __future__ import annotations

from fastapi import APIRouter, HTTPException

from app.models.schemas import ParseRequest, ParseResponse
from app.services.zhipu import parse_to_operations


router = APIRouter()


@router.post("/parse", response_model=ParseResponse)
def parse_message(payload: ParseRequest):
    try:
        parsed = parse_to_operations(payload.message, payload.sheet)
    except RuntimeError as exc:
        if str(exc) == "missing_api_key":
            raise HTTPException(status_code=400, detail="missing_api_key")
        raise HTTPException(status_code=502, detail="llm_error")
    except Exception:
        raise HTTPException(status_code=400, detail="parse_failed")
    return parsed
