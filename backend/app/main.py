from fastapi import FastAPI, HTTPException

from app.schemas import HealthResponse, QuoteRequest, QuoteResponse
from app.services.engines import get_engine


app = FastAPI(
    title="Insurance Financing Quote API",
    version="0.1.0",
    description="PoC API for premium-financing insurance quote scenarios.",
)


@app.get("/health", response_model=HealthResponse)
def healthcheck():
    engine = get_engine()
    return HealthResponse(
        status="ok",
        workbookEngine=engine.name,
        supportedProducts=["WUS"],
    )


@app.post("/api/v1/quotes", response_model=QuoteResponse)
def create_quote(request: QuoteRequest):
    engine = get_engine()
    if not engine.supports(request):
        raise HTTPException(status_code=400, detail="Unsupported product or scenario for the current engine.")
    try:
        return engine.quote(request)
    except NotImplementedError as exc:
        raise HTTPException(status_code=501, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
