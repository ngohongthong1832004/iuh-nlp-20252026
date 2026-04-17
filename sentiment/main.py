import logging

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

from predictor import load_artifacts, predict

logging.basicConfig(level=logging.INFO)

app = FastAPI(title="Tiki Sentiment API", version="1.0")


class PredictRequest(BaseModel):
    text: str


class PredictResponse(BaseModel):
    emotion: str
    confidence: float
    all_scores: dict


@app.on_event("startup")
async def startup():
    load_artifacts()


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/predict", response_model=PredictResponse)
def predict_sentiment(req: PredictRequest):
    if not req.text.strip():
        raise HTTPException(status_code=400, detail="Text cannot be empty")
    return predict(req.text)


@app.post("/predict_batch")
def predict_batch(texts: list[str]):
    if len(texts) > 100:
        raise HTTPException(status_code=400, detail="Max batch size is 100")
    return [predict(t) for t in texts]
