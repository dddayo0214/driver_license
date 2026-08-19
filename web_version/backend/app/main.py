from pathlib import Path

from fastapi import FastAPI, HTTPException, Response, status
from fastapi.middleware.cors import CORSMiddleware

from .config import LICENSE_TYPES, STATIONS
from .jobs import JobManager
from .schemas import JobStatus, RegistrationData
from .secure_store import EncryptedStore

BASE_DIR = Path(__file__).resolve().parents[1]
store = EncryptedStore(BASE_DIR / ".data")
jobs = JobManager()
app = FastAPI(title="駕照報名工具 API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://127.0.0.1:5173", "http://localhost:5173"],
    allow_credentials=False,
    allow_methods=["GET", "PUT", "POST"],
    allow_headers=["Content-Type"],
)


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.get("/api/options")
def options():
    return {"license_types": LICENSE_TYPES, "stations": STATIONS}


@app.get("/api/profile", response_model=RegistrationData | None)
def get_profile():
    try:
        return store.load()
    except ValueError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.put("/api/profile", status_code=status.HTTP_204_NO_CONTENT)
def save_profile(data: RegistrationData):
    store.save(data.model_dump(mode="json"))
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@app.get("/api/registration/status", response_model=JobStatus)
def registration_status():
    return jobs.status()


@app.post("/api/registration/start", response_model=JobStatus, status_code=status.HTTP_202_ACCEPTED)
def start_registration(data: RegistrationData):
    store.save(data.model_dump(mode="json"))
    if not jobs.start(data):
        raise HTTPException(status_code=409, detail="報名作業已在執行中")
    return jobs.status()


@app.post("/api/registration/stop", response_model=JobStatus)
def stop_registration():
    if not jobs.stop():
        raise HTTPException(status_code=409, detail="目前沒有執行中的報名作業")
    return jobs.status()
