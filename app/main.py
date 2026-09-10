import json
import tempfile
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from app.api.auth import router as auth_router
from app.api.categories import router as categories_router
from app.api.consumptions import router as consumptions_router
from app.api.equipment import router as equipment_router
from app.api.mechanic_reports import router as mechanic_reports_router
from app.api.projects import router as projects_router
from app.api.reports import router as reports_router
from app.api.supervisors import router as supervisors_router
from app.api.tire_inspections import router as tire_inspections_router
from app.api.uploads import router as uploads_router
from app.api.users import router as users_router
from app.core.config import get_settings
from app.db.init_db import init_db
from app.services import file_store

settings = get_settings()
app = FastAPI(title=settings.app_name)

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOADS_DIR = Path(settings.upload_dir) if settings.upload_dir.strip() else Path(tempfile.gettempdir()) / "disponibilidad_uploads"
REPORTS_DIR = Path(settings.reports_dir) if settings.reports_dir.strip() else Path(tempfile.gettempdir()) / "disponibilidad_reports"
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

if settings.cors_origin_list:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origin_list,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

init_db()

if file_store.is_cloud():
    file_store.ensure_bucket()

@app.get("/health")
def health():
    return {"status": "healthy"}


@app.get("/uploads/{path:path}")
def get_upload(path: str):
    if file_store.is_cloud():
        return RedirectResponse(file_store.public_url(f"uploads/{path}"))
    target = (UPLOADS_DIR / path).resolve()
    if not str(target).startswith(str(UPLOADS_DIR.resolve())) or not target.is_file():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    return FileResponse(target)


@app.get("/repositorio/files/{name}")
def get_report_file(name: str):
    safe = Path(name).name
    if file_store.is_cloud():
        return RedirectResponse(file_store.public_url(f"reports/{safe}"))
    target = (REPORTS_DIR / safe).resolve()
    if not str(target).startswith(str(REPORTS_DIR.resolve())) or not target.is_file():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    return FileResponse(target)


app.mount("/repositorio", StaticFiles(directory=str(BASE_DIR / "static" / "repositorio"), html=True), name="repositorio")
app.mount("/admin", StaticFiles(directory=str(BASE_DIR / "static" / "admin"), html=True), name="admin")

PWA_DIR = BASE_DIR / "static" / "disponibilidad"


@app.get("/", response_class=HTMLResponse)
def pwa_index():
    target = PWA_DIR / "index.html"
    if not target.is_file():
        raise HTTPException(status_code=503, detail="PWA no disponible")
    return target.read_text(encoding="utf-8")


@app.get("/manifest.json")
def pwa_manifest():
    target = PWA_DIR / "manifest.json"
    if not target.is_file():
        raise HTTPException(status_code=404)
    return FileResponse(target, media_type="application/json")


@app.get("/sw.js")
def pwa_sw():
    target = PWA_DIR / "sw.js"
    if not target.is_file():
        raise HTTPException(status_code=404)
    return FileResponse(target, media_type="application/javascript")


@app.get("/icons/{name}")
def pwa_icon(name: str):
    safe = Path(name).name
    target = PWA_DIR / "icons" / safe
    if not target.is_file():
        raise HTTPException(status_code=404)
    if safe.endswith(".png"):
        return FileResponse(target, media_type="image/png")
    if safe.endswith(".svg"):
        return FileResponse(target, media_type="image/svg+xml")
    return FileResponse(target)

APK_DIR = BASE_DIR / "static" / "app"
APK_VERSION_PATH = Path(__file__).resolve().parent / "app_version.json"


@app.get("/api/v1/app-version", tags=["system"])
def app_version():
    if not APK_VERSION_PATH.is_file():
        raise HTTPException(status_code=404, detail="Versión no configurada")
    data = json.loads(APK_VERSION_PATH.read_text(encoding="utf-8"))
    data["download_url"] = "/api/v1/app-download"
    return data


@app.get("/api/v1/app-download", tags=["system"])
def app_download():
    target = APK_DIR / "MGA_Control_Equipos.apk"
    if not target.is_file():
        raise HTTPException(status_code=404, detail="APK no disponible")
    return FileResponse(target, media_type="application/vnd.android.package-archive", filename="MGA_Control_Equipos.apk")


app.include_router(auth_router, prefix="/api/v1")
app.include_router(categories_router, prefix="/api/v1")
app.include_router(projects_router, prefix="/api/v1")
app.include_router(equipment_router, prefix="/api/v1")
app.include_router(consumptions_router, prefix="/api/v1")
app.include_router(mechanic_reports_router, prefix="/api/v1")
app.include_router(tire_inspections_router, prefix="/api/v1")
app.include_router(uploads_router, prefix="/api/v1")
app.include_router(reports_router, prefix="/api/v1")
app.include_router(supervisors_router, prefix="/api/v1")
app.include_router(users_router, prefix="/api/v1")
