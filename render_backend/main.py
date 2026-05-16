from __future__ import annotations

import os
from datetime import datetime, timezone
from typing import Any

from fastapi import FastAPI, Header, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, create_engine, func, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker


SERVICE_NAME = "mga-cloud-sync"


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def database_url() -> str:
    raw = os.getenv("DATABASE_URL", "").strip()
    if not raw:
        return "sqlite:///cloud_sync.db"
    if raw.startswith("postgres://"):
        return raw.replace("postgres://", "postgresql+psycopg://", 1)
    if raw.startswith("postgresql://"):
        return raw.replace("postgresql://", "postgresql+psycopg://", 1)
    return raw


class Base(DeclarativeBase):
    pass


class MobileCapture(Base):
    __tablename__ = "mobile_capture"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    mobile_id: Mapped[str] = mapped_column(String(140), unique=True, index=True)
    source_device: Mapped[str] = mapped_column(Text, default="")
    user_name: Mapped[str] = mapped_column(String(160), default="")
    equipment_code: Mapped[str] = mapped_column(String(120), default="", index=True)
    component_name: Mapped[str] = mapped_column(String(160), default="")
    work_date: Mapped[str] = mapped_column(String(20), default="", index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    received_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    desktop_imported_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    payload_json: Mapped[str] = mapped_column(Text, default="{}")

    photos: Mapped[list["MobilePhoto"]] = relationship(
        back_populates="capture",
        cascade="all, delete-orphan",
        lazy="selectin",
    )


class MobilePhoto(Base):
    __tablename__ = "mobile_photo"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    capture_id: Mapped[int] = mapped_column(ForeignKey("mobile_capture.id", ondelete="CASCADE"), index=True)
    file_name: Mapped[str] = mapped_column(String(260), default="")
    mime_type: Mapped[str] = mapped_column(String(120), default="image/jpeg")
    captured_at: Mapped[str] = mapped_column(String(40), default="")
    data_url: Mapped[str] = mapped_column(Text)

    capture: Mapped[MobileCapture] = relationship(back_populates="photos")


class CatalogSnapshot(Base):
    __tablename__ = "catalog_snapshot"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(80), unique=True, default="default")
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    payload_json: Mapped[str] = mapped_column(Text, default="{}")


engine = create_engine(database_url(), pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
Base.metadata.create_all(engine)

app = FastAPI(title="MGA Cloud Sync", version="1.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def json_dumps(value: Any) -> str:
    import json

    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def json_loads(value: str) -> Any:
    import json

    if not value:
        return {}
    return json.loads(value)


def require_api_key(x_mga_api_key: str | None = Header(default=None)) -> None:
    if os.getenv("MGA_REQUIRE_API_KEY", "").strip().lower() not in {"1", "true", "yes", "si"}:
        return
    expected = os.getenv("MGA_API_KEY", "").strip()
    if not expected:
        raise HTTPException(status_code=500, detail="MGA_API_KEY no configurada en Render.")
    if x_mga_api_key != expected:
        raise HTTPException(status_code=401, detail="API key invalida.")


def database_status() -> dict[str, Any]:
    return {
        "engine": engine.dialect.name,
        "persistent": bool(os.getenv("DATABASE_URL", "").strip()),
    }


def capture_counts(session: Session) -> dict[str, int]:
    total = session.scalar(select(func.count(MobileCapture.id))) or 0
    pending = session.scalar(
        select(func.count(MobileCapture.id)).where(MobileCapture.desktop_imported_at.is_(None))
    ) or 0
    imported = session.scalar(
        select(func.count(MobileCapture.id)).where(MobileCapture.desktop_imported_at.is_not(None))
    ) or 0
    photos = session.scalar(select(func.count(MobilePhoto.id))) or 0
    return {
        "total": int(total),
        "pending": int(pending),
        "imported": int(imported),
        "photos": int(photos),
    }


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "ok": True,
        "service": SERVICE_NAME,
        "version": app.version,
        "database": database_status(),
        "generated_at": utc_now().isoformat(timespec="seconds"),
    }


@app.get("/api/stats")
def cloud_stats(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        return {
            "ok": True,
            "service": SERVICE_NAME,
            "version": app.version,
            "database": database_status(),
            "captures": capture_counts(session),
            "generated_at": utc_now().isoformat(timespec="seconds"),
        }


@app.get("/api/catalog")
def get_catalog(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        snapshot = session.scalar(select(CatalogSnapshot).where(CatalogSnapshot.name == "default"))
        if snapshot is None:
            return {"ok": True, "source": "cloud-empty", "equipment": [], "generated_at": utc_now().isoformat(timespec="seconds")}
        payload = json_loads(snapshot.payload_json)
        if isinstance(payload, dict):
            payload.setdefault("ok", True)
            payload.setdefault("source", "cloud")
            payload["generated_at"] = snapshot.updated_at.isoformat(timespec="seconds")
            return payload
        return {"ok": True, "source": "cloud", "equipment": []}


@app.post("/api/catalog")
async def publish_catalog(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Catalogo invalido.")
    equipment = payload.get("equipment") or []
    if not isinstance(equipment, list):
        raise HTTPException(status_code=400, detail="El catalogo debe contener una lista de equipos.")
    payload["ok"] = True
    payload["source"] = "cloud"
    payload["updated_at"] = utc_now().isoformat(timespec="seconds")
    with SessionLocal() as session:
        snapshot = session.scalar(select(CatalogSnapshot).where(CatalogSnapshot.name == "default"))
        if snapshot is None:
            snapshot = CatalogSnapshot(name="default")
            session.add(snapshot)
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(payload)
        session.commit()
    return {"ok": True, "equipment": len(equipment)}


@app.post("/api/sync")
async def sync_mobile_records(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Paquete invalido.")
    records = payload.get("records") or []
    if not isinstance(records, list):
        raise HTTPException(status_code=400, detail="El paquete no contiene lista de capturas.")

    source_device = str(payload.get("device") or "")
    user_name = str(payload.get("user") or "")
    results: list[dict[str, Any]] = []
    created = 0
    skipped = 0
    errors = 0
    with SessionLocal() as session:
        for record in records:
            try:
                if not isinstance(record, dict):
                    raise ValueError("Captura invalida.")
                mobile_id = str(record.get("mobile_id") or "").strip()
                if not mobile_id:
                    raise ValueError("Captura sin mobile_id.")
                existing = session.scalar(select(MobileCapture).where(MobileCapture.mobile_id == mobile_id))
                if existing is not None:
                    skipped += 1
                    results.append(
                        {
                            "mobile_id": mobile_id,
                            "capture_id": existing.id,
                            "created": False,
                            "stored": True,
                            "desktop_imported": existing.desktop_imported_at is not None,
                        }
                    )
                    continue

                photos = record.get("photos") or []
                stored_record = dict(record)
                stored_record["photos"] = []
                capture = MobileCapture(
                    mobile_id=mobile_id,
                    source_device=source_device,
                    user_name=str(record.get("user_name") or user_name),
                    equipment_code=str(record.get("equipment_code") or ""),
                    component_name=str(record.get("component_name") or ""),
                    work_date=str(record.get("work_date") or ""),
                    payload_json=json_dumps(stored_record),
                )
                session.add(capture)
                session.flush()
                evidence_count = 0
                for photo in photos if isinstance(photos, list) else []:
                    if not isinstance(photo, dict):
                        continue
                    data_url = str(photo.get("data") or "")
                    if not data_url:
                        continue
                    session.add(
                        MobilePhoto(
                            capture_id=capture.id,
                            file_name=str(photo.get("name") or f"{mobile_id}.jpg")[:260],
                            mime_type=str(photo.get("mime_type") or "image/jpeg")[:120],
                            captured_at=str(photo.get("captured_at") or ""),
                            data_url=data_url,
                        )
                    )
                    evidence_count += 1
                created += 1
                results.append(
                    {
                        "mobile_id": mobile_id,
                        "capture_id": capture.id,
                        "created": True,
                        "stored": True,
                        "desktop_imported": False,
                        "evidence": evidence_count,
                    }
                )
            except Exception as exc:
                errors += 1
                results.append({"mobile_id": str(record.get("mobile_id") or "") if isinstance(record, dict) else "", "created": False, "error": str(exc)})
        session.commit()

        counts = capture_counts(session)

    return {
        "ok": errors == 0,
        "created": created,
        "skipped": skipped,
        "errors": errors,
        "stored": created + skipped,
        "captures": counts,
        "results": results,
    }


@app.get("/api/desktop/pending")
def desktop_pending(
    limit: int = Query(default=200, ge=1, le=1000),
    include_imported: bool = Query(default=False),
    _auth: str | None = Header(default=None, alias="X-MGA-API-Key"),
) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        query = select(MobileCapture).order_by(MobileCapture.id.asc()).limit(limit)
        if not include_imported:
            query = query.where(MobileCapture.desktop_imported_at.is_(None))
        rows = session.scalars(query).all()
        records = []
        for row in rows:
            payload = json_loads(row.payload_json)
            photos = [
                {
                    "name": photo.file_name,
                    "mime_type": photo.mime_type,
                    "captured_at": photo.captured_at,
                    "data": photo.data_url,
                }
                for photo in row.photos
            ]
            records.append(
                {
                    "id": row.id,
                    "mobile_id": row.mobile_id,
                    "source_device": row.source_device,
                    "user_name": row.user_name,
                    "payload": payload,
                    "photos": photos,
                    "received_at": row.received_at.isoformat(timespec="seconds"),
                }
            )
        counts = capture_counts(session)
        return {"ok": True, "records": records, "count": len(records), "captures": counts}


@app.post("/api/desktop/ack")
async def desktop_ack(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    ids = payload.get("ids") if isinstance(payload, dict) else []
    if not isinstance(ids, list):
        raise HTTPException(status_code=400, detail="ids invalido.")
    clean_ids = [int(item) for item in ids if str(item).isdigit()]
    if not clean_ids:
        return {"ok": True, "updated": 0}
    with SessionLocal() as session:
        rows = session.scalars(select(MobileCapture).where(MobileCapture.id.in_(clean_ids))).all()
        now = utc_now()
        for row in rows:
            row.desktop_imported_at = now
        session.commit()
        counts = capture_counts(session)
    return {"ok": True, "updated": len(clean_ids), "captures": counts}


@app.post("/api/mobile/status")
async def mobile_status(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    ids = payload.get("mobile_ids") if isinstance(payload, dict) else []
    if not isinstance(ids, list):
        raise HTTPException(status_code=400, detail="mobile_ids invalido.")
    clean_ids = [str(item).strip() for item in ids if str(item).strip()]
    if not clean_ids:
        return {"ok": True, "statuses": []}
    with SessionLocal() as session:
        rows = session.scalars(select(MobileCapture).where(MobileCapture.mobile_id.in_(clean_ids))).all()
        by_id = {row.mobile_id: row for row in rows}
        statuses = []
        for mobile_id in clean_ids:
            row = by_id.get(mobile_id)
            if row is None:
                statuses.append({"mobile_id": mobile_id, "uploaded": False, "desktop_imported": False})
                continue
            statuses.append(
                {
                    "mobile_id": mobile_id,
                    "uploaded": True,
                    "capture_id": row.id,
                    "desktop_imported": row.desktop_imported_at is not None,
                    "desktop_imported_at": row.desktop_imported_at.isoformat(timespec="seconds") if row.desktop_imported_at else "",
                }
            )
    return {"ok": True, "statuses": statuses}
