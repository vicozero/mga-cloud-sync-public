from __future__ import annotations

import base64
import json
import math
import os
import re
from copy import copy
import tempfile
import unicodedata
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import code128, qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas as pdf_canvas
from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, create_engine, delete, func, select, text as sql_text
from sqlalchemy import Float
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker

try:
    from .kpi_editable_excel import build_editable_kpi_excel
except ImportError:
    from kpi_editable_excel import build_editable_kpi_excel

try:
    import fitz
except Exception:
    fitz = None


SERVICE_NAME = "mga-cloud-sync"
EPP_REPLACEMENT_SOON_DAYS = 30


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
    __tablename__ = "mga_mobile_capture"

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
    __tablename__ = "mga_mobile_photo"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    capture_id: Mapped[int] = mapped_column(ForeignKey("mga_mobile_capture.id", ondelete="CASCADE"), index=True)
    file_name: Mapped[str] = mapped_column(String(260), default="")
    mime_type: Mapped[str] = mapped_column(String(120), default="image/jpeg")
    captured_at: Mapped[str] = mapped_column(String(40), default="")
    data_url: Mapped[str] = mapped_column(Text)

    capture: Mapped[MobileCapture] = relationship(back_populates="photos")


class CaptureDeletion(Base):
    __tablename__ = "mga_capture_deletion"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    mobile_id: Mapped[str] = mapped_column(String(140), default="", index=True)
    source_device: Mapped[str] = mapped_column(Text, default="")
    user_name: Mapped[str] = mapped_column(String(160), default="")
    equipment_code: Mapped[str] = mapped_column(String(120), default="", index=True)
    component_name: Mapped[str] = mapped_column(String(160), default="")
    work_date: Mapped[str] = mapped_column(String(20), default="", index=True)
    shift: Mapped[str] = mapped_column(String(80), default="")
    payload_json: Mapped[str] = mapped_column(Text, default="{}")
    deleted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class DesktopCaptureChange(Base):
    __tablename__ = "mga_desktop_capture_change"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    sync_id: Mapped[str] = mapped_column(String(80), index=True)
    logical_key: Mapped[str] = mapped_column(String(500), index=True)
    source_device: Mapped[str] = mapped_column(String(180), default="")
    source_updated_at: Mapped[str] = mapped_column(String(40), default="")
    deleted: Mapped[int] = mapped_column(Integer, default=0)
    payload_json: Mapped[str] = mapped_column(Text, default="{}")
    received_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class CatalogSnapshot(Base):
    __tablename__ = "mga_catalog_snapshot"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(80), unique=True, default="default")
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    payload_json: Mapped[str] = mapped_column(Text, default="{}")


class PortalSnapshot(Base):
    __tablename__ = "mga_portal_snapshot"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(80), unique=True, default="default")
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    payload_json: Mapped[str] = mapped_column(Text, default="{}")


class FilterInventoryItem(Base):
    __tablename__ = "mga_filter_inventory_item"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    part_key: Mapped[str] = mapped_column(String(180), unique=True, index=True)
    part_number: Mapped[str] = mapped_column(String(180), default="", index=True)
    description: Mapped[str] = mapped_column(Text, default="")
    quantity: Mapped[float] = mapped_column(Float, default=0)
    unit: Mapped[str] = mapped_column(String(40), default="PZA")
    min_stock: Mapped[float] = mapped_column(Float, default=0)
    location: Mapped[str] = mapped_column(String(180), default="")
    source_file: Mapped[str] = mapped_column(String(260), default="")
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)

    movements: Mapped[list["FilterInventoryMovement"]] = relationship(
        back_populates="item",
        cascade="all, delete-orphan",
        lazy="selectin",
    )


class FilterInventoryMovement(Base):
    __tablename__ = "mga_filter_inventory_movement"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    item_id: Mapped[int] = mapped_column(ForeignKey("mga_filter_inventory_item.id", ondelete="CASCADE"), index=True)
    movement_date: Mapped[str] = mapped_column(String(20), default="")
    movement_type: Mapped[str] = mapped_column(String(20), default="ENTRADA")
    quantity: Mapped[float] = mapped_column(Float, default=0)
    balance_after: Mapped[float] = mapped_column(Float, default=0)
    reference: Mapped[str] = mapped_column(String(180), default="")
    equipment_code: Mapped[str] = mapped_column(String(120), default="")
    service_interval: Mapped[str] = mapped_column(String(80), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    created_by: Mapped[str] = mapped_column(String(160), default="")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)

    item: Mapped[FilterInventoryItem] = relationship(back_populates="movements")


class EppItem(Base):
    __tablename__ = "mga_epp_item"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    code_key: Mapped[str] = mapped_column(String(180), unique=True, index=True)
    code: Mapped[str] = mapped_column(String(180), default="", index=True)
    description: Mapped[str] = mapped_column(Text, default="")
    category: Mapped[str] = mapped_column(String(160), default="")
    size: Mapped[str] = mapped_column(String(80), default="")
    unit: Mapped[str] = mapped_column(String(40), default="PZA")
    quantity: Mapped[float] = mapped_column(Float, default=0)
    min_stock: Mapped[float] = mapped_column(Float, default=0)
    useful_life_days: Mapped[int] = mapped_column(Integer, default=0)
    risk_area: Mapped[str] = mapped_column(String(180), default="")
    location: Mapped[str] = mapped_column(String(180), default="")
    training_required: Mapped[int] = mapped_column(Integer, default=0)
    maintenance_notes: Mapped[str] = mapped_column(Text, default="")
    source_file: Mapped[str] = mapped_column(String(260), default="")
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)

    movements: Mapped[list["EppMovement"]] = relationship(
        back_populates="item",
        cascade="all, delete-orphan",
        lazy="selectin",
    )
    deliveries: Mapped[list["EppDelivery"]] = relationship(
        back_populates="item",
        cascade="all, delete-orphan",
        lazy="selectin",
    )


class EppMovement(Base):
    __tablename__ = "mga_epp_movement"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    item_id: Mapped[int] = mapped_column(ForeignKey("mga_epp_item.id", ondelete="CASCADE"), index=True)
    movement_date: Mapped[str] = mapped_column(String(20), default="")
    movement_type: Mapped[str] = mapped_column(String(20), default="ENTRADA")
    quantity: Mapped[float] = mapped_column(Float, default=0)
    balance_after: Mapped[float] = mapped_column(Float, default=0)
    worker_name: Mapped[str] = mapped_column(String(180), default="")
    employee_id: Mapped[str] = mapped_column(String(80), default="")
    area: Mapped[str] = mapped_column(String(180), default="")
    reference: Mapped[str] = mapped_column(String(180), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)

    item: Mapped[EppItem] = relationship(back_populates="movements")


class EppDelivery(Base):
    __tablename__ = "mga_epp_delivery"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    item_id: Mapped[int] = mapped_column(ForeignKey("mga_epp_item.id", ondelete="CASCADE"), index=True)
    delivery_date: Mapped[str] = mapped_column(String(20), default="")
    worker_name: Mapped[str] = mapped_column(String(180), default="")
    employee_id: Mapped[str] = mapped_column(String(80), default="")
    area: Mapped[str] = mapped_column(String(180), default="")
    quantity: Mapped[float] = mapped_column(Float, default=0)
    useful_life_days: Mapped[int] = mapped_column(Integer, default=0)
    due_date: Mapped[str] = mapped_column(String(20), default="")
    received_by: Mapped[str] = mapped_column(String(180), default="")
    signature: Mapped[str] = mapped_column(String(180), default="")
    training_done: Mapped[int] = mapped_column(Integer, default=0)
    condition_status: Mapped[str] = mapped_column(String(80), default="ENTREGADO")
    notes: Mapped[str] = mapped_column(Text, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)

    item: Mapped[EppItem] = relationship(back_populates="deliveries")


class EppWorker(Base):
    __tablename__ = "mga_epp_worker"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    employee_id: Mapped[str] = mapped_column(String(80), default="", index=True)
    worker_name: Mapped[str] = mapped_column(String(180), default="", index=True)
    area: Mapped[str] = mapped_column(String(180), default="")
    position: Mapped[str] = mapped_column(String(180), default="")
    active: Mapped[int] = mapped_column(Integer, default=1)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class CloudRequisition(Base):
    __tablename__ = "mga_requisition"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    folio: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    request_date: Mapped[str] = mapped_column(String(20), default="")
    authorization_date: Mapped[str] = mapped_column(String(20), default="")
    equipment: Mapped[str] = mapped_column(String(180), default="")
    cost_center: Mapped[str] = mapped_column(String(180), default="")
    request_area: Mapped[str] = mapped_column(String(180), default="MTTO")
    location: Mapped[str] = mapped_column(String(180), default="PROVIDENCIA")
    requesting_unit: Mapped[str] = mapped_column(String(180), default="TALLER CENTRAL")
    operating_unit: Mapped[str] = mapped_column(String(180), default="PROVIDENCIA")
    priority: Mapped[str] = mapped_column(String(80), default="URGENTE")
    recommendation: Mapped[str] = mapped_column(String(80), default="ORIGINAL")
    status: Mapped[str] = mapped_column(String(80), default="Abierta")
    notes: Mapped[str] = mapped_column(Text, default="")
    purchase_status: Mapped[str] = mapped_column(String(120), default="")
    purchase_order: Mapped[str] = mapped_column(String(120), default="")
    purchase_order_date: Mapped[str] = mapped_column(String(20), default="")
    supplier: Mapped[str] = mapped_column(String(220), default="")
    buyer: Mapped[str] = mapped_column(String(180), default="")
    expected_date: Mapped[str] = mapped_column(String(20), default="")
    received_date: Mapped[str] = mapped_column(String(20), default="")
    tracking_notes: Mapped[str] = mapped_column(Text, default="")
    tracking_source_file: Mapped[str] = mapped_column(String(260), default="")
    tracking_updated_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)

    items: Mapped[list["CloudRequisitionItem"]] = relationship(
        back_populates="requisition",
        cascade="all, delete-orphan",
        lazy="selectin",
    )


class CloudRequisitionItem(Base):
    __tablename__ = "mga_requisition_item"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    requisition_id: Mapped[int] = mapped_column(ForeignKey("mga_requisition.id", ondelete="CASCADE"), index=True)
    quantity: Mapped[float] = mapped_column(Float, default=1)
    unit: Mapped[str] = mapped_column(String(40), default="PZA")
    part_number: Mapped[str] = mapped_column(String(180), default="")
    description: Mapped[str] = mapped_column(Text, default="")
    sort_order: Mapped[int] = mapped_column(Integer, default=0)
    active: Mapped[int] = mapped_column(Integer, default=1)

    requisition: Mapped[CloudRequisition] = relationship(back_populates="items")


class HoseChange(Base):
    __tablename__ = "mga_hose_change"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    change_date: Mapped[str] = mapped_column(String(20), default="", index=True)
    equipment: Mapped[str] = mapped_column(String(120), default="", index=True)
    system: Mapped[str] = mapped_column(String(120), default="")
    part_type: Mapped[str] = mapped_column(String(80), default="MANGUERA", index=True)
    diameter: Mapped[str] = mapped_column(String(80), default="")
    length_m: Mapped[float] = mapped_column(Float, default=0)
    quantity: Mapped[float] = mapped_column(Float, default=1)
    unit_cost: Mapped[float] = mapped_column(Float, default=0)
    estimated_life_days: Mapped[float] = mapped_column(Float, default=30)
    estimated_weekly_qty: Mapped[float] = mapped_column(Float, default=0)
    failure_reason: Mapped[str] = mapped_column(String(220), default="")
    technician: Mapped[str] = mapped_column(String(180), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    source: Mapped[str] = mapped_column(String(80), default="web")
    external_id: Mapped[str] = mapped_column(String(180), default="", index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class DieselRecord(Base):
    __tablename__ = "mga_diesel_record"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    work_date: Mapped[str] = mapped_column(String(20), default="", index=True)
    equipment: Mapped[str] = mapped_column(String(120), default="", index=True)
    condition: Mapped[str] = mapped_column(String(80), default="DISPONIBLE")
    shift: Mapped[str] = mapped_column(String(40), default="1")
    horometer_initial: Mapped[float] = mapped_column(Float, default=0)
    horometer_final: Mapped[float] = mapped_column(Float, default=0)
    worked_hours: Mapped[float] = mapped_column(Float, default=0)
    diesel_liters: Mapped[float] = mapped_column(Float, default=0)
    operator: Mapped[str] = mapped_column(String(180), default="")
    dispatcher: Mapped[str] = mapped_column(String(180), default="")
    supervisor: Mapped[str] = mapped_column(String(180), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    source: Mapped[str] = mapped_column(String(80), default="web")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class DieselDay(Base):
    __tablename__ = "mga_diesel_day"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    work_date: Mapped[str] = mapped_column(String(20), default="", unique=True, index=True)
    diesel_received: Mapped[float] = mapped_column(Float, default=0)
    initial_stock: Mapped[float] = mapped_column(Float, default=0)
    final_stock: Mapped[float] = mapped_column(Float, default=0)
    prosermin_stock: Mapped[float] = mapped_column(Float, default=0)
    supplier: Mapped[str] = mapped_column(String(180), default="")
    notes: Mapped[str] = mapped_column(Text, default="")
    source: Mapped[str] = mapped_column(String(80), default="web")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class DieselDeletedRecord(Base):
    __tablename__ = "mga_diesel_deleted_record"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    work_date: Mapped[str] = mapped_column(String(20), default="", index=True)
    equipment: Mapped[str] = mapped_column(String(120), default="", index=True)
    shift: Mapped[str] = mapped_column(String(40), default="1")
    deleted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


class DieselDeletedDay(Base):
    __tablename__ = "mga_diesel_deleted_day"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    work_date: Mapped[str] = mapped_column(String(20), default="", index=True)
    deleted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utc_now)


engine = create_engine(database_url(), pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
Base.metadata.create_all(engine)


def ensure_cloud_schema() -> None:
    requisition_columns = {
        "purchase_status": "VARCHAR(120) DEFAULT ''",
        "purchase_order": "VARCHAR(120) DEFAULT ''",
        "purchase_order_date": "VARCHAR(20) DEFAULT ''",
        "supplier": "VARCHAR(220) DEFAULT ''",
        "buyer": "VARCHAR(180) DEFAULT ''",
        "expected_date": "VARCHAR(20) DEFAULT ''",
        "received_date": "VARCHAR(20) DEFAULT ''",
        "tracking_notes": "TEXT DEFAULT ''",
        "tracking_source_file": "VARCHAR(260) DEFAULT ''",
        "tracking_updated_at": "TIMESTAMP",
    }
    hose_columns = {
        "external_id": "VARCHAR(180) DEFAULT ''",
    }
    if database_url().startswith("sqlite"):
        try:
            with engine.begin() as conn:
                columns = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(mga_diesel_day)")).fetchall()}
                if "prosermin_stock" not in columns:
                    conn.execute(sql_text("ALTER TABLE mga_diesel_day ADD COLUMN prosermin_stock FLOAT DEFAULT 0"))
                req_columns = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(mga_requisition)")).fetchall()}
                for column, definition in requisition_columns.items():
                    if column not in req_columns:
                        conn.execute(sql_text(f"ALTER TABLE mga_requisition ADD COLUMN {column} {definition}"))
                hose_existing = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(mga_hose_change)")).fetchall()}
                for column, definition in hose_columns.items():
                    if column not in hose_existing:
                        conn.execute(sql_text(f"ALTER TABLE mga_hose_change ADD COLUMN {column} {definition}"))
        except Exception:
            pass
        return
    try:
        with engine.begin() as conn:
            conn.execute(sql_text("ALTER TABLE mga_diesel_day ADD COLUMN IF NOT EXISTS prosermin_stock DOUBLE PRECISION DEFAULT 0"))
            for column, definition in requisition_columns.items():
                pg_definition = definition.replace("VARCHAR", "VARCHAR")
                conn.execute(sql_text(f"ALTER TABLE mga_requisition ADD COLUMN IF NOT EXISTS {column} {pg_definition}"))
            for column, definition in hose_columns.items():
                conn.execute(sql_text(f"ALTER TABLE mga_hose_change ADD COLUMN IF NOT EXISTS {column} {definition}"))
    except Exception:
        pass


ensure_cloud_schema()

app = FastAPI(title="MGA Cloud Sync", version="1.4.29")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)
EPP_MODULE_ENABLED = False


@app.middleware("http")
async def block_removed_epp_module(request: Request, call_next):
    if not EPP_MODULE_ENABLED and request.url.path.startswith("/api/epp"):
        return JSONResponse({"detail": "Modulo EPP almacen retirado."}, status_code=404)
    return await call_next(request)


STATIC_DIR = Path(__file__).resolve().parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
PRODUCT_CATALOG_PATH = STATIC_DIR / "productos_catalog.json"
REQUISITION_TEMPLATE_PATH = STATIC_DIR / "requisition_template.pdf"
REQUISITION_TRACKING_TEMPLATE_PATH = STATIC_DIR / "requisition_tracking_template.xlsx"
DIESEL_TEMPLATE_PATH = STATIC_DIR / "diesel_control_template.xlsx"
DIESEL_LOGO_PATH = STATIC_DIR / "mga-corner-logo.jfif"
MONTHLY_REPORT_TEMPLATE_PATH = STATIC_DIR / "monthly_report_template.pptx"
KPI_FORMAT_PDFS = {
    "barrenacion": STATIC_DIR / "kpi_barrenacion_format.pdf",
    "rezagado": STATIC_DIR / "kpi_rezagado_format.pdf",
    "aceites": STATIC_DIR / "kpi_aceites_format.pdf",
    "llantas": STATIC_DIR / "kpi_llantas_format.pdf",
}
KPI_FORMAT_FILENAMES = {
    "barrenacion": "Formato_KPI_Equipos_de_Barrenacion.pdf",
    "rezagado": "Formato_KPI_Equipos_de_Rezagado.pdf",
    "aceites": "Formato_KPI_KPI_Aceites.pdf",
    "llantas": "Formato_KPI_KPI_Llantas.pdf",
}
REQUISITION_UNITS = [
    "PZA", "JGO", "KIT", "SERV", "LT", "L", "GAL", "ML", "TAMBO", "TAMBOR",
    "CUBETA", "BOTE", "LATA", "CAJA", "PAQUETE", "BOLSA", "MTS", "M2", "M3",
    "KG", "GR", "TON", "ROLLO",
]


def json_dumps(value: Any) -> str:
    import json

    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def json_loads(value: str) -> Any:
    import json

    if not value:
        return {}
    try:
        return json.loads(value)
    except (TypeError, ValueError):
        return {}


def parse_float(value: Any, default: float = 0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def normalize_text(value: Any) -> str:
    return " ".join(str(value or "").strip().upper().split())


TIRE_REMAINING_WARNING_PERCENT = 35.0
TIRE_REMAINING_CRITICAL_PERCENT = 20.0
TIRE_REMAINING_WARNING_HOURS = 250.0


def normalize_capture_shift_py(value: Any) -> str:
    text = normalize_text(value)
    compact = text.replace(" ", "")
    if compact in {"1", "T1", "TURNO1", "PRIMERO", "DIA"}:
        return "Turno 1"
    if compact in {"2", "T2", "TURNO2", "SEGUNDO", "NOCHE"}:
        return "Turno 2"
    if text == "GENERAL":
        return "General"
    return "Turno 1"


def normalize_part_key(value: Any) -> str:
    text = normalize_text(value)
    return "".join(ch for ch in text if ch.isalnum())


REQUISITION_REFERENCE_RE = re.compile(r"\b(REQ|SER|SERV)\s*[\.\-_]?\s*0*(\d+)\b", re.IGNORECASE)


def product_search_text(item: dict[str, Any]) -> str:
    return normalize_text(f"{item.get('code') or item.get('clave') or ''} {item.get('product') or item.get('producto') or ''}")


def static_products() -> list[dict[str, str]]:
    if not PRODUCT_CATALOG_PATH.exists():
        return []
    try:
        payload = json.loads(PRODUCT_CATALOG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []
    rows = payload.get("products") if isinstance(payload, dict) else payload
    if not isinstance(rows, list):
        return []
    products: list[dict[str, str]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        code = str(row.get("code") or row.get("clave") or "").strip().upper()
        product = str(row.get("product") or row.get("producto") or "").strip().upper()
        if code and product:
            products.append({"code": code, "product": product})
    return products


def product_rows(session: Session | None = None, query: str = "", limit: int = 120) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    if session is not None:
        try:
            portal = latest_portal_payload(session)
            portal_products = portal.get("products") if isinstance(portal, dict) else []
            if isinstance(portal_products, list):
                for item in portal_products:
                    if not isinstance(item, dict):
                        continue
                    code = str(item.get("code") or item.get("clave") or "").strip().upper()
                    product = str(item.get("product") or item.get("producto") or "").strip().upper()
                    if code and product:
                        rows.append({"code": code, "product": product})
        except Exception:
            rows = []
    if not rows:
        rows = static_products()
    needle = normalize_text(query)
    if needle:
        rows = [row for row in rows if needle in product_search_text(row)]
    return rows[: max(1, min(int(limit or 120), 25000))]


def fmt_qty(value: Any) -> str:
    qty = parse_float(value, 0)
    return str(int(qty)) if float(qty).is_integer() else f"{qty:g}"


def requisition_field(row: Any, key: str) -> str:
    if isinstance(row, dict):
        return str(row.get(key) or "")
    return str(getattr(row, key, "") or "")


def requisition_item_field(item: Any, key: str) -> Any:
    if isinstance(item, dict):
        return item.get(key)
    return getattr(item, key, None)


def inventory_field_for(header: Any) -> str | None:
    text = normalize_part_key(header)
    mapping = {
        "NOPARTE": "part_number",
        "NUMPARTE": "part_number",
        "NUMEROPARTE": "part_number",
        "PARTNUMBER": "part_number",
        "CODIGO": "part_number",
        "SKU": "part_number",
        "DONALDSON": "part_number",
        "DESCRIPCION": "description",
        "DESCRIPTION": "description",
        "FILTRO": "description",
        "CANTIDAD": "quantity",
        "EXISTENCIA": "quantity",
        "EXISTENCIAS": "quantity",
        "STOCK": "quantity",
        "DISPONIBLE": "quantity",
        "SALDO": "quantity",
        "UNIDAD": "unit",
        "UDM": "unit",
        "MINIMO": "min_stock",
        "MINSTOCK": "min_stock",
        "STOCKMINIMO": "min_stock",
        "UBICACION": "location",
        "LOCALIZACION": "location",
        "RACK": "location",
    }
    return mapping.get(text)


def epp_field_for(header: Any) -> str | None:
    text = normalize_part_key(header)
    mapping = {
        "CODIGO": "code",
        "CLAVE": "code",
        "SKU": "code",
        "NOPARTE": "code",
        "NUMEROPARTE": "code",
        "PARTE": "code",
        "PRODUCTO": "code",
        "EPP": "code",
        "DESCRIPCION": "description",
        "DESCRIPCIONPRODUCTO": "description",
        "ARTICULO": "description",
        "ITEM": "description",
        "MATERIAL": "description",
        "CANTIDAD": "quantity",
        "EXISTENCIA": "quantity",
        "EXISTENCIAS": "quantity",
        "STOCK": "quantity",
        "SALDO": "quantity",
        "UNIDAD": "unit",
        "UM": "unit",
        "UDM": "unit",
        "MINIMO": "min_stock",
        "MINSTOCK": "min_stock",
        "STOCKMINIMO": "min_stock",
        "UBICACION": "location",
        "LOCALIZACION": "location",
        "RACK": "location",
        "CATEGORIA": "category",
        "TIPO": "category",
        "FAMILIA": "category",
        "TALLA": "size",
        "TAMANO": "size",
        "SIZE": "size",
        "VIDAUTIL": "useful_life_days",
        "VIDAUTILDIAS": "useful_life_days",
        "REPOSICIONDIAS": "useful_life_days",
        "AREA": "risk_area",
        "RIESGO": "risk_area",
        "AREARIESGO": "risk_area",
        "PUESTO": "risk_area",
        "NOTAS": "maintenance_notes",
        "OBSERVACIONES": "maintenance_notes",
        "MANTENIMIENTO": "maintenance_notes",
    }
    return mapping.get(text)


def requisition_tracking_field_for(header: Any) -> str | None:
    text = normalize_part_key(header)
    mapping = {
        "FOLIO": "folio",
        "REQUISICION": "folio",
        "REQUISICIONES": "folio",
        "REQUISICIÓN": "folio",
        "REQ": "folio",
        "NUMREQ": "folio",
        "NOREQ": "folio",
        "NOREQUISICION": "folio",
        "NOREQUISICIÓN": "folio",
        "NUMEROREQ": "folio",
        "NUMEROREQUISICION": "folio",
        "NUMEROREQUISICIÓN": "folio",
        "FOLIOREQ": "folio",
        "FOLIOREQUISICION": "folio",
        "NO": "folio",
        "DESCRIPCION": "description",
        "DESCRIPCIÃ“N": "description",
        "DESCRIPCIONREQ": "description",
        "CONCEPTO": "description",
        "NOECON": "equipment",
        "NOECONOMICO": "equipment",
        "NUMEROECONOMICO": "equipment",
        "EQUIPO": "equipment",
        "ESTATUS": "purchase_status",
        "ESTADO": "purchase_status",
        "STATUS": "purchase_status",
        "ESTATUSCOMPRAS": "purchase_status",
        "STATUSCOMPRAS": "purchase_status",
        "SEGUIMIENTO": "purchase_status",
        "ORDENCOMPRA": "purchase_order",
        "ORDENDECOMPRA": "purchase_order",
        "OC": "purchase_order",
        "OCSAP": "purchase_order",
        "PO": "purchase_order",
        "PONUMBER": "purchase_order",
        "PEDIDO": "purchase_order",
        "NUMOC": "purchase_order",
        "NOOC": "purchase_order",
        "FECHAOC": "purchase_order_date",
        "FECHAORDENCOMPRA": "purchase_order_date",
        "PROVEEDOR": "supplier",
        "PROVEEDORASIGNADO": "supplier",
        "SUPPLIER": "supplier",
        "COMPRADOR": "buyer",
        "BUYER": "buyer",
        "RESPONSABLE": "buyer",
        "TE": "expected_date",
        "TENTREGA": "expected_date",
        "TIEMPOENTREGA": "expected_date",
        "TIEMPODEENTREGA": "expected_date",
        "FECHAENTREGA": "expected_date",
        "FECHAPROMESA": "expected_date",
        "FECHAESTIMADA": "expected_date",
        "FECHAESTIMADAENTREGA": "expected_date",
        "PROMESA": "expected_date",
        "ETA": "expected_date",
        "ENTREGAESTIMADA": "expected_date",
        "FECHARECEPCION": "received_date",
        "FECHARECIBIDO": "received_date",
        "RECIBIDO": "received_date",
        "OBSERVACIONES": "tracking_notes",
        "OBS": "tracking_notes",
        "NOTAS": "tracking_notes",
        "COMENTARIOS": "tracking_notes",
        "COMENTARIOSCOMPRAS": "tracking_notes",
        "COMENTARIOSPROYECTO": "tracking_notes",
        "COMENTARIOSCOMPRASYOPROYECTO": "tracking_notes",
    }
    return mapping.get(text)


def excel_cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def split_requisition_reference(value: Any) -> tuple[str, str]:
    text = " ".join(str(value or "").strip().split())
    if not text:
        return "", ""
    match = REQUISITION_REFERENCE_RE.search(text)
    if not match:
        return normalize_requisition_folio(text), ""
    prefix = "SER" if match.group(1).upper().startswith("SER") else "REQ"
    folio = f"{prefix}-{int(match.group(2)):04d}"
    description = f"{text[:match.start()]} {text[match.end():]}".strip(" .-:/")
    return folio, normalize_text(description)


def normalize_requisition_folio(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    match = REQUISITION_REFERENCE_RE.search(text)
    if match:
        prefix = "SER" if match.group(1).startswith("SER") else "REQ"
        return f"{prefix}-{int(match.group(2)):04d}"
    text = text.replace(" ", "")
    return text


def requisition_lookup(session: Session, folio: str) -> CloudRequisition | None:
    normalized = normalize_requisition_folio(folio)
    if not normalized:
        return None
    row = session.scalar(select(CloudRequisition).where(CloudRequisition.folio == normalized))
    if row is not None:
        return row
    compact = normalize_part_key(normalized)
    rows = session.scalars(select(CloudRequisition)).all()
    return next((candidate for candidate in rows if normalize_part_key(candidate.folio) == compact), None)


def add_unique(values: list[str], value: Any) -> None:
    text = str(value or "").strip()
    if text and text not in values:
        values.append(text)


def inventory_item_payload(item: FilterInventoryItem, details: dict[str, Any] | None = None) -> dict[str, Any]:
    details = details or {}
    description = item.description or str(details.get("description") or "")
    return {
        "id": item.id,
        "part_key": item.part_key,
        "part_number": item.part_number,
        "description": description,
        "catalog_description": str(details.get("description") or ""),
        "equipment": str(details.get("equipment") or ""),
        "equipment_codes": str(details.get("equipment_codes") or ""),
        "item_type": str(details.get("item_type") or ""),
        "service_interval": str(details.get("service_interval") or ""),
        "quantity": item.quantity,
        "unit": item.unit,
        "min_stock": item.min_stock,
        "location": item.location,
        "source_file": item.source_file,
        "updated_at": item.updated_at.isoformat(timespec="seconds") if item.updated_at else "",
    }


def inventory_rows_payload(session: Session) -> list[dict[str, Any]]:
    rows = session.scalars(select(FilterInventoryItem).order_by(FilterInventoryItem.part_number.asc())).all()
    return [inventory_item_payload(row) for row in rows]


def upsert_inventory_item(
    session: Session,
    *,
    part_number: str,
    description: str = "",
    quantity: float = 0,
    unit: str = "PZA",
    min_stock: float = 0,
    location: str = "",
    source_file: str = "",
) -> FilterInventoryItem:
    part_number = normalize_text(part_number)
    part_key = normalize_part_key(part_number)
    if not part_key:
        raise HTTPException(status_code=400, detail="Numero de parte requerido.")
    item = session.scalar(select(FilterInventoryItem).where(FilterInventoryItem.part_key == part_key))
    if item is None:
        item = FilterInventoryItem(part_key=part_key, part_number=part_number)
        session.add(item)
    item.part_number = part_number
    if description:
        item.description = normalize_text(description)
    item.quantity = max(parse_float(quantity, 0), 0)
    item.unit = normalize_text(unit) or "PZA"
    item.min_stock = max(parse_float(min_stock, 0), 0)
    item.location = str(location or "").strip().upper()
    item.source_file = source_file
    item.updated_at = utc_now()
    return item


def epp_worker_payload(row: EppWorker) -> dict[str, Any]:
    return {
        "id": row.id,
        "employee_id": row.employee_id,
        "worker_name": row.worker_name,
        "area": row.area,
        "position": row.position,
        "active": int(row.active or 0),
        "created_at": row.created_at.isoformat(timespec="seconds") if row.created_at else "",
        "updated_at": row.updated_at.isoformat(timespec="seconds") if row.updated_at else "",
    }


def upsert_epp_worker(
    session: Session,
    *,
    worker_name: str,
    employee_id: str = "",
    area: str = "",
    position: str = "",
    active: int = 1,
) -> EppWorker:
    normalized_name = normalize_text(worker_name or employee_id)
    normalized_employee = normalize_text(employee_id)[:80]
    if not normalized_name:
        raise HTTPException(status_code=400, detail="Trabajador requerido.")
    if normalized_employee:
        row = session.scalar(select(EppWorker).where(EppWorker.employee_id == normalized_employee))
    else:
        row = session.scalar(select(EppWorker).where(EppWorker.employee_id == "", EppWorker.worker_name == normalized_name))
    if row is None:
        row = EppWorker(employee_id=normalized_employee, worker_name=normalized_name, created_at=utc_now())
        session.add(row)
    row.employee_id = normalized_employee
    row.worker_name = normalized_name[:180]
    row.area = normalize_text(area)[:180]
    row.position = normalize_text(position)[:180]
    row.active = 1 if active else 0
    row.updated_at = utc_now()
    return row


def ensure_epp_worker_from_values(session: Session, worker_name: str = "", employee_id: str = "", area: str = "") -> None:
    if not str(worker_name or "").strip() and not str(employee_id or "").strip():
        return
    try:
        upsert_epp_worker(session, worker_name=str(worker_name or employee_id), employee_id=str(employee_id or ""), area=str(area or ""))
    except Exception:
        pass


def epp_delivery_alerts_by_code(deliveries: list[EppDelivery], item_by_id: dict[int, EppItem]) -> dict[str, dict[str, Any]]:
    today = utc_now().date()
    soon_limit = today + timedelta(days=EPP_REPLACEMENT_SOON_DAYS)
    alerts: dict[str, dict[str, Any]] = {}
    for row in deliveries:
        if not row.due_date or row.condition_status in {"BAJA", "DEVUELTO"}:
            continue
        item = item_by_id.get(row.item_id)
        if item is None:
            continue
        try:
            due = datetime.fromisoformat(str(row.due_date)).date()
        except Exception:
            continue
        bucket = alerts.setdefault(
            item.code_key,
            {
                "overdue": 0,
                "due_soon": 0,
                "next_due_date": "",
                "next_worker": "",
                "next_employee_id": "",
                "next_area": "",
            },
        )
        if due < today:
            bucket["overdue"] += 1
        elif today <= due <= soon_limit:
            bucket["due_soon"] += 1
        if not bucket["next_due_date"] or due.isoformat() < bucket["next_due_date"]:
            bucket["next_due_date"] = due.isoformat()
            bucket["next_worker"] = row.worker_name
            bucket["next_employee_id"] = row.employee_id
            bucket["next_area"] = row.area
    return alerts


def epp_status(item: EppItem, alerts: dict[str, Any] | None = None) -> str:
    quantity = parse_float(item.quantity, 0)
    minimum = parse_float(item.min_stock, 0)
    alerts = alerts or {}
    if int(alerts.get("overdue", 0) or 0) > 0:
        return "VENCIDO POR VIDA UTIL"
    if int(alerts.get("due_soon", 0) or 0) > 0:
        return "PROXIMO A REPOSICION"
    if quantity <= 0:
        return "SIN STOCK"
    if minimum > 0 and quantity <= minimum:
        return "BAJO MINIMO"
    return "OK"


def epp_item_payload(item: EppItem, alerts: dict[str, Any] | None = None) -> dict[str, Any]:
    alerts = alerts or {}
    return {
        "id": item.id,
        "code_key": item.code_key,
        "code": item.code,
        "description": item.description,
        "category": item.category,
        "size": item.size,
        "unit": item.unit,
        "quantity": item.quantity,
        "min_stock": item.min_stock,
        "useful_life_days": item.useful_life_days,
        "risk_area": item.risk_area,
        "location": item.location,
        "training_required": int(item.training_required or 0),
        "maintenance_notes": item.maintenance_notes,
        "source_file": item.source_file,
        "updated_at": item.updated_at.isoformat(timespec="seconds") if item.updated_at else "",
        "status": epp_status(item, alerts),
        "next_due_date": alerts.get("next_due_date", ""),
        "next_worker": alerts.get("next_worker", ""),
        "overdue_deliveries": int(alerts.get("overdue", 0) or 0),
        "due_soon_deliveries": int(alerts.get("due_soon", 0) or 0),
        "qr_payload": f"EPP|{item.code}|{item.description}",
    }


def upsert_epp_item(
    session: Session,
    *,
    code: str,
    description: str = "",
    category: str = "",
    size: str = "",
    unit: str = "PZA",
    quantity: float = 0,
    min_stock: float = 0,
    useful_life_days: int = 0,
    risk_area: str = "",
    location: str = "",
    training_required: int = 0,
    maintenance_notes: str = "",
    source_file: str = "",
) -> EppItem:
    code = normalize_text(code)
    code_key = normalize_part_key(code)
    if not code_key:
        raise HTTPException(status_code=400, detail="Codigo EPP requerido.")
    item = session.scalar(select(EppItem).where(EppItem.code_key == code_key))
    if item is None:
        item = EppItem(code_key=code_key, code=code)
        session.add(item)
    item.code = code
    item.description = normalize_text(description) if description else item.description
    item.category = normalize_text(category)
    item.size = normalize_text(size)
    item.unit = normalize_text(unit) or "PZA"
    item.quantity = max(parse_float(quantity, 0), 0)
    item.min_stock = max(parse_float(min_stock, 0), 0)
    item.useful_life_days = max(int(parse_float(useful_life_days, 0)), 0)
    item.risk_area = normalize_text(risk_area)
    item.location = normalize_text(location)
    item.training_required = 1 if training_required else 0
    item.maintenance_notes = normalize_text(maintenance_notes)
    item.source_file = source_file
    item.updated_at = utc_now()
    return item


def epp_payload(session: Session) -> dict[str, Any]:
    portal = latest_portal_payload(session)
    portal_epp = portal.get("epp") if isinstance(portal, dict) else {}
    db_items = session.scalars(select(EppItem).order_by(EppItem.category.asc(), EppItem.description.asc(), EppItem.code.asc())).all()
    db_workers = session.scalars(select(EppWorker).order_by(EppWorker.worker_name.asc(), EppWorker.employee_id.asc())).all()
    if db_items or db_workers:
        item_by_id = {item.id: item for item in db_items}
        all_deliveries = session.scalars(select(EppDelivery).order_by(EppDelivery.id.desc())).all()
        alerts_by_code = epp_delivery_alerts_by_code(all_deliveries, item_by_id)
        items = [epp_item_payload(item, alerts_by_code.get(item.code_key, {})) for item in db_items]
        movements = session.scalars(select(EppMovement).order_by(EppMovement.id.desc()).limit(300)).all()
        deliveries = all_deliveries[:300]
        movement_rows = [
            {
                "id": row.id,
                "item_code": item_by_id.get(row.item_id).code if item_by_id.get(row.item_id) else "",
                "movement_date": row.movement_date,
                "movement_type": row.movement_type,
                "quantity": row.quantity,
                "balance_after": row.balance_after,
                "worker_name": row.worker_name,
                "employee_id": row.employee_id,
                "area": row.area,
                "reference": row.reference,
                "notes": row.notes,
                "created_at": row.created_at.isoformat(timespec="seconds") if row.created_at else "",
            }
            for row in movements
        ]
        delivery_rows = [
            {
                "id": row.id,
                "item_code": item_by_id.get(row.item_id).code if item_by_id.get(row.item_id) else "",
                "delivery_date": row.delivery_date,
                "worker_name": row.worker_name,
                "employee_id": row.employee_id,
                "area": row.area,
                "quantity": row.quantity,
                "useful_life_days": row.useful_life_days,
                "due_date": row.due_date,
                "received_by": row.received_by,
                "signature": row.signature,
                "training_done": int(row.training_done or 0),
                "condition_status": row.condition_status,
                "notes": row.notes,
                "created_at": row.created_at.isoformat(timespec="seconds") if row.created_at else "",
            }
            for row in deliveries
        ]
        workers = [epp_worker_payload(row) for row in db_workers]
    elif isinstance(portal_epp, dict):
        items = portal_epp.get("items") if isinstance(portal_epp.get("items"), list) else []
        movement_rows = portal_epp.get("movements") if isinstance(portal_epp.get("movements"), list) else []
        delivery_rows = portal_epp.get("deliveries") if isinstance(portal_epp.get("deliveries"), list) else []
        workers = portal_epp.get("workers") if isinstance(portal_epp.get("workers"), list) else []
    else:
        items, movement_rows, delivery_rows, workers = [], [], [], []
    summary = {
        "items": len(items),
        "total_quantity": sum(parse_float(row.get("quantity"), 0) for row in items if isinstance(row, dict)),
        "low_stock": sum(1 for row in items if isinstance(row, dict) and row.get("status") == "BAJO MINIMO"),
        "out_stock": sum(1 for row in items if isinstance(row, dict) and row.get("status") == "SIN STOCK"),
        "due_soon": sum(1 for row in items if isinstance(row, dict) and row.get("status") == "PROXIMO A REPOSICION"),
        "overdue": sum(1 for row in items if isinstance(row, dict) and row.get("status") == "VENCIDO POR VIDA UTIL"),
        "workers": len(workers),
        "deliveries": len(delivery_rows),
    }
    return {"ok": True, "items": items, "movements": movement_rows, "deliveries": delivery_rows, "workers": workers, "summary": summary}


def pdf_wrap_lines(c: pdf_canvas.Canvas, text: str, max_width: float, font_name: str = "Helvetica", font_size: float = 9) -> list[str]:
    words = str(text or "").replace("\n", " ").split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if c.stringWidth(candidate, font_name, font_size) <= max_width or not current:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def draw_pdf_qr(c: pdf_canvas.Canvas, payload: str, x: float, y: float, size: float) -> None:
    widget = qr.QrCodeWidget(payload or "EPP")
    bounds = widget.getBounds()
    width = bounds[2] - bounds[0]
    height = bounds[3] - bounds[1]
    drawing = Drawing(size, size, transform=[size / width, 0, 0, size / height, 0, 0])
    drawing.add(widget)
    renderPDF.draw(drawing, c, x, y)


def draw_pdf_code128(c: pdf_canvas.Canvas, payload: str, x: float, y: float, bar_width: float = 0.8, bar_height: float = 28) -> None:
    code = code128.Code128(payload or "EPP", barWidth=bar_width, barHeight=bar_height, humanReadable=True)
    code.drawOn(c, x, y)


def epp_qr_pdf_bytes(item: EppItem) -> bytes:
    stream = BytesIO()
    c = pdf_canvas.Canvas(stream, pagesize=letter)
    width, height = letter
    c.setFillColor(colors.HexColor("#071f49"))
    c.rect(0, height - 78, width, 78, fill=1, stroke=0)
    if DIESEL_LOGO_PATH.exists():
        c.drawImage(str(DIESEL_LOGO_PATH), 42, height - 58, width=92, height=36, preserveAspectRatio=True, mask="auto")
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(150, height - 45, "ETIQUETA EPP")
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(60, height - 130, item.code)
    c.setFont("Helvetica", 12)
    c.drawString(60, height - 152, item.description[:80])
    c.drawString(60, height - 172, f"Categoria: {item.category or 'S/D'}")
    payload = f"EPP|{item.code}|{item.description}"
    draw_pdf_qr(c, payload, 390, height - 250, 140)
    draw_pdf_code128(c, item.code, 60, height - 240, bar_width=0.9, bar_height=42)
    c.setFont("Helvetica", 8)
    c.setFillColor(colors.HexColor("#64748b"))
    c.drawString(60, 60, f"Generado: {utc_now().isoformat(timespec='seconds')}")
    c.showPage()
    c.save()
    return stream.getvalue()


def epp_delivery_pdf_bytes(delivery: EppDelivery, item: EppItem) -> bytes:
    stream = BytesIO()
    c = pdf_canvas.Canvas(stream, pagesize=letter)
    width, height = letter
    margin = 42
    c.setFillColor(colors.HexColor("#071f49"))
    c.rect(0, height - 82, width, 82, fill=1, stroke=0)
    if DIESEL_LOGO_PATH.exists():
        c.drawImage(str(DIESEL_LOGO_PATH), margin, height - 61, width=95, height=38, preserveAspectRatio=True, mask="auto")
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 34, "CONSTANCIA DE ENTREGA DE EPP")
    c.setFont("Helvetica", 9)
    c.drawRightString(width - margin, height - 58, f"Folio EPP-{delivery.id}")

    def field(label: str, value: str, x: float, y: float, w: float, h: float = 30) -> None:
        c.setFillColor(colors.HexColor("#f8fafc"))
        c.rect(x, y - h, w, h, fill=1, stroke=1)
        c.setFillColor(colors.HexColor("#64748b"))
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(x + 6, y - 11, label)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        c.drawString(x + 6, y - 24, str(value or "")[:58])

    top = height - 105
    field("Fecha de entrega", delivery.delivery_date, margin, top, 130)
    field("Fecha de reposicion", delivery.due_date or "S/D", margin + 140, top, 145)
    field("Trabajador", delivery.worker_name, margin, top - 40, 275)
    field("No. empleado", delivery.employee_id, margin + 285, top - 40, 120)
    field("Area / puesto", delivery.area, margin + 415, top - 40, 135)
    field("Codigo EPP", item.code, margin, top - 80, 130)
    field("Descripcion", item.description, margin + 140, top - 80, 285)
    field("Cantidad", f"{delivery.quantity:g} {item.unit or 'PZA'}", margin + 435, top - 80, 115)
    field("Categoria", item.category, margin, top - 120, 130)
    field("Talla", item.size, margin + 140, top - 120, 100)
    field("Vida util dias", str(delivery.useful_life_days or 0), margin + 250, top - 120, 100)
    field("Capacitacion", "SI" if delivery.training_done else "NO", margin + 360, top - 120, 90)
    field("Estado", delivery.condition_status, margin + 460, top - 120, 90)

    notes_top = top - 175
    c.setFillColor(colors.HexColor("#f8fafc"))
    c.rect(margin, notes_top - 72, width - (margin * 2), 72, fill=1, stroke=1)
    c.setFillColor(colors.HexColor("#64748b"))
    c.setFont("Helvetica-Bold", 8)
    c.drawString(margin + 6, notes_top - 13, "Notas y mantenimiento")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    notes = delivery.notes or item.maintenance_notes or ""
    for idx, line in enumerate(pdf_wrap_lines(c, notes, width - (margin * 2) - 16, "Helvetica", 9)[:4]):
        c.drawString(margin + 6, notes_top - 30 - (idx * 12), line)

    qr_y = notes_top - 190
    payload = f"EPP-ENTREGA|{delivery.id}|{item.code}|{delivery.employee_id}|{delivery.delivery_date}"
    draw_pdf_qr(c, payload, width - margin - 112, qr_y, 104)
    draw_pdf_code128(c, f"EPP-{delivery.id}", margin, qr_y + 18, bar_width=0.8, bar_height=34)
    c.setFont("Helvetica", 9)
    legal = "Recibi el equipo de proteccion personal descrito y me comprometo a usarlo, conservarlo y reportar cualquier dano, perdida o desgaste."
    for idx, line in enumerate(pdf_wrap_lines(c, legal, width - (margin * 2) - 135, "Helvetica", 9)[:3]):
        c.drawString(margin, qr_y - 20 - (idx * 12), line)

    sig_y = 145
    c.line(margin, sig_y, margin + 210, sig_y)
    c.line(width - margin - 210, sig_y, width - margin, sig_y)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(margin + 105, sig_y - 14, "Firma trabajador")
    c.drawCentredString(width - margin - 105, sig_y - 14, "Entrega / Almacen")
    c.setFont("Helvetica", 8)
    c.drawCentredString(margin + 105, sig_y - 28, delivery.signature or delivery.worker_name)
    c.drawCentredString(width - margin - 105, sig_y - 28, delivery.received_by or "")
    c.setFillColor(colors.HexColor("#64748b"))
    c.drawString(margin, 48, f"Generado: {utc_now().isoformat(timespec='seconds')}")
    c.showPage()
    c.save()
    return stream.getvalue()


def latest_catalog_payload(session: Session) -> dict[str, Any]:
    snapshot = session.scalar(select(CatalogSnapshot).where(CatalogSnapshot.name == "default"))
    if snapshot is None:
        return {"ok": True, "source": "cloud-empty", "equipment": []}
    payload = json_loads(snapshot.payload_json)
    if not isinstance(payload, dict):
        return {"ok": True, "source": "cloud", "equipment": []}
    payload.setdefault("ok", True)
    payload.setdefault("source", "cloud")
    payload.setdefault("equipment", [])
    return payload


def mobile_capture_portal_row(row: MobileCapture) -> dict[str, Any]:
    payload = json_loads(row.payload_json)
    if not isinstance(payload, dict):
        payload = {}
    return {
        "id": row.id,
        "mobile_id": row.mobile_id,
        "source": row.source_device or str(payload.get("source") or ""),
        "received_at": row.received_at.isoformat(timespec="seconds") if row.received_at else "",
        "work_date": row.work_date or str(payload.get("work_date") or ""),
        "shift": normalize_capture_shift_py(payload.get("shift") or payload.get("turno")),
        "equipment_code": row.equipment_code or str(payload.get("equipment_code") or payload.get("equipment") or ""),
        "equipment_description": "",
        "component": row.component_name or str(payload.get("component_name") or payload.get("component") or ""),
        "hi": parse_float(payload.get("hi"), 0),
        "hf": parse_float(payload.get("hf"), 0),
        "worked_hours": parse_float(payload.get("worked_hours"), 0),
        "mp_hours": parse_float(payload.get("mp_hours"), 0),
        "mc_hours": parse_float(payload.get("mc_hours"), 0),
        "standby_hours": parse_float(payload.get("standby_hours"), 0),
        "stops": int(parse_float(payload.get("stops"), 0)),
        "oil_liters": parse_float(payload.get("oil_liters"), 0),
        "oil_motor_15w40": parse_float(payload.get("oil_motor_15w40"), 0),
        "oil_hco_iso68": parse_float(payload.get("oil_hco_iso68"), 0),
        "oil_trans_sae30": parse_float(payload.get("oil_trans_sae30"), 0),
        "oil_sae50": parse_float(payload.get("oil_sae50"), 0),
        "oil_85w140": parse_float(payload.get("oil_85w140"), 0),
        "almo_liters": parse_float(payload.get("almo_liters"), 0),
        "coolant_liters": parse_float(payload.get("coolant_liters"), 0),
        "oil_hyd_vg100": parse_float(payload.get("oil_hyd_vg100"), 0),
        "atf_liters": parse_float(payload.get("atf_liters"), 0),
        "fault": str(payload.get("fault") or ""),
        "wear": str(payload.get("wear") or ""),
        "status": str(payload.get("status") or "Disponible"),
        "observations": str(payload.get("observations") or payload.get("details") or ""),
        "evidence_count": len(row.photos or []),
    }


def capture_merge_key(row: dict[str, Any]) -> tuple[str, str, str, str]:
    equipment_code = str(row.get("equipment_code") or row.get("equipment") or "").strip().upper()
    component = str(row.get("component") or row.get("component_name") or "").strip().upper()
    if not component:
        component = kpi_required_component_py(equipment_code)
    return (
        str(row.get("work_date") or "").strip(),
        normalize_capture_shift_py(row.get("shift")),
        equipment_code,
        component,
    )


def capture_delete_key(payload: dict[str, Any]) -> tuple[str, str, str, str]:
    return capture_merge_key(
        {
            "work_date": payload.get("work_date"),
            "shift": payload.get("shift"),
            "equipment_code": payload.get("equipment_code") or payload.get("equipment"),
            "component": payload.get("component") or payload.get("component_name"),
        }
    )


def mobile_capture_by_merge_key(session: Session, record: dict[str, Any]) -> MobileCapture | None:
    key = capture_merge_key(record)
    if not any(key):
        return None
    candidates = session.scalars(
        select(MobileCapture)
        .where(MobileCapture.work_date == key[0])
        .where(MobileCapture.equipment_code == key[2])
        .order_by(MobileCapture.id.desc())
    ).all()
    for row in candidates:
        if capture_merge_key(mobile_capture_portal_row(row)) == key:
            return row
    return None


def remove_capture_rows_from_portal_payload(payload: dict[str, Any], delete_key: tuple[str, str, str, str], mobile_id: str, capture_id: int) -> int:
    captures = payload.get("captures")
    if not isinstance(captures, list):
        return 0
    kept: list[Any] = []
    removed = 0
    for row in captures:
        if not isinstance(row, dict):
            kept.append(row)
            continue
        row_mobile_id = str(row.get("mobile_id") or "").strip()
        row_id = int(parse_float(row.get("id"), 0) or 0)
        mobile_match = bool(mobile_id and row_mobile_id == mobile_id)
        id_match = bool(capture_id and row_id == capture_id)
        key_match = bool(any(delete_key) and capture_merge_key(row) == delete_key)
        if mobile_match or id_match or key_match:
            removed += 1
            continue
        kept.append(row)
    if removed:
        payload["captures"] = kept
    return removed


def capture_deletion_payload(row: CaptureDeletion) -> dict[str, Any]:
    payload = json_loads(row.payload_json)
    if not isinstance(payload, dict):
        payload = {}
    payload.setdefault("mobile_id", row.mobile_id)
    payload.setdefault("work_date", row.work_date)
    payload.setdefault("shift", row.shift)
    payload.setdefault("equipment_code", row.equipment_code)
    payload.setdefault("component_name", row.component_name)
    payload.setdefault("component", row.component_name)
    return payload


def mobile_capture_portal_rows(session: Session, limit: int = 1000, start: str = "", end: str = "") -> list[dict[str, Any]]:
    query = select(MobileCapture).order_by(MobileCapture.work_date.desc(), MobileCapture.id.desc())
    if start:
        query = query.where(MobileCapture.work_date >= start)
    if end:
        query = query.where(MobileCapture.work_date <= end)
    if limit:
        query = query.limit(limit)
    rows = session.scalars(query).all()
    return [mobile_capture_portal_row(row) for row in rows]


def parse_iso_datetime(value: Any) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def portal_capture_is_closed_month(row: dict[str, Any]) -> bool:
    text = str(row.get("work_date") or "").strip()
    if not text:
        return False
    try:
        day = date.fromisoformat(text[:10])
    except ValueError:
        return False
    current_month_start = utc_now().date().replace(day=1)
    return day < current_month_start


def closed_capture_period_error(row: dict[str, Any]) -> str:
    text = str(row.get("work_date") or "").strip()
    if not portal_capture_is_closed_month(row):
        return ""
    return (
        f"El periodo {text[:7]} ya esta cerrado. "
        "No se puede modificar, eliminar ni importar capturas de meses anteriores."
    )


def replace_mobile_photos(session: Session, capture_id: int, mobile_id: str, photos: Any) -> int:
    if not isinstance(photos, list) or not photos:
        return 0
    session.execute(delete(MobilePhoto).where(MobilePhoto.capture_id == capture_id))
    evidence_count = 0
    for photo in photos:
        if not isinstance(photo, dict):
            continue
        data_url = str(photo.get("data") or "")
        if not data_url:
            continue
        session.add(
            MobilePhoto(
                capture_id=capture_id,
                file_name=str(photo.get("name") or f"{mobile_id}.jpg")[:260],
                mime_type=str(photo.get("mime_type") or "image/jpeg")[:120],
                captured_at=str(photo.get("captured_at") or ""),
                data_url=data_url,
            )
        )
        evidence_count += 1
    return evidence_count


def merge_mobile_captures_into_portal(session: Session, portal: dict[str, Any]) -> dict[str, Any]:
    captures = portal.get("captures")
    if not isinstance(captures, list):
        captures = []
    merged = [row for row in captures if isinstance(row, dict)]
    index_by_key = {capture_merge_key(row): idx for idx, row in enumerate(merged) if any(capture_merge_key(row))}
    portal_updated = parse_iso_datetime(portal.get("updated_at"))

    equipment_rows = portal.get("equipment") if isinstance(portal.get("equipment"), list) else []
    descriptions = {
        str(e.get("code") or e.get("equipment_code") or "").strip().upper(): str(e.get("description") or e.get("family") or "")
        for e in equipment_rows
        if isinstance(e, dict)
    }
    for row in mobile_capture_portal_rows(session):
        key = capture_merge_key(row)
        if not any(key):
            continue
        code = str(row.get("equipment_code") or "").strip().upper()
        if code and descriptions.get(code):
            row["equipment_description"] = descriptions[code]
        existing_index = index_by_key.get(key)
        received_at = parse_iso_datetime(row.get("received_at"))
        if portal_updated and portal_capture_is_closed_month(row):
            continue
        if existing_index is None:
            merged.append(row)
            index_by_key[key] = len(merged) - 1
        elif portal_updated and received_at and received_at > portal_updated:
            merged[existing_index] = row

    merged.sort(key=lambda row: (str(row.get("work_date") or ""), int(parse_float(row.get("id"), 0))), reverse=True)
    portal["captures"] = merged
    return portal


def tire_event_list(portal: dict[str, Any]) -> list[dict[str, Any]]:
    tracking = portal.get("tire_tracking") if isinstance(portal.get("tire_tracking"), dict) else {}
    events = tracking.get("events") if isinstance(tracking.get("events"), list) else []
    return [event for event in events if isinstance(event, dict)]


def tire_rows_list(portal: dict[str, Any]) -> list[dict[str, Any]]:
    tire = portal.get("tire_kpi") if isinstance(portal.get("tire_kpi"), dict) else {}
    rows = tire.get("rows") if isinstance(tire.get("rows"), list) else []
    return [row for row in rows if isinstance(row, dict)]


def tire_tread_remaining_percent_py(tread_initial: float, tread_current: float) -> float:
    if tread_initial <= 0 or tread_current <= 0:
        return 0.0
    return max(min((tread_current / tread_initial) * 100, 100), 0)


def tire_control_status_py(row: dict[str, Any]) -> tuple[str, str]:
    status = normalize_text(row.get("status"))
    if status == "BAJA":
        return "BAJA", "Fuera de servicio"
    if status == "REPARACION":
        return "REVISION", "Validar reparacion"
    if status == "ALMACEN":
        return "ALMACEN", "Disponible en almacen"
    life_percent = parse_float(row.get("life_percent") or row.get("tread_remaining_percent"), 0)
    remaining_hours = parse_float(row.get("life_remaining_hours"), 0)
    tread_initial = parse_float(row.get("tread_initial"), 0)
    tread_current = parse_float(row.get("tread_current"), 0)
    tread_percent = tire_tread_remaining_percent_py(tread_initial, tread_current)
    basis_percent = tread_percent if tread_percent > 0 else life_percent
    if basis_percent and basis_percent <= TIRE_REMAINING_CRITICAL_PERCENT:
        return "CRITICA", "Cambiar / dar de baja"
    if tread_initial > 0 and tread_current <= 0:
        return "REVISION", "Capturar piso actual"
    if (basis_percent and basis_percent <= TIRE_REMAINING_WARNING_PERCENT) or (remaining_hours and remaining_hours <= TIRE_REMAINING_WARNING_HOURS):
        return "PROXIMA", "Programar cambio"
    return "OK", "Seguimiento normal"


def normalize_tire_row(row: dict[str, Any]) -> dict[str, Any]:
    clean = dict(row)
    clean["tire_code"] = normalize_text(clean.get("tire_code") or clean.get("code"))
    clean["equipment_code"] = normalize_text(clean.get("equipment_code") or clean.get("equipment"))
    clean["position"] = normalize_text(clean.get("position"))[:80]
    clean["brand"] = normalize_text(clean.get("brand"))[:120]
    clean["model"] = normalize_text(clean.get("model"))[:120]
    clean["size"] = normalize_text(clean.get("size"))[:80]
    clean["status"] = normalize_text(clean.get("status") or "MONTADA")[:80]
    clean["install_date"] = str(clean.get("install_date") or "")[:20]
    clean["install_meter"] = max(parse_float(clean.get("install_meter"), 0), 0)
    clean["current_meter"] = max(parse_float(clean.get("current_meter"), 0), 0)
    clean["target_life_hours"] = max(parse_float(clean.get("target_life_hours"), 0), 0)
    clean["tread_initial"] = max(parse_float(clean.get("tread_initial"), 0), 0)
    clean["tread_current"] = max(parse_float(clean.get("tread_current"), 0), 0)
    clean["pressure_current"] = max(parse_float(clean.get("pressure_current"), 0), 0)
    clean["hours_used"] = max(clean["current_meter"] - clean["install_meter"], 0) if clean["current_meter"] or clean["install_meter"] else max(parse_float(clean.get("hours_used"), 0), 0)
    clean["life_remaining_hours"] = max(clean["target_life_hours"] - clean["hours_used"], 0) if clean["target_life_hours"] else max(parse_float(clean.get("life_remaining_hours"), 0), 0)
    tread_percent = tire_tread_remaining_percent_py(clean["tread_initial"], clean["tread_current"])
    if tread_percent > 0:
        clean["life_percent"] = tread_percent
        clean["tread_remaining_percent"] = tread_percent
        clean["life_text"] = f"{tread_percent:.0f}%"
    else:
        clean["life_percent"] = max(min(parse_float(clean.get("life_percent") or clean.get("tread_remaining_percent"), 0), 100), 0)
        clean["tread_remaining_percent"] = clean["life_percent"]
        clean["life_text"] = f"{clean['life_percent']:.0f}%" if clean["life_percent"] else "S/D"
    control_status, recommendation = tire_control_status_py(clean)
    clean["control_status"] = control_status
    clean["recommendation"] = normalize_text(clean.get("recommendation")) or recommendation
    clean["notes"] = str(clean.get("notes") or "").strip()
    return clean


def tire_summary_from_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    normalized = [normalize_tire_row(row) for row in rows if row.get("tire_code")]
    total = len(normalized)
    counts: dict[str, int] = {}
    for row in normalized:
        status = normalize_text(row.get("control_status") or "S/D")
        counts[status] = counts.get(status, 0) + 1
    avg_life = sum(parse_float(row.get("life_percent"), 0) for row in normalized) / total if total else 0
    avg_remaining = sum(parse_float(row.get("life_remaining_hours"), 0) for row in normalized) / total if total else 0
    return {
        "total": total,
        "critical": counts.get("CRITICA", 0),
        "soon": counts.get("PROXIMA", 0),
        "review": counts.get("REVISION", 0),
        "ok": counts.get("OK", 0),
        "warehouse": counts.get("ALMACEN", 0),
        "retired": counts.get("BAJA", 0),
        "avg_life": avg_life,
        "avg_remaining": avg_remaining,
    }


def enrich_tire_tracking(portal: dict[str, Any]) -> dict[str, Any]:
    rows = [normalize_tire_row(row) for row in tire_rows_list(portal) if row.get("tire_code")]
    rows.sort(
        key=lambda row: (
            row.get("equipment_code") or "ZZZ",
            {"CRITICA": 0, "PROXIMA": 1, "REVISION": 2, "OK": 3, "ALMACEN": 4, "BAJA": 5}.get(normalize_text(row.get("control_status")), 9),
            row.get("position") or "",
            row.get("tire_code") or "",
        )
    )
    portal["tire_kpi"] = {"rows": rows, "summary": tire_summary_from_rows(rows)}
    tracking = portal.get("tire_tracking") if isinstance(portal.get("tire_tracking"), dict) else {}
    events = sorted(tire_event_list(portal), key=lambda event: (str(event.get("event_date") or ""), int(parse_float(event.get("id"), 0))), reverse=True)
    portal["tire_tracking"] = {"events": events[:500]}
    return portal


PREVENTIVE_SERVICE_HOURS = {"PM1": 250, "PM2": 500, "PM3": 750, "PM4": 1000}
PREVENTIVE_CLOSED_STATUSES = {"CERRADO", "CERRADA", "TERMINADO", "TERMINADA", "FINALIZADO", "FINALIZADA"}


def preventive_execution_records(portal: dict[str, Any]) -> list[dict[str, Any]]:
    payload = portal.get("preventive_execution") if isinstance(portal.get("preventive_execution"), dict) else {}
    records = payload.get("records") if isinstance(payload, dict) else []
    return records if isinstance(records, list) else []


def normalize_preventive_execution_record(row: dict[str, Any]) -> dict[str, Any]:
    now = utc_now().isoformat(timespec="seconds")
    service_type = normalize_text(row.get("service_type") or "PM1")[:12] or "PM1"
    if service_type not in PREVENTIVE_SERVICE_HOURS:
        service_match = re.search(r"PM\s*([1-4])", service_type)
        service_type = f"PM{service_match.group(1)}" if service_match else service_type.replace(" ", "").upper()[:12] or "PM1"
    status = normalize_text(row.get("status") or "ABIERTO")[:30] or "ABIERTO"
    equipment_code = str(row.get("equipment_code") or row.get("equipment") or "").strip()[:120]
    record_id = str(row.get("id") or "").strip()
    if not record_id:
        record_id = f"WEB-PM-{int(datetime.now(timezone.utc).timestamp() * 1000)}"
    normalized = {
        "id": record_id[:80],
        "service_date": str(row.get("service_date") or row.get("date") or utc_now().date().isoformat())[:10],
        "close_date": str(row.get("close_date") or "")[:10],
        "equipment_code": equipment_code,
        "equipment_description": str(row.get("equipment_description") or "")[:220],
        "supervisor": str(row.get("supervisor") or "").strip()[:180],
        "mechanic": str(row.get("mechanic") or "").strip()[:180],
        "service_type": service_type,
        "service_hours": PREVENTIVE_SERVICE_HOURS.get(service_type, 0),
        "service_interval": f"{PREVENTIVE_SERVICE_HOURS.get(service_type, 0)}H" if PREVENTIVE_SERVICE_HOURS.get(service_type, 0) else "",
        "attribute_type": str(row.get("attribute_type") or "").strip()[:120],
        "status": status,
        "completed_meter": parse_float(row.get("completed_meter"), 0),
        "parts_used": str(row.get("parts_used") or "").strip(),
        "lubricants_used": str(row.get("lubricants_used") or "").strip(),
        "notes": str(row.get("notes") or "").strip(),
        "checklist": row.get("checklist") if isinstance(row.get("checklist"), dict) else {},
        "evidence_note": str(row.get("evidence_note") or "").strip(),
        "source": str(row.get("source") or "web")[:40],
        "created_at": str(row.get("created_at") or now)[:40],
        "updated_at": now,
    }
    oil_total = 0.0
    for key, _part_number, _description in OIL_STOCK_FIELDS:
        value = max(parse_float(row.get(key), 0), 0)
        normalized[key] = value
        oil_total += value
    normalized["oil_liters"] = oil_total
    if not normalized["lubricants_used"] and oil_total:
        normalized["lubricants_used"] = "; ".join(
            f"{part_number}: {normalized[key]:g} L"
            for key, part_number, _description in OIL_STOCK_FIELDS
            if normalized.get(key)
        )
    if normalized["status"] in PREVENTIVE_CLOSED_STATUSES and not normalized["close_date"]:
        normalized["close_date"] = normalized["service_date"]
    if normalized["status"] not in PREVENTIVE_CLOSED_STATUSES:
        normalized["close_date"] = ""
    return normalized


def preventive_record_to_service_history(row: dict[str, Any], equipment_lookup: dict[str, dict[str, Any]]) -> dict[str, Any]:
    service_type = normalize_text(row.get("service_type") or "PM") or "PM"
    service_hours = PREVENTIVE_SERVICE_HOURS.get(service_type, parse_float(row.get("service_hours"), 0))
    service_interval = f"{int(service_hours)}H" if service_hours else str(row.get("service_interval") or "")
    equipment_code = str(row.get("equipment_code") or "")
    equipment = equipment_lookup.get(equipment_code) or {}
    description = row.get("equipment_description") or equipment.get("description") or equipment.get("family") or ""
    people = [row.get("supervisor"), row.get("mechanic")]
    detail = " | ".join([f"{label}: {value}" for label, value in (("Supervisor", people[0]), ("Mecanico", people[1])) if value])
    if row.get("notes"):
        detail = f"{detail} | {row.get('notes')}" if detail else str(row.get("notes"))
    oils_used = {part_number: row.get(key) for key, part_number, _description in OIL_STOCK_FIELDS if parse_float(row.get(key), 0) > 0}
    return {
        "web_id": row.get("id"),
        "completed_date": row.get("close_date") or row.get("service_date") or "",
        "service_date": row.get("service_date") or "",
        "service_type": "Programado",
        "stage": service_type,
        "equipment_code": equipment_code,
        "equipment_description": description,
        "component": row.get("attribute_type") or "Preventivo",
        "service_name": service_type,
        "service_interval": service_interval,
        "scheduled_meter": "",
        "completed_meter": row.get("completed_meter") or 0,
        "due_date": "",
        "status": "A TIEMPO",
        "order_number": "",
        "document_name": "Registro web preventivo",
        "filters_text": row.get("parts_used") or "",
        "oils_used": oils_used or row.get("lubricants_used") or "",
        "notes": detail,
        "source": "preventive_execution_web",
    }


def preventive_execution_as_oil_capture(row: dict[str, Any]) -> dict[str, Any]:
    capture = {
        "work_date": row.get("close_date") or row.get("service_date") or "",
        "equipment_code": row.get("equipment_code") or "",
        "equipment_description": row.get("equipment_description") or "",
        "worked_hours": 0,
        "source": "preventive_execution_web",
    }
    total = 0.0
    for key, _part_number, _description in OIL_STOCK_FIELDS:
        value = max(parse_float(row.get(key), 0), 0)
        capture[key] = value
        total += value
    capture["oil_liters"] = total
    return capture


def merge_preventive_execution_into_portal(portal: dict[str, Any]) -> dict[str, Any]:
    records = [normalize_preventive_execution_record(row) for row in preventive_execution_records(portal) if isinstance(row, dict)]
    records.sort(key=lambda row: (str(row.get("service_date") or ""), str(row.get("id") or "")), reverse=True)
    portal["preventive_execution"] = {"records": records[:1000]}
    equipment_lookup: dict[str, dict[str, Any]] = {}
    for equipment in portal.get("equipment") or []:
        if isinstance(equipment, dict):
            code = str(equipment.get("code") or equipment.get("equipment_code") or "")
            if code:
                equipment_lookup[code] = equipment
    raw_history = portal.get("service_history") if isinstance(portal.get("service_history"), list) else []
    history = [
        row for row in raw_history
        if not (isinstance(row, dict) and str(row.get("source") or "") == "preventive_execution_web")
    ]
    closed_records = [
        row for row in records
        if normalize_text(row.get("status")) in PREVENTIVE_CLOSED_STATUSES
    ]
    additions = [preventive_record_to_service_history(row, equipment_lookup) for row in closed_records]
    portal["service_history"] = additions + history
    return portal


WORK_ORDER_CLOSED_STATUSES = {"CERRADA", "CERRADO", "CANCELADA", "CANCELADO"}


def work_order_records(portal: dict[str, Any]) -> list[dict[str, Any]]:
    payload = portal.get("work_orders") if isinstance(portal.get("work_orders"), dict) else {}
    records = payload.get("records") if isinstance(payload, dict) else []
    return records if isinstance(records, list) else []


def next_work_order_folio(records: list[dict[str, Any]]) -> str:
    current = 0
    for row in records:
        match = re.search(r"OT-(\d+)", str(row.get("folio") or row.get("id") or ""))
        if match:
            current = max(current, int(match.group(1)))
    return f"OT-{current + 1:05d}"


def normalize_work_order_record(row: dict[str, Any], existing: dict[str, Any] | None = None, folio: str | None = None) -> dict[str, Any]:
    now = utc_now().isoformat(timespec="seconds")
    status = normalize_text(row.get("status") or "ABIERTA")[:30] or "ABIERTA"
    record_id = str(row.get("id") or (existing.get("id") if existing else "") or "").strip()
    if not record_id:
        record_id = folio or str(row.get("folio") or "") or f"OT-{int(datetime.now(timezone.utc).timestamp() * 1000)}"
    clean = {
        "id": record_id[:80],
        "folio": str(row.get("folio") or folio or (existing.get("folio") if existing else "") or record_id)[:80],
        "date": str(row.get("date") or row.get("work_date") or utc_now().date().isoformat())[:10],
        "close_date": str(row.get("close_date") or "")[:10],
        "equipment_code": str(row.get("equipment_code") or row.get("equipment") or "").strip()[:120],
        "equipment_description": str(row.get("equipment_description") or "")[:220],
        "origin": normalize_text(row.get("origin") or "MANUAL")[:80],
        "priority": normalize_text(row.get("priority") or "MEDIA")[:30],
        "responsible": str(row.get("responsible") or "").strip()[:180],
        "mechanic": str(row.get("mechanic") or "").strip()[:180],
        "supervisor": str(row.get("supervisor") or "").strip()[:180],
        "status": status,
        "description": str(row.get("description") or row.get("detail") or "").strip(),
        "action": str(row.get("action") or "").strip(),
        "parts_used": str(row.get("parts_used") or "").strip(),
        "lubricants_used": str(row.get("lubricants_used") or "").strip(),
        "evidence_note": str(row.get("evidence_note") or "").strip(),
        "source_ref": str(row.get("source_ref") or "").strip()[:180],
        "created_at": str(existing.get("created_at") if existing else row.get("created_at") or now)[:40],
        "updated_at": now,
    }
    if clean["status"] in WORK_ORDER_CLOSED_STATUSES and not clean["close_date"]:
        clean["close_date"] = clean["date"]
    if clean["status"] not in WORK_ORDER_CLOSED_STATUSES:
        clean["close_date"] = ""
    if clean["priority"] not in {"ALTA", "MEDIA", "BAJA", "URGENTE"}:
        clean["priority"] = "MEDIA"
    return clean


def merge_work_orders_into_backlog(portal: dict[str, Any]) -> dict[str, Any]:
    records = [normalize_work_order_record(row) for row in work_order_records(portal) if isinstance(row, dict)]
    records.sort(key=lambda row: (str(row.get("date") or ""), str(row.get("folio") or "")), reverse=True)
    portal["work_orders"] = {"records": records[:1500]}
    return portal


def portal_fallback_payload(session: Session) -> dict[str, Any]:
    catalog = latest_catalog_payload(session)
    equipment = catalog.get("equipment") if isinstance(catalog, dict) else []
    if not isinstance(equipment, list):
        equipment = []
    captures = mobile_capture_portal_rows(session)
    now = utc_now()
    start = now.replace(day=1).date().isoformat()
    return {
        "ok": True,
        "source": "cloud-fallback",
        "generated_at": now.isoformat(timespec="seconds"),
        "period": {"start": start, "end": now.date().isoformat(), "year": now.year, "month": now.month},
        "settings": {
            "shift_hours": 9,
            "turns_per_day": 2,
            "meta_availability": 85,
            "meta_utilization": 75,
            "meta_reliability": 80,
            "meta_tmef": 8,
            "meta_tmpr": 4,
            "meta_diesel_lh": 25,
            "reliability_mission_hours": 24,
        },
        "equipment": equipment,
        "preventives": [],
        "service_history": [],
        "preventive_execution": {"records": []},
        "work_orders": {"records": []},
        "captures": captures,
        "availability": [],
        "kpi_groups": ["Todos los equipos", "Equipos de Barrenacion", "Equipos de Rezagado", "Acarreo", "Equipo Utilitario", "KPI Aceites", "KPI Llantas"],
        "kpi_reports": {},
        "oil_kpi": {"rows": [], "totals": {}, "columns": []},
        "tire_kpi": {"rows": [], "summary": {}},
        "tire_tracking": {"events": []},
        "diesel": {"records": [], "days": [], "rows": [], "totals": {}},
    }


def latest_portal_payload(session: Session) -> dict[str, Any]:
    snapshot = session.scalar(select(PortalSnapshot).where(PortalSnapshot.name == "default"))
    if snapshot is None:
        return portal_fallback_payload(session)
    payload = json_loads(snapshot.payload_json)
    if not isinstance(payload, dict):
        return portal_fallback_payload(session)
    payload.setdefault("ok", True)
    payload.setdefault("source", "cloud-portal")
    payload.setdefault("period", {})
    payload.setdefault("settings", {})
    payload.setdefault("equipment", [])
    payload.setdefault("captures", [])
    payload.setdefault("availability", [])
    payload.setdefault("preventives", [])
    payload.setdefault("service_history", [])
    payload.setdefault("preventive_execution", {"records": []})
    payload.setdefault("work_orders", {"records": []})
    payload.setdefault("kpi_groups", [])
    payload.setdefault("kpi_reports", {})
    payload.setdefault("oil_kpi", {"rows": [], "totals": {}, "columns": []})
    payload.setdefault("tire_kpi", {"rows": [], "summary": {}})
    payload.setdefault("tire_tracking", {"events": []})
    payload.setdefault("diesel", {"records": [], "days": [], "rows": [], "totals": {}})
    payload["updated_at"] = snapshot.updated_at.isoformat(timespec="seconds") if snapshot.updated_at else ""
    return enrich_tire_tracking(merge_work_orders_into_backlog(merge_preventive_execution_into_portal(merge_mobile_captures_into_portal(session, payload))))


def diesel_iso_or_none(value: Any) -> str | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text[:10] if len(text) >= 10 else None


def diesel_period_bounds(session: Session, start: str = "", end: str = "") -> tuple[str, str]:
    start_iso = diesel_iso_or_none(start)
    end_iso = diesel_iso_or_none(end)
    if not start_iso or not end_iso:
        portal = latest_portal_payload(session)
        diesel = portal.get("diesel") if isinstance(portal, dict) else {}
        if isinstance(diesel, dict):
            start_iso = start_iso or diesel_iso_or_none(diesel.get("start"))
            end_iso = end_iso or diesel_iso_or_none(diesel.get("end"))
        period = portal.get("period") if isinstance(portal, dict) else {}
        if isinstance(period, dict):
            start_iso = start_iso or diesel_iso_or_none(period.get("start"))
            end_iso = end_iso or diesel_iso_or_none(period.get("end"))
    today = utc_now().date()
    month_start = today.replace(day=1)
    if today.month == 12:
        month_end = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        month_end = today.replace(month=today.month + 1, day=1) - timedelta(days=1)
    start_iso = start_iso or month_start.isoformat()
    end_iso = end_iso or month_end.isoformat()
    if end_iso < start_iso:
        start_iso, end_iso = end_iso, start_iso
    return start_iso, end_iso


def diesel_record_key(row: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(row.get("work_date") or ""),
        normalize_text(row.get("equipment")),
        normalize_text(row.get("shift") or "1"),
    )


def diesel_deleted_record_keys(session: Session) -> set[tuple[str, str, str]]:
    rows = session.scalars(select(DieselDeletedRecord)).all()
    return {
        (
            row.work_date or "",
            normalize_text(row.equipment),
            normalize_text(row.shift or "1"),
        )
        for row in rows
    }


def diesel_deleted_day_dates(session: Session) -> set[str]:
    return {row.work_date for row in session.scalars(select(DieselDeletedDay)).all() if row.work_date}


def mark_diesel_record_deleted(session: Session, row: dict[str, Any]) -> tuple[str, str, str]:
    key = diesel_record_key(row)
    if not key[0] or not key[1]:
        raise HTTPException(status_code=400, detail="Captura diesel invalida.")
    existing = session.scalar(
        select(DieselDeletedRecord).where(
            DieselDeletedRecord.work_date == key[0],
            DieselDeletedRecord.equipment == key[1],
            DieselDeletedRecord.shift == key[2],
        )
    )
    if existing is None:
        session.add(DieselDeletedRecord(work_date=key[0], equipment=key[1], shift=key[2], deleted_at=utc_now()))
    else:
        existing.deleted_at = utc_now()
    return key


def unmark_diesel_record_deleted(session: Session, row: dict[str, Any]) -> None:
    key = diesel_record_key(row)
    for deleted in session.scalars(
        select(DieselDeletedRecord).where(
            DieselDeletedRecord.work_date == key[0],
            DieselDeletedRecord.equipment == key[1],
            DieselDeletedRecord.shift == key[2],
        )
    ).all():
        session.delete(deleted)


def mark_diesel_day_deleted(session: Session, work_date: str) -> None:
    if not work_date:
        raise HTTPException(status_code=400, detail="Fecha requerida.")
    existing = session.scalar(select(DieselDeletedDay).where(DieselDeletedDay.work_date == work_date))
    if existing is None:
        session.add(DieselDeletedDay(work_date=work_date, deleted_at=utc_now()))
    else:
        existing.deleted_at = utc_now()


def unmark_diesel_day_deleted(session: Session, work_date: str) -> None:
    if not work_date:
        return
    for deleted in session.scalars(select(DieselDeletedDay).where(DieselDeletedDay.work_date == work_date)).all():
        session.delete(deleted)


def hose_iso_or_none(value: Any) -> str:
    try:
        text = str(value or "").strip()
        if not text:
            return ""
        return datetime.fromisoformat(text[:10]).date().isoformat()
    except Exception:
        return ""


def hose_period_bounds(start: str = "", end: str = "") -> tuple[str, str]:
    start_iso = hose_iso_or_none(start)
    end_iso = hose_iso_or_none(end)
    if not start_iso or not end_iso:
        today = utc_now().date()
        start_date = today.replace(day=1)
        if start_date.month == 12:
            end_date = date(start_date.year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(start_date.year, start_date.month + 1, 1) - timedelta(days=1)
        start_iso = start_iso or start_date.isoformat()
        end_iso = end_iso or end_date.isoformat()
    if end_iso < start_iso:
        start_iso, end_iso = end_iso, start_iso
    return start_iso, end_iso


def mobile_hose_change_rows(record: dict[str, Any], source_device: str = "", user_name: str = "") -> list[dict[str, Any]]:
    if normalize_text(record.get("kind")) != "MANGUERA":
        return []
    hoses = record.get("hoses") or []
    if not isinstance(hoses, list):
        return []
    mobile_id = str(record.get("mobile_id") or "").strip()
    equipment = normalize_text(record.get("equipment_code") or record.get("equipment"))
    change_date = hose_iso_or_none(record.get("work_date")) or utc_now().date().isoformat()
    system = normalize_text(record.get("system") or "HIDRAULICO")
    location = str(record.get("location") or "").strip()
    work_type = str(record.get("work_type") or "").strip()
    failure_reason = str(record.get("failure_reason") or "").strip()
    meter = max(parse_float(record.get("meter"), 0), 0)
    downtime_hours = max(parse_float(record.get("downtime_hours"), 0), 0)
    details = str(record.get("details") or record.get("observations") or "").strip()
    supervisor = str(record.get("supervisor") or "").strip()
    mechanic = normalize_text(record.get("mechanic") or record.get("technician") or user_name)
    folio = str(record.get("capture_folio") or "").strip()
    rows: list[dict[str, Any]] = []
    for idx, item in enumerate(hoses, start=1):
        if not isinstance(item, dict):
            continue
        connection_type = str(item.get("connection_type") or "").strip()
        connection_number = str(item.get("connection_number") or "").strip()
        length = max(parse_float(item.get("length"), 0), 0)
        layers = str(item.get("hose_layers") or "").strip()
        if not connection_number and length <= 0:
            continue
        notes = " | ".join(
            part
            for part in (
                f"Conexion: {connection_type}" if connection_type else "",
                f"Malla: {layers}" if layers else "",
                f"Tipo: {work_type}" if work_type else "",
                f"Motivo: {failure_reason}" if failure_reason else "",
                f"Horometro consulta: {meter:g}" if meter > 0 else "",
                f"Horas paro: {downtime_hours:g}" if downtime_hours > 0 else "",
                f"Supervisor: {supervisor}" if supervisor else "",
                f"Dispositivo: {source_device}" if source_device else "",
                f"Folio: {folio}" if folio else "",
                details,
            )
            if part
        )
        rows.append(
            {
                "change_date": change_date,
                "equipment": equipment,
                "system": system,
                "part_type": "MANGUERA",
                "diameter": connection_number,
                "length_m": length,
                "quantity": 1,
                "unit_cost": 0,
                "estimated_life_days": 30,
                "estimated_weekly_qty": 0,
                "failure_reason": " - ".join(part for part in (failure_reason, location) if part),
                "technician": mechanic,
                "notes": notes,
                "source": "mobile",
                "external_id": f"mobile:{mobile_id}:{idx}" if mobile_id else "",
            }
        )
    return rows


def hose_change_payload(row: HoseChange) -> dict[str, Any]:
    return {
        "id": row.id,
        "source": row.source or "web",
        "external_id": row.external_id or "",
        "change_date": row.change_date,
        "equipment": row.equipment,
        "system": row.system,
        "part_type": row.part_type,
        "diameter": row.diameter,
        "length_m": row.length_m,
        "quantity": row.quantity,
        "unit_cost": row.unit_cost,
        "estimated_life_days": row.estimated_life_days,
        "estimated_weekly_qty": row.estimated_weekly_qty,
        "failure_reason": row.failure_reason,
        "technician": row.technician,
        "notes": row.notes,
        "updated_at": row.updated_at.isoformat(timespec="seconds") if row.updated_at else "",
    }


def upsert_hose_change(session: Session, payload: dict[str, Any]) -> tuple[HoseChange, bool]:
    external_id = str(payload.get("external_id") or "").strip()
    row = None
    if external_id:
        row = session.scalar(select(HoseChange).where(HoseChange.external_id == external_id))
    created = False
    if row is None:
        row = HoseChange(created_at=utc_now())
        session.add(row)
        created = True
    row.change_date = hose_iso_or_none(payload.get("change_date")) or utc_now().date().isoformat()
    row.equipment = normalize_text(payload.get("equipment"))
    row.system = normalize_text(payload.get("system"))
    row.part_type = normalize_text(payload.get("part_type") or "MANGUERA")[:80] or "MANGUERA"
    row.diameter = normalize_text(payload.get("diameter"))[:80]
    row.length_m = max(parse_float(payload.get("length_m"), 0), 0)
    row.quantity = max(parse_float(payload.get("quantity"), 1), 0)
    row.unit_cost = max(parse_float(payload.get("unit_cost"), 0), 0)
    row.estimated_life_days = max(parse_float(payload.get("estimated_life_days"), 30), 0)
    row.estimated_weekly_qty = max(parse_float(payload.get("estimated_weekly_qty"), 0), 0)
    row.failure_reason = normalize_text(payload.get("failure_reason"))[:220]
    row.technician = normalize_text(payload.get("technician"))[:180]
    row.notes = str(payload.get("notes") or "").strip()
    row.source = str(payload.get("source") or "web").strip()[:80]
    row.external_id = external_id
    row.updated_at = utc_now()
    return row, created


def hose_report(session: Session, start: str = "", end: str = "", equipment: str = "") -> dict[str, Any]:
    start_iso, end_iso = hose_period_bounds(start, end)
    query = select(HoseChange).where(HoseChange.change_date >= start_iso, HoseChange.change_date <= end_iso)
    selected_equipment = normalize_text(equipment)
    if selected_equipment and selected_equipment not in {"TODOS", "TODOS LOS EQUIPOS"}:
        query = query.where(HoseChange.equipment == selected_equipment)
    rows = session.scalars(query.order_by(HoseChange.change_date.desc(), HoseChange.id.desc())).all()
    start_date = datetime.fromisoformat(start_iso).date()
    end_date = datetime.fromisoformat(end_iso).date()
    period_days = max((end_date - start_date).days + 1, 1)
    records: list[dict[str, Any]] = []
    summary: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        payload = hose_change_payload(row)
        qty = max(parse_float(payload.get("quantity"), 0), 0)
        unit_cost = max(parse_float(payload.get("unit_cost"), 0), 0)
        life_days = max(parse_float(payload.get("estimated_life_days"), 0), 0)
        weekly = max(parse_float(payload.get("estimated_weekly_qty"), 0), 0)
        if weekly <= 0 and life_days > 0:
            weekly = qty * 7 / life_days
        estimated_qty = max(weekly * period_days / 7, 0)
        total_cost = qty * unit_cost
        payload["estimated_qty_period"] = estimated_qty
        payload["total_cost"] = total_cost
        records.append(payload)
        key = (payload.get("equipment") or "SIN EQUIPO", payload.get("part_type") or "MANGUERA")
        bucket = summary.setdefault(
            key,
            {
                "equipment": key[0],
                "part_type": key[1],
                "systems": set(),
                "changes": 0,
                "quantity": 0.0,
                "estimated_qty": 0.0,
                "variance": 0.0,
                "total_cost": 0.0,
            },
        )
        bucket["changes"] += 1
        bucket["quantity"] += qty
        bucket["estimated_qty"] += estimated_qty
        bucket["total_cost"] += total_cost
        if payload.get("system"):
            bucket["systems"].add(payload["system"])
    summary_rows = []
    for bucket in summary.values():
        quantity = bucket["quantity"]
        estimated_qty = bucket["estimated_qty"]
        summary_rows.append(
            {
                **bucket,
                "systems": ", ".join(sorted(bucket["systems"])),
                "variance": quantity - estimated_qty,
            }
        )
    summary_rows.sort(key=lambda row: (-row["quantity"], row["equipment"], row["part_type"]))
    real_qty = sum(row["quantity"] for row in summary_rows)
    estimated_qty = sum(row["estimated_qty"] for row in summary_rows)
    return {
        "start": start_iso,
        "end": end_iso,
        "period_days": period_days,
        "records": records,
        "summary": summary_rows,
        "totals": {
            "changes": len(records),
            "real_qty": real_qty,
            "estimated_qty": estimated_qty,
            "variance": real_qty - estimated_qty,
            "total_cost": sum(row["total_cost"] for row in summary_rows),
            "critical_equipment": summary_rows[0]["equipment"] if summary_rows else "S/D",
        },
    }


def diesel_record_payload(row: DieselRecord) -> dict[str, Any]:
    return {
        "id": row.id,
        "source": row.source or "web",
        "work_date": row.work_date,
        "equipment": row.equipment,
        "condition": row.condition,
        "shift": row.shift,
        "horometer_initial": row.horometer_initial,
        "horometer_final": row.horometer_final,
        "worked_hours": row.worked_hours,
        "diesel_liters": row.diesel_liters,
        "operator": row.operator,
        "dispatcher": row.dispatcher,
        "supervisor": row.supervisor,
        "notes": row.notes,
        "updated_at": row.updated_at.isoformat(timespec="seconds") if row.updated_at else "",
    }


def diesel_day_payload(row: DieselDay) -> dict[str, Any]:
    return {
        "id": row.id,
        "source": row.source or "web",
        "work_date": row.work_date,
        "diesel_received": row.diesel_received,
        "initial_stock": row.initial_stock,
        "final_stock": row.final_stock,
        "prosermin_stock": row.prosermin_stock,
        "supplier": row.supplier,
        "notes": row.notes,
        "updated_at": row.updated_at.isoformat(timespec="seconds") if row.updated_at else "",
    }


def clean_diesel_record_dict(row: dict[str, Any], source: str = "desktop") -> dict[str, Any]:
    hi = parse_float(row.get("horometer_initial"), 0)
    hf = parse_float(row.get("horometer_final"), 0)
    worked = parse_float(row.get("worked_hours"), -1)
    if worked < 0:
        worked = max(hf - hi, 0) if hi and hf and hf >= hi else 0
    return {
        "id": row.get("id") or "",
        "source": row.get("source") or source,
        "work_date": diesel_iso_or_none(row.get("work_date")) or "",
        "equipment": normalize_text(row.get("equipment")),
        "condition": normalize_text(row.get("condition") or "DISPONIBLE"),
        "shift": normalize_text(row.get("shift") or "1"),
        "horometer_initial": hi,
        "horometer_final": hf,
        "worked_hours": max(worked, 0),
        "diesel_liters": max(parse_float(row.get("diesel_liters"), 0), 0),
        "operator": normalize_text(row.get("operator")),
        "dispatcher": normalize_text(row.get("dispatcher")),
        "supervisor": normalize_text(row.get("supervisor")),
        "notes": str(row.get("notes") or "").strip(),
        "updated_at": str(row.get("updated_at") or ""),
    }


def clean_diesel_day_dict(row: dict[str, Any], source: str = "desktop") -> dict[str, Any]:
    return {
        "id": row.get("id") or "",
        "source": row.get("source") or source,
        "work_date": diesel_iso_or_none(row.get("work_date")) or "",
        "diesel_received": max(parse_float(row.get("diesel_received"), 0), 0),
        "initial_stock": max(parse_float(row.get("initial_stock"), 0), 0),
        "final_stock": max(parse_float(row.get("final_stock"), 0), 0),
        "prosermin_stock": max(parse_float(row.get("prosermin_stock"), 0), 0),
        "supplier": normalize_text(row.get("supplier")),
        "notes": str(row.get("notes") or "").strip(),
        "updated_at": str(row.get("updated_at") or ""),
    }


def diesel_supplier_owner(value: Any) -> str:
    return "PROSERMIN" if "PROSERMIN" in normalize_text(value) else "MGA"


def diesel_payload(session: Session, start: str = "", end: str = "", meta_lh: float | None = None) -> dict[str, Any]:
    start_iso, end_iso = diesel_period_bounds(session, start, end)
    portal = latest_portal_payload(session)
    settings = portal.get("settings") if isinstance(portal, dict) else {}
    diesel_portal = portal.get("diesel") if isinstance(portal, dict) else {}
    meta = parse_float(meta_lh, 0) or parse_float((settings or {}).get("meta_diesel_lh"), 25) or 25
    deleted_record_keys = diesel_deleted_record_keys(session)
    deleted_day_dates = diesel_deleted_day_dates(session)
    records_by_key: dict[tuple[str, str, str], dict[str, Any]] = {}
    if isinstance(diesel_portal, dict):
        for raw in diesel_portal.get("records") or []:
            if not isinstance(raw, dict):
                continue
            record = clean_diesel_record_dict(raw, "desktop")
            if record["work_date"] and start_iso <= record["work_date"] <= end_iso and diesel_record_key(record) not in deleted_record_keys:
                records_by_key[diesel_record_key(record)] = record
    db_records = session.scalars(
        select(DieselRecord)
        .where(DieselRecord.work_date >= start_iso, DieselRecord.work_date <= end_iso)
        .order_by(DieselRecord.work_date.desc(), DieselRecord.equipment.asc(), DieselRecord.shift.asc())
    ).all()
    for row in db_records:
        record = diesel_record_payload(row)
        if diesel_record_key(record) not in deleted_record_keys:
            records_by_key[diesel_record_key(record)] = record
    equipment_aliases = diesel_equipment_alias_map(portal, list(records_by_key.values()))
    records = []
    for record in records_by_key.values():
        normalized_record = dict(record)
        raw_equipment = normalize_text(normalized_record.get("equipment"))
        canonical_equipment = diesel_canonical_equipment(raw_equipment, equipment_aliases)
        if canonical_equipment and raw_equipment and canonical_equipment != raw_equipment:
            normalized_record["raw_equipment"] = raw_equipment
        normalized_record["equipment"] = canonical_equipment or raw_equipment
        if diesel_record_key(normalized_record) in deleted_record_keys:
            continue
        records.append(normalized_record)
    records = sorted(records, key=lambda row: (row["work_date"], row["equipment"], row["shift"]), reverse=True)
    bitacora_hours = diesel_bitacora_hours_by_equipment(portal, start_iso, end_iso, equipment_aliases)

    days_by_date: dict[str, dict[str, Any]] = {}
    if isinstance(diesel_portal, dict):
        for raw in diesel_portal.get("days") or []:
            if not isinstance(raw, dict):
                continue
            day = clean_diesel_day_dict(raw, "desktop")
            if day["work_date"] and start_iso <= day["work_date"] <= end_iso and day["work_date"] not in deleted_day_dates:
                days_by_date[day["work_date"]] = day
    db_days = session.scalars(
        select(DieselDay).where(DieselDay.work_date >= start_iso, DieselDay.work_date <= end_iso)
    ).all()
    for row in db_days:
        day = diesel_day_payload(row)
        if day["work_date"] not in deleted_day_dates:
            days_by_date[day["work_date"]] = day

    buckets: dict[str, dict[str, Any]] = {}
    consumption_by_date: dict[str, float] = {}
    for record in records:
        equipment = normalize_text(record.get("equipment"))
        if not equipment:
            continue
        bucket = buckets.setdefault(
            equipment,
            {
                "equipment": equipment,
                "condition": record.get("condition") or "DISPONIBLE",
                "hi_values": [],
                "hf_values": [],
                "worked_hours": 0.0,
                "diesel_liters": 0.0,
            },
        )
        bucket["condition"] = record.get("condition") or bucket["condition"]
        hi = parse_float(record.get("horometer_initial"), 0)
        hf = parse_float(record.get("horometer_final"), 0)
        if hi > 0:
            bucket["hi_values"].append(hi)
        if hf > 0:
            bucket["hf_values"].append(hf)
        bucket["worked_hours"] += parse_float(record.get("worked_hours"), 0)
        bucket["diesel_liters"] += parse_float(record.get("diesel_liters"), 0)
        work_date = record.get("work_date") or ""
        consumption_by_date[work_date] = consumption_by_date.get(work_date, 0) + parse_float(record.get("diesel_liters"), 0)

    rows = []
    for bucket in buckets.values():
        bitacora = bitacora_hours.get(bucket["equipment"], {})
        bitacora_worked = parse_float(bitacora.get("worked_hours"), 0)
        worked = bitacora_worked if bitacora_worked > 0 else bucket["worked_hours"]
        liters = bucket["diesel_liters"]
        hi_values = bitacora.get("hi_values") or bucket["hi_values"]
        hf_values = bitacora.get("hf_values") or bucket["hf_values"]
        rendimiento = liters / worked if worked > 0 else None
        if liters <= 0:
            status = "SIN CONSUMO"
        elif worked <= 0:
            status = "SIN HORAS"
        elif rendimiento is not None and rendimiento > meta:
            status = "ALTO"
        else:
            status = "OK"
        rows.append(
            {
                "equipment": bucket["equipment"],
                "condition": bucket["condition"],
                "horometer_initial": min(hi_values) if hi_values else 0,
                "horometer_final": max(hf_values) if hf_values else 0,
                "worked_hours": worked,
                "hours_source": "bitacora" if bitacora_worked > 0 else "diesel",
                "diesel_liters": liters,
                "rendimiento_lh": rendimiento,
                "status": status,
            }
        )
    rows.sort(key=lambda item: (-item["diesel_liters"], item["equipment"]))

    daily_rows = []
    cursor = datetime.strptime(start_iso, "%Y-%m-%d").date()
    finish = datetime.strptime(end_iso, "%Y-%m-%d").date()
    while cursor <= finish:
        key = cursor.isoformat()
        if key in deleted_day_dates:
            cursor += timedelta(days=1)
            continue
        day = days_by_date.get(key, {})
        supplier = day.get("supplier") or ""
        owner = diesel_supplier_owner(supplier)
        daily_rows.append(
            {
                "id": day.get("id") or "",
                "source": day.get("source") or ("desktop" if day else ""),
                "work_date": key,
                "diesel_liters": consumption_by_date.get(key, 0),
                "diesel_received": parse_float(day.get("diesel_received"), 0),
                "initial_stock": parse_float(day.get("initial_stock"), 0),
                "final_stock": parse_float(day.get("final_stock"), 0),
                "prosermin_stock": parse_float(day.get("prosermin_stock"), 0),
                "supplier": supplier,
                "supplier_owner": owner,
                "notes": day.get("notes") or "",
            }
        )
        cursor += timedelta(days=1)

    total_liters = sum(parse_float(row["diesel_liters"], 0) for row in rows)
    total_hours = sum(parse_float(row["worked_hours"], 0) for row in rows)
    mga_liters = sum(parse_float(row.get("diesel_liters"), 0) for row in daily_rows if row.get("supplier_owner") == "MGA")
    prosermin_liters = sum(parse_float(row.get("diesel_liters"), 0) for row in daily_rows if row.get("supplier_owner") == "PROSERMIN")
    mga_received = sum(parse_float(row.get("diesel_received"), 0) for row in daily_rows if row.get("supplier_owner") == "MGA")
    prosermin_received = sum(parse_float(row.get("diesel_received"), 0) for row in daily_rows if row.get("supplier_owner") == "PROSERMIN")
    stock_candidates = [
        row for row in daily_rows
        if parse_float(row.get("final_stock"), 0) > 0
        or parse_float(row.get("initial_stock"), 0) > 0
        or parse_float(row.get("diesel_received"), 0) > 0
        or parse_float(row.get("diesel_liters"), 0) > 0
    ]
    latest_stock_row = stock_candidates[-1] if stock_candidates else {}
    total_stock = parse_float(latest_stock_row.get("final_stock"), 0)
    prosermin_stock = parse_float(latest_stock_row.get("prosermin_stock"), 0)
    if prosermin_stock <= 0 and isinstance(diesel_portal, dict):
        prosermin_stock = parse_float((diesel_portal.get("totals") or {}).get("prosermin_stock"), 0)
    prosermin_stock = max(min(prosermin_stock, total_stock), 0)
    mga_stock = max(total_stock - prosermin_stock, 0)
    equipment_codes = sorted(
        {
            normalize_text(item.get("code") or item.get("equipment_code") or "")
            for item in (portal.get("equipment") if isinstance(portal, dict) else []) or []
            if normalize_text(item.get("code") or item.get("equipment_code") or "")
        }
        | {row["equipment"] for row in rows}
    )
    return {
        "ok": True,
        "start": start_iso,
        "end": end_iso,
        "meta_lh": meta,
        "equipment": equipment_codes,
        "records": records,
        "days": daily_rows,
        "rows": rows,
        "totals": {
            "diesel_liters": total_liters,
            "worked_hours": total_hours,
            "rendimiento_lh": total_liters / total_hours if total_hours > 0 else None,
            "received": sum(parse_float(row.get("diesel_received"), 0) for row in daily_rows),
            "mga_liters": mga_liters,
            "prosermin_liters": prosermin_liters,
            "mga_received": mga_received,
            "prosermin_received": prosermin_received,
            "total_stock": total_stock,
            "mga_stock": mga_stock,
            "prosermin_stock": prosermin_stock,
            "critical": sum(1 for row in rows if row["status"] in {"ALTO", "SIN HORAS"}),
        },
        "updated_at": portal.get("updated_at") or portal.get("generated_at") or utc_now().isoformat(timespec="seconds"),
    }


def diesel_equipment_match_key(value: Any) -> str:
    return "".join(ch for ch in normalize_text(value) if ch.isalnum())


def diesel_equipment_code_key(value: Any) -> str:
    match = re.search(r"([A-Z]{1,4})[- ]?0*(\d{1,4})", normalize_text(value))
    if not match:
        return ""
    return f"{match.group(1)}{int(match.group(2))}"


def diesel_canonical_equipment(value: Any, aliases: dict[str, str] | None = None) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    aliases = aliases or {}
    for key in diesel_template_equipment_keys(text):
        if key in aliases:
            return aliases[key]
    without_parentheses = re.sub(r"\([^)]*\)", "", text).strip()
    first_token = re.split(r"[\s(]+", text, 1)[0].strip()
    if first_token and any(ch.isdigit() for ch in first_token):
        return first_token
    return without_parentheses or text


def diesel_equipment_alias_map(
    portal: dict[str, Any] | None,
    records: list[dict[str, Any]] | None = None,
) -> dict[str, str]:
    aliases: dict[str, str] = {}

    def add_alias(canonical: Any, *values: Any) -> None:
        canonical_text = normalize_text(canonical)
        if not canonical_text:
            return
        for value in (canonical_text, *values):
            for key in diesel_template_equipment_keys(value):
                aliases.setdefault(key, canonical_text)
            key = diesel_equipment_match_key(value)
            if key:
                aliases.setdefault(key, canonical_text)

    equipment_rows = portal.get("equipment") if isinstance(portal, dict) else []
    if isinstance(equipment_rows, list):
        for item in equipment_rows:
            if not isinstance(item, dict):
                continue
            code = normalize_text(item.get("code") or item.get("equipment_code") or "")
            description = normalize_text(item.get("description") or item.get("family") or "")
            if not code:
                continue
            add_alias(code, f"{code} {description}", f"{code} ({description})")
            if code.startswith("RET-"):
                short_code = code.replace("RET-", "RE-", 1)
                add_alias(code, short_code, f"{short_code} {description}", f"{short_code} ({description})")
            elif code.startswith("RE-"):
                long_code = code.replace("RE-", "RET-", 1)
                add_alias(code, long_code, f"{long_code} {description}", f"{long_code} ({description})")

    diesel_portal = portal.get("diesel") if isinstance(portal, dict) else {}
    diesel_equipment = diesel_portal.get("equipment") if isinstance(diesel_portal, dict) else []
    if isinstance(diesel_equipment, list):
        for item in diesel_equipment:
            text = normalize_text(item)
            if text:
                add_alias(diesel_canonical_equipment(text, aliases), text)

    for record in records or []:
        text = normalize_text(record.get("equipment") if isinstance(record, dict) else record)
        if text:
            add_alias(diesel_canonical_equipment(text, aliases), text)

    return aliases


def diesel_bitacora_hours_by_equipment(
    portal: dict[str, Any] | None,
    start: str,
    end: str,
    aliases: dict[str, str],
) -> dict[str, dict[str, Any]]:
    captures = portal.get("captures") if isinstance(portal, dict) else []
    if not isinstance(captures, list):
        return {}
    buckets: dict[str, dict[str, Any]] = {}
    for capture in captures:
        if not isinstance(capture, dict):
            continue
        work_date = str(capture.get("work_date") or "")
        if not work_date or work_date < start or work_date > end:
            continue
        equipment = diesel_canonical_equipment(
            capture.get("equipment_code") or capture.get("equipment") or capture.get("code") or capture.get("eco"),
            aliases,
        )
        if not equipment:
            continue
        bucket = buckets.setdefault(equipment, {"hi_values": [], "hf_values": [], "worked_hours": 0.0})
        hi = parse_float(capture.get("hi") or capture.get("horometer_initial"), 0)
        hf = parse_float(capture.get("hf") or capture.get("horometer_final"), 0)
        worked = parse_float(capture.get("worked_hours"), 0)
        if worked <= 0 and hi > 0 and hf >= hi:
            worked = hf - hi
        if hi > 0:
            bucket["hi_values"].append(hi)
        if hf > 0:
            bucket["hf_values"].append(hf)
        bucket["worked_hours"] += max(worked, 0)
    return buckets


def diesel_template_equipment_keys(value: Any) -> list[str]:
    text = normalize_text(value)
    if not text or text.startswith("=") or text == "TOTALES":
        return []
    candidates = [text, re.sub(r"\([^)]*\)", "", text).strip()]
    first_token = re.split(r"[\s(]+", text, 1)[0].strip()
    if first_token:
        candidates.append(first_token)
    candidates.extend(re.findall(r"[A-Z]{1,4}[- ]?\d{2,4}", text))
    keys: list[str] = []
    for candidate in candidates:
        for key in (diesel_equipment_match_key(candidate), diesel_equipment_code_key(candidate)):
            if key and key not in keys:
                keys.append(key)
    return keys


def diesel_shift_key(value: Any) -> str:
    text = normalize_text(value)
    if "2" in text or "SEG" in text:
        return "2"
    return "1"


def diesel_month_year_from_start(start: str) -> tuple[int, int]:
    try:
        parsed = datetime.strptime(start, "%Y-%m-%d").date()
    except Exception:
        parsed = utc_now().date()
    return parsed.month, parsed.year


def diesel_template_row_map(wb) -> dict[str, int]:
    if "DIA 01" not in wb.sheetnames:
        return {}
    ws = wb["DIA 01"]
    row_map: dict[str, int] = {}
    for row_idx in range(7, min(ws.max_row, 35)):
        for key in diesel_template_equipment_keys(ws.cell(row_idx, 1).value):
            row_map.setdefault(key, row_idx)
    return row_map


def diesel_aggregate_records(records: list[dict[str, Any]]) -> dict[tuple[str, str, str], dict[str, Any]]:
    buckets: dict[tuple[str, str, str], dict[str, Any]] = {}
    for record in records:
        work_date = diesel_iso_or_none(record.get("work_date")) or ""
        key = diesel_equipment_match_key(record.get("equipment"))
        if not work_date or not key:
            continue
        shift = diesel_shift_key(record.get("shift"))
        bucket = buckets.setdefault(
            (work_date, key, shift),
            {
                "horometer_initial": 0.0,
                "horometer_final": 0.0,
                "diesel_liters": 0.0,
                "operator": "",
                "dispatcher": "",
                "supervisor": "",
            },
        )
        hi = parse_float(record.get("horometer_initial"), 0)
        hf = parse_float(record.get("horometer_final"), 0)
        if hi > 0 and (not bucket["horometer_initial"] or hi < bucket["horometer_initial"]):
            bucket["horometer_initial"] = hi
        if hf > 0 and hf > bucket["horometer_final"]:
            bucket["horometer_final"] = hf
        bucket["diesel_liters"] += max(parse_float(record.get("diesel_liters"), 0), 0)
        for field in ("operator", "dispatcher", "supervisor"):
            value = normalize_text(record.get(field))
            if value:
                bucket[field] = value
    return buckets


def set_excel_value_or_blank(ws, row_idx: int, col_idx: int, value: Any) -> None:
    number = parse_float(value, 0)
    ws.cell(row_idx, col_idx).value = number if number else None


def populate_diesel_workbook(wb, payload: dict[str, Any]) -> None:
    month_names = [
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
    ]
    start = payload.get("start") or utc_now().date().isoformat()
    month, year = diesel_month_year_from_start(start)
    row_map = diesel_template_row_map(wb)
    buckets = diesel_aggregate_records(payload.get("records") or [])
    days_by_date = {row.get("work_date"): row for row in (payload.get("days") or []) if isinstance(row, dict)}
    report_rows = {}
    for row in payload.get("rows") or []:
        if isinstance(row, dict):
            key = diesel_equipment_match_key(row.get("equipment"))
            if key:
                report_rows[key] = row

    for day in range(1, 32):
        sheet_name = f"DIA {day:02d}"
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        try:
            work_date = datetime(year, month, day).date().isoformat()
        except ValueError:
            continue
        ws["B3"] = day
        ws["C3"] = month_names[month - 1]
        ws["D3"] = year
        day_meta = days_by_date.get(work_date, {})
        ws["O2"] = parse_float(day_meta.get("diesel_received"), 0)
        ws["P2"] = parse_float(day_meta.get("initial_stock"), 0)
        ws["Q2"] = "=P2-O35"
        for row_idx in range(7, 35):
            for col_idx in (2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14):
                ws.cell(row_idx, col_idx).value = None
            ws.cell(row_idx, 15).value = f"=C{row_idx}+J{row_idx}"
            ws.cell(row_idx, 16).value = f"=G{row_idx}-B{row_idx}"
            ws.cell(row_idx, 17).value = f"=N{row_idx}-I{row_idx}"
        ws["C35"] = "=SUM(C7:C34)"
        ws["J35"] = "=SUM(J7:J34)"
        ws["O35"] = "=C35+J35"
        for key, row_idx in row_map.items():
            for shift, columns in (("1", (2, 3, 4, 5, 6, 7)), ("2", (9, 10, 11, 12, 13, 14))):
                bucket = buckets.get((work_date, key, shift))
                if not bucket:
                    continue
                hi_col, liters_col, operator_col, dispatcher_col, supervisor_col, hf_col = columns
                set_excel_value_or_blank(ws, row_idx, hi_col, bucket["horometer_initial"])
                set_excel_value_or_blank(ws, row_idx, liters_col, bucket["diesel_liters"])
                ws.cell(row_idx, operator_col).value = bucket["operator"] or None
                ws.cell(row_idx, dispatcher_col).value = bucket["dispatcher"] or None
                ws.cell(row_idx, supervisor_col).value = bucket["supervisor"] or None
                set_excel_value_or_blank(ws, row_idx, hf_col, bucket["horometer_final"])

    if "CONSUMO DIARIO" in wb.sheetnames:
        ws = wb["CONSUMO DIARIO"]
        for day in range(1, 32):
            row_idx = day + 3
            ws.cell(row_idx, 1).value = day
            ws.cell(row_idx, 2).value = f"='DIA {day:02d}'!$O$35"
            ws.cell(row_idx, 3).value = f"='DIA {day:02d}'!$O$2"
            ws.cell(row_idx, 4).value = f"='DIA {day:02d}'!$Q$2"
            try:
                work_date = datetime(year, month, day).date().isoformat()
            except ValueError:
                continue
            ws.cell(row_idx, 6).value = (days_by_date.get(work_date, {}).get("supplier") or None)
        totals = payload.get("totals") or {}
        ws["F23"] = "MGA"
        ws["G23"] = parse_float(totals.get("mga_stock"), 0) or 0
        ws["F36"] = "PROSERMIN"
        ws["G36"] = parse_float(totals.get("prosermin_stock"), 0) or 0

    if "RENDIMIENTO" in wb.sheetnames:
        ws = wb["RENDIMIENTO"]
        dia = wb["DIA 01"] if "DIA 01" in wb.sheetnames else None
        for report_row_idx, day_row_idx in zip(range(6, 34), range(7, 35)):
            keys = diesel_template_equipment_keys(dia.cell(day_row_idx, 1).value) if dia else []
            row = next((report_rows[key] for key in keys if key in report_rows), None)
            ws.cell(report_row_idx, 2).value = f"='DIA 01'!A{day_row_idx}"
            ws.cell(report_row_idx, 3).value = (row or {}).get("condition") or "DISPONIBLE"
            hi = parse_float((row or {}).get("horometer_initial"), 0)
            hf = parse_float((row or {}).get("horometer_final"), 0)
            ws.cell(report_row_idx, 4).value = hi if hi else None
            ws.cell(report_row_idx, 5).value = hf if hf else None
            ws.cell(report_row_idx, 6).value = f"=E{report_row_idx}-D{report_row_idx}"
            ws.cell(report_row_idx, 7).value = "=" + "+".join([f"'DIA {day:02d}'!O{day_row_idx}" for day in range(1, 32)])
            ws.cell(report_row_idx, 8).value = f"=G{report_row_idx}/F{report_row_idx}"

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass


def add_diesel_logo_to_workbook(wb) -> None:
    if not DIESEL_LOGO_PATH.exists():
        return
    for ws in wb.worksheets:
        try:
            if getattr(ws, "_images", []):
                continue
            logo = ExcelImage(str(DIESEL_LOGO_PATH))
            logo.width = 118
            logo.height = 54
            ws.add_image(logo, "A1")
        except Exception:
            continue


def diesel_workbook_bytes(payload: dict[str, Any]) -> bytes:
    if DIESEL_TEMPLATE_PATH.exists():
        wb = load_workbook(DIESEL_TEMPLATE_PATH, data_only=False)
    else:
        wb = Workbook()
        wb.active.title = "DIA 01"
        for day in range(2, 32):
            wb.create_sheet(f"DIA {day:02d}")
        wb.create_sheet("CONSUMO DIARIO")
        wb.create_sheet("RENDIMIENTO")
    try:
        populate_diesel_workbook(wb, payload)
        add_diesel_logo_to_workbook(wb)
        out = BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()
    finally:
        wb.close()


def filter_match_keys(item: dict[str, Any]) -> list[str]:
    keys: list[str] = []
    for field in ("part_number", "donaldson_part", "matched_part"):
        key = normalize_part_key(item.get(field))
        if key and key not in keys:
            keys.append(key)
    return keys


def inventory_catalog_details(equipment: list[Any]) -> dict[str, dict[str, str]]:
    raw: dict[str, dict[str, list[str]]] = {}
    for equipment_item in equipment:
        if not isinstance(equipment_item, dict):
            continue
        code = str(equipment_item.get("code") or "").strip()
        description = str(equipment_item.get("description") or equipment_item.get("family") or "").strip()
        equipment_label = f"{code} - {description}" if code and description else code or description
        filters = equipment_item.get("filters")
        if not isinstance(filters, list):
            continue
        for filter_item in filters:
            if not isinstance(filter_item, dict):
                continue
            for key in filter_match_keys(filter_item):
                bucket = raw.setdefault(
                    key,
                    {
                        "equipment": [],
                        "equipment_codes": [],
                        "item_type": [],
                        "description": [],
                        "service_interval": [],
                    },
                )
                add_unique(bucket["equipment"], equipment_label)
                add_unique(bucket["equipment_codes"], code)
                add_unique(bucket["item_type"], filter_item.get("item_type"))
                add_unique(bucket["description"], filter_item.get("description"))
                add_unique(bucket["service_interval"], filter_item.get("service_interval"))
    return {key: {field: "; ".join(values) for field, values in bucket.items()} for key, bucket in raw.items()}


def catalog_with_inventory(session: Session) -> dict[str, Any]:
    catalog = latest_catalog_payload(session)
    equipment = catalog.get("equipment") if isinstance(catalog, dict) else []
    if not isinstance(equipment, list):
        equipment = []
    inventory = {item.part_key: item for item in session.scalars(select(FilterInventoryItem)).all()}
    inventory_details = inventory_catalog_details(equipment)
    inventory_list = [
        inventory_item_payload(item, inventory_details.get(item.part_key))
        for item in sorted(inventory.values(), key=lambda row: row.part_number)
    ]
    matched = shortages = unknown = 0
    for equipment_item in equipment:
        if not isinstance(equipment_item, dict):
            continue
        filters = equipment_item.get("filters")
        if not isinstance(filters, list):
            equipment_item["filters"] = []
            continue
        for filter_item in filters:
            if not isinstance(filter_item, dict):
                continue
            keys = filter_match_keys(filter_item)
            stock = next((inventory[key] for key in keys if key in inventory), None)
            required = parse_float(filter_item.get("quantity"), 0)
            if stock is None:
                filter_item["available"] = None
                filter_item["shortage"] = 0
                filter_item["inventory_status"] = "Sin inventario"
                filter_item["stock_part"] = ""
                unknown += 1
                continue
            shortage = max(required - parse_float(stock.quantity, 0), 0)
            filter_item["available"] = stock.quantity
            filter_item["shortage"] = shortage
            filter_item["inventory_status"] = "Faltante" if shortage > 0 else "Disponible"
            filter_item["stock_part"] = stock.part_number
            filter_item["stock_description"] = stock.description or inventory_details.get(stock.part_key, {}).get("description", "")
            filter_item["location"] = stock.location
            matched += 1
            if shortage > 0:
                shortages += 1
    result = dict(catalog) if isinstance(catalog, dict) else {}
    result.update(
        {
            "ok": True,
            "source": "cloud-inventory",
            "generated_at": utc_now().isoformat(timespec="seconds"),
            "equipment": equipment,
            "inventory": inventory_list,
            "summary": {
                "equipment": len(equipment),
                "filters": sum(len(item.get("filters") or []) for item in equipment if isinstance(item, dict)),
                "inventory_parts": len(inventory_list),
                "matched_filters": matched,
                "shortage_filters": shortages,
                "unknown_filters": unknown,
            },
        }
    )
    return result


def require_api_key(x_mga_api_key: str | None = Header(default=None)) -> None:
    if os.getenv("MGA_REQUIRE_API_KEY", "").strip().lower() not in {"1", "true", "yes", "si"}:
        return
    allowed = {
        value
        for value in (
            os.getenv("MGA_API_KEY", "").strip(),
            os.getenv("MGA_ALMACEN_KEY", "").strip(),
            "MGA4lmacen",
        )
        if value
    }
    if not allowed:
        raise HTTPException(status_code=500, detail="MGA_API_KEY no configurada en Render.")
    if x_mga_api_key not in allowed:
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


def iso_date(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return utc_now().date().isoformat()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text[:20]


def next_requisition_folio(session: Session) -> str:
    rows = session.scalars(select(CloudRequisition.folio).where(CloudRequisition.folio.like("REQ-%"))).all()
    last = 0
    for folio in rows:
        try:
            last = max(last, int(str(folio).split("-")[-1]))
        except ValueError:
            pass
    return f"REQ-{last + 1:04d}"


def requisition_payload(row: CloudRequisition, include_items: bool = False) -> dict[str, Any]:
    payload = {
        "id": row.id,
        "folio": row.folio,
        "request_date": row.request_date,
        "authorization_date": row.authorization_date,
        "equipment": row.equipment,
        "cost_center": row.cost_center,
        "request_area": row.request_area,
        "location": row.location,
        "requesting_unit": row.requesting_unit,
        "operating_unit": row.operating_unit,
        "priority": row.priority,
        "recommendation": row.recommendation,
        "status": row.status,
        "notes": row.notes,
        "purchase_status": row.purchase_status,
        "purchase_order": row.purchase_order,
        "purchase_order_date": row.purchase_order_date,
        "supplier": row.supplier,
        "buyer": row.buyer,
        "expected_date": row.expected_date,
        "received_date": row.received_date,
        "tracking_notes": row.tracking_notes,
        "tracking_source_file": row.tracking_source_file,
        "tracking_updated_at": row.tracking_updated_at.isoformat(timespec="seconds") if row.tracking_updated_at else "",
        "items_count": len([item for item in row.items if item.active]),
        "created_at": row.created_at.isoformat(timespec="seconds") if row.created_at else "",
        "updated_at": row.updated_at.isoformat(timespec="seconds") if row.updated_at else "",
    }
    if include_items:
        payload["items"] = [
            {
                "id": item.id,
                "quantity": item.quantity,
                "unit": item.unit,
                "part_number": item.part_number,
                "description": item.description,
                "sort_order": item.sort_order,
            }
            for item in sorted(row.items, key=lambda item: (item.sort_order, item.id))
            if item.active
        ]
    return payload


def tracking_equipment_code(value: Any) -> str:
    text = normalize_text(value)
    match = re.search(r"([A-Z]{1,4})[- ]?0*(\d{1,4})", text)
    if match:
        return f"{match.group(1)}-{int(match.group(2)):03d}"
    return text or "PARA STOCK"


def tracking_category_for_equipment(value: Any) -> str:
    code = tracking_equipment_code(value)
    text = normalize_text(value)
    if code.startswith("JL") or code.startswith("JA") or "JUMBO" in text:
        return "Jumbos:"
    if code.startswith("ST") or "SCOOP" in text:
        return "Scoop Tram"
    if code.startswith("RET") or "RETRO" in text:
        return "Retroexcavadoras"
    if code.startswith(("MG", "CBP")) or "CAMION" in text or "VEHIC" in text:
        return "Camiones y vehiculos"
    if "STOCK" in text or "TALLER" in text:
        return "Almacen"
    return "Otros"


def tracking_requisition_description(req: CloudRequisition, item: CloudRequisitionItem | None = None) -> str:
    parts = [req.folio]
    if item is not None:
        if item.part_number:
            parts.append(item.part_number)
        if item.description:
            parts.append(item.description)
    elif req.notes:
        parts.append(req.notes)
    return " ".join(str(part or "").strip() for part in parts if str(part or "").strip())


def tracking_comments(req: CloudRequisition) -> str:
    parts = []
    if req.purchase_status:
        parts.append(req.purchase_status)
    elif req.status:
        parts.append(req.status)
    if req.tracking_notes:
        parts.append(req.tracking_notes)
    return " / ".join(part for part in parts if part)


def infer_purchase_status(purchase_order: Any, notes: Any) -> str:
    if excel_cell_text(purchase_order):
        return "CON ORDEN DE COMPRA"
    text = normalize_text(notes)
    if not text:
        return ""
    if "CANCEL" in text:
        return "CANCELADO"
    if "RECIB" in text or "COLOCADO" in text:
        return "RECIBIDO"
    if "COTIZ" in text:
        return "COTIZANDO"
    if "ORDEN DE COMPRA" in text or "CON OC" in text:
        return "CON ORDEN DE COMPRA"
    return "EN SEGUIMIENTO"


def requisition_tracking_export_rows(session: Session) -> list[dict[str, Any]]:
    requisitions = session.scalars(
        select(CloudRequisition).order_by(CloudRequisition.equipment.asc(), CloudRequisition.folio.asc())
    ).all()
    rows: list[dict[str, Any]] = []
    for req in requisitions:
        active_items = [item for item in sorted(req.items, key=lambda item: (item.sort_order, item.id)) if item.active]
        items = active_items or [None]
        for item in items:
            rows.append(
                {
                    "category": tracking_category_for_equipment(req.equipment),
                    "equipment": tracking_equipment_code(req.equipment),
                    "description": tracking_requisition_description(req, item),
                    "purchase_order": req.purchase_order,
                    "delivery_time": req.expected_date or req.purchase_order_date or req.received_date,
                    "supplier": req.supplier,
                    "comments": tracking_comments(req),
                }
            )
    order = {"Jumbos:": 0, "Scoop Tram": 1, "Retroexcavadoras": 2, "Camiones y vehiculos": 3, "Almacen": 4, "Otros": 5}
    return sorted(rows, key=lambda row: (order.get(row["category"], 9), row["equipment"], row["description"]))


def requisition_tracking_week_title() -> str:
    month_names = [
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
    ]
    today = utc_now().date()
    start = today - timedelta(days=today.weekday())
    end = start + timedelta(days=6)
    if start.month == end.month:
        date_text = f"DEL {start.day:02d} AL {end.day:02d} DE {month_names[end.month - 1]}"
    else:
        date_text = f"DEL {start.day:02d} DE {month_names[start.month - 1]} AL {end.day:02d} DE {month_names[end.month - 1]}"
    return f"SEMANA {date_text} PROVIDENCIA"


def copy_cell_format(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)
    if source.protection:
        target.protection = copy(source.protection)


def prepare_tracking_export_sheet(ws, row_count: int) -> None:
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= 6:
            ws.unmerge_cells(str(merged))
    needed_rows = max(6 + max(row_count, 1) + 3, 40)
    if ws.max_row < needed_rows:
        ws.insert_rows(ws.max_row + 1, needed_rows - ws.max_row)
    for row_idx in range(6, needed_rows + 1):
        ws.row_dimensions[row_idx].height = ws.row_dimensions[6].height or 27
        for col_idx in range(1, 8):
            copy_cell_format(ws.cell(6, col_idx), ws.cell(row_idx, col_idx))
            ws.cell(row_idx, col_idx).value = None


def build_requisition_tracking_workbook(session: Session) -> bytes:
    if REQUISITION_TRACKING_TEMPLATE_PATH.exists():
        wb = load_workbook(REQUISITION_TRACKING_TEMPLATE_PATH)
        ws = wb.worksheets[-1]
        ws.title = "SEGUIMIENTO REQ"
        for sheet in list(wb.worksheets):
            if sheet is not ws:
                wb.remove(sheet)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "SEGUIMIENTO REQ"
        ws.append(["Equipos"])
        ws.append([])
        ws.append(["Descripcion", "No. Econ", requisition_tracking_week_title(), "", "", "", ""])
        ws.append(["", "", "Requisiciones", "Orden de Compra", "T.E", "Proveedor", "Comentarios Compras Y/O Proyecto"])
        ws.append(["Equipos:", "SEGUIMIENTO DE REQUISICIONES"])
    rows = requisition_tracking_export_rows(session)
    prepare_tracking_export_sheet(ws, len(rows))
    ws["C3"] = requisition_tracking_week_title()
    ws["A5"] = "Equipos:"
    ws["B5"] = "SEGUIMIENTO DE REQUISICIONES"

    if not rows:
        rows = [
            {
                "category": "Sin datos",
                "equipment": "",
                "description": "SIN REQUISICIONES PARA EXPORTAR",
                "purchase_order": "",
                "delivery_time": "",
                "supplier": "",
                "comments": "",
            }
        ]

    start_row = 6
    category_start = start_row
    equipment_start = start_row
    previous_category = rows[0]["category"]
    previous_equipment = rows[0]["equipment"]
    for offset, row in enumerate(rows):
        row_idx = start_row + offset
        category = row["category"]
        equipment = row["equipment"]
        if category != previous_category:
            if row_idx - category_start > 1:
                ws.merge_cells(start_row=category_start, start_column=1, end_row=row_idx - 1, end_column=1)
            previous_category = category
            category_start = row_idx
        if equipment != previous_equipment:
            if row_idx - equipment_start > 1:
                ws.merge_cells(start_row=equipment_start, start_column=2, end_row=row_idx - 1, end_column=2)
            previous_equipment = equipment
            equipment_start = row_idx
        ws.cell(row_idx, 1).value = category
        ws.cell(row_idx, 2).value = equipment
        ws.cell(row_idx, 3).value = row["description"]
        ws.cell(row_idx, 4).value = row["purchase_order"]
        ws.cell(row_idx, 5).value = row["delivery_time"]
        ws.cell(row_idx, 6).value = row["supplier"]
        ws.cell(row_idx, 7).value = row["comments"]
        for col_idx in range(1, 8):
            ws.cell(row_idx, col_idx).alignment = copy(ws.cell(row_idx, col_idx).alignment)
            ws.cell(row_idx, col_idx).alignment = Alignment(
                horizontal=ws.cell(row_idx, col_idx).alignment.horizontal or "center",
                vertical="center",
                wrap_text=True,
            )
    end_row = start_row + len(rows) - 1
    if end_row - category_start >= 1:
        ws.merge_cells(start_row=category_start, start_column=1, end_row=end_row, end_column=1)
    if end_row - equipment_start >= 1:
        ws.merge_cells(start_row=equipment_start, start_column=2, end_row=end_row, end_column=2)

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def text_width(c: pdf_canvas.Canvas, text: str, size: float, bold: bool = False) -> float:
    return c.stringWidth(str(text or ""), "Helvetica-Bold" if bold else "Helvetica", size)


def draw_text(c: pdf_canvas.Canvas, text: str, x: float, top: float, size: float = 7, bold: bool = False, align: str = "left") -> None:
    font = "Helvetica-Bold" if bold else "Helvetica"
    c.setFont(font, size)
    value = str(text or "")
    y = landscape(letter)[1] - top - size
    if align == "center":
        x -= text_width(c, value, size, bold) / 2
    elif align == "right":
        x -= text_width(c, value, size, bold)
    c.drawString(x, y, value)


def draw_box(c: pdf_canvas.Canvas, x: float, top: float, w: float, h: float, fill=None, stroke: int = 1) -> None:
    width, height = landscape(letter)
    if fill is not None:
        c.setFillColor(fill)
    else:
        c.setFillColor(colors.white)
    c.setStrokeColor(colors.black)
    c.rect(x, height - top - h, w, h, fill=1 if fill is not None else 0, stroke=stroke)


def draw_wrapped(c: pdf_canvas.Canvas, text: str, x: float, top: float, w: float, h: float, size: float = 6.2, bold: bool = False, align: str = "left") -> None:
    words = str(text or "").split()
    lines: list[str] = []
    line = ""
    for word in words:
        candidate = f"{line} {word}".strip()
        if text_width(c, candidate, size, bold) <= w or not line:
            line = candidate
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    max_lines = max(1, int(h // (size + 2)))
    for idx, line_text in enumerate(lines[:max_lines]):
        y_top = top + 3 + idx * (size + 2)
        if align == "center":
            draw_text(c, line_text, x + w / 2, y_top, size, bold, "center")
        else:
            draw_text(c, line_text, x, y_top, size, bold)


MONTH_ABBR_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


def requisition_pdf_parse_date(value: str | None) -> str | None:
    if value is None or value == "":
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def requisition_pdf_header_date(value: str | None) -> str:
    parsed = requisition_pdf_parse_date(value)
    if not parsed:
        return value or ""
    parsed_date = datetime.strptime(parsed, "%Y-%m-%d").date()
    return f"{parsed_date.day:02d}-{MONTH_ABBR_ES[parsed_date.month - 1]}-{str(parsed_date.year)[-2:]}"


def requisition_pdf_cell_date(value: str | None) -> str:
    parsed = requisition_pdf_parse_date(value)
    if not parsed:
        return value or ""
    parsed_date = datetime.strptime(parsed, "%Y-%m-%d").date()
    return f"{parsed_date.day:02d}/{parsed_date.month:02d}/{parsed_date.year}"


def requisition_pdf_template_bytes(row: CloudRequisition, items: list[CloudRequisitionItem], template_path: Path) -> bytes:
    if fitz is None:
        raise RuntimeError("PyMuPDF no esta disponible para usar la plantilla PDF.")
    template = fitz.open(str(template_path))
    output = fitz.open()
    rows_per_page = 14
    chunks = [items[idx:idx + rows_per_page] for idx in range(0, len(items), rows_per_page)] or [[]]

    def clear(page, rect) -> None:
        page.draw_rect(fitz.Rect(*rect), color=None, fill=(1, 1, 1), overlay=True)

    def text(page, rect, value: str, size: float = 7.2, align: int | None = None) -> None:
        page.insert_textbox(
            fitz.Rect(*rect),
            str(value or ""),
            fontsize=size,
            fontname="helv",
            color=(0, 0, 0),
            align=fitz.TEXT_ALIGN_CENTER if align is None else align,
            overlay=True,
        )

    def centered_line(page, rect, value: str, size: float = 7.0) -> None:
        value = str(value or "")
        if not value:
            return
        x0, y0, x1, y1 = rect
        while size > 4.8 and fitz.get_text_length(value, fontname="helv", fontsize=size) > (x1 - x0 - 2):
            size -= 0.3
        text_width_value = fitz.get_text_length(value, fontname="helv", fontsize=size)
        x = x0 + max((x1 - x0 - text_width_value) / 2, 0)
        y = y0 + ((y1 - y0 + size) / 2) - 1
        page.insert_text((x, y), value, fontsize=size, fontname="helv", color=(0, 0, 0), overlay=True)

    def line_text(page, rect, value: str, size: float = 7.0, align: int = 1) -> None:
        value = str(value or "")
        if not value:
            return
        x0, y0, x1, y1 = rect
        while size > 4.8 and fitz.get_text_length(value, fontname="helv", fontsize=size) > (x1 - x0 - 2):
            size -= 0.3
        text_width_value = fitz.get_text_length(value, fontname="helv", fontsize=size)
        if align == 0:
            x = x0
        elif align == 2:
            x = x1 - text_width_value
        else:
            x = x0 + max((x1 - x0 - text_width_value) / 2, 0)
        y = y0 + ((y1 - y0 + size) / 2) - 1
        page.insert_text((x, y), value, fontsize=size, fontname="helv", color=(0, 0, 0), overlay=True)

    def split_two_lines(value: str) -> tuple[str, str]:
        words = str(value or "").split()
        if len(words) <= 1:
            return str(value or ""), ""
        if len(words) == 2:
            return words[0], words[1]
        split_at = max(1, len(words) // 2)
        return " ".join(words[:split_at]), " ".join(words[split_at:])

    def two_line_value(page, rect_top, rect_bottom, value: str, size: float = 7.0) -> None:
        first, second = split_two_lines(value)
        line_text(page, rect_top, first, size)
        if second:
            line_text(page, rect_bottom, second, size)

    def clear_template_values(page) -> None:
        rects = [
            (452, 89, 482, 98),
            (141, 165, 180, 185),
            (312, 169, 366, 181),
            (462, 168, 503, 181),
            (558, 168, 600, 181),
            (145, 208, 196, 221),
            (395, 208, 434, 221),
            (102, 234, 141, 245),
            (326, 230, 365, 242),
            (529, 230, 582, 242),
            (29, 294, 38, 305),
            (69, 294, 86, 305),
            (342, 294, 455, 305),
            (29, 310, 38, 321),
            (69, 310, 86, 321),
            (340, 310, 488, 321),
        ]
        for rect in rects:
            clear(page, rect)

    try:
        for chunk in chunks:
            output.insert_pdf(template, from_page=0, to_page=0)
            page = output[-1]
            clear_template_values(page)

            line_text(page, (453, 91, 486, 97), requisition_pdf_header_date(requisition_field(row, "request_date")), 5.3, 0)
            two_line_value(page, (140, 166, 181, 174), (140, 175, 181, 183), requisition_field(row, "requesting_unit"), 7.2)
            line_text(page, (303, 168, 373, 183), requisition_field(row, "operating_unit"), 8.0)
            line_text(page, (456, 169, 509, 181), requisition_pdf_cell_date(requisition_field(row, "authorization_date")), 8.0)
            line_text(page, (562, 169, 607, 181), requisition_pdf_cell_date(requisition_field(row, "request_date")), 7.1)
            line_text(page, (95, 193, 593, 201), requisition_field(row, "cost_center"), 8.0)
            line_text(page, (107, 211, 238, 220), requisition_field(row, "equipment"), 8.2)
            line_text(page, (378, 211, 450, 220), requisition_field(row, "priority"), 8.2)
            line_text(page, (105, 236, 149, 243), requisition_field(row, "folio"), 7.1)
            line_text(page, (227, 235, 450, 243), requisition_field(row, "request_area"), 7.8)
            line_text(page, (514, 235, 593, 243), requisition_field(row, "location"), 7.6)

            notes = requisition_field(row, "notes")
            if notes:
                line_text(page, (80, 540, 594, 548), notes, 7.0, 0)
            else:
                clear(page, (80, 540, 594, 548))

            row_top = 292.0
            row_h = 16.2
            for idx in range(rows_per_page):
                y = row_top + idx * row_h
                clear(page, (18, y + 2.0, 49, y + 14.0))
                clear(page, (54, y + 2.0, 101, y + 14.0))
                clear(page, (106, y + 2.0, 237, y + 14.0))
                clear(page, (241, y + 2.0, 595, y + 14.0))
                if idx >= len(chunk):
                    continue
                item = chunk[idx]
                centered_line(page, (18, y + 1.4, 49, y + 15.0), fmt_qty(requisition_item_field(item, "quantity")), 7.0)
                centered_line(page, (54, y + 1.4, 101, y + 15.0), requisition_item_field(item, "unit"), 7.0)
                centered_line(page, (106, y + 1.4, 237, y + 15.0), requisition_item_field(item, "part_number"), 6.8)
                centered_line(page, (241, y + 1.2, 595, y + 15.2), requisition_item_field(item, "description"), 6.8)

        pdf = output.tobytes(garbage=4, deflate=True)
    finally:
        output.close()
        template.close()
    return pdf


def requisition_pdf_bytes(row: CloudRequisition, items: list[CloudRequisitionItem]) -> bytes:
    if fitz is not None and REQUISITION_TEMPLATE_PATH.exists():
        return requisition_pdf_template_bytes(row, items, REQUISITION_TEMPLATE_PATH)

    stream = BytesIO()
    width, height = landscape(letter)
    c = pdf_canvas.Canvas(stream, pagesize=landscape(letter))
    c.setLineWidth(0.75)
    chunks = [items[idx:idx + 8] for idx in range(0, len(items), 8)] or [[]]

    def cell(x: float, top: float, w: float, h: float, value: str = "", size: float = 6.2, bold: bool = False, fill=None, align: str = "left") -> None:
        draw_box(c, x, top, w, h, fill=fill)
        if value:
            if align == "center":
                draw_wrapped(c, value, x + 2, top + 2, w - 4, h - 4, size, bold, "center")
            else:
                draw_wrapped(c, value, x + 3, top + 3, w - 6, h - 6, size, bold)

    def check_option(label: str, selected: str, x: float, top: float, w: float) -> None:
        cell(x, top, w, 18, label, 5.8, align="center")
        marked = "X" if normalize_text(label) in normalize_text(selected) else ""
        cell(x + w, top, 18, 18, marked, 8, bold=True, align="center")

    def header(page_number: int, total_pages: int) -> None:
        c.setFillColor(colors.white)
        c.rect(0, 0, width, height, fill=1, stroke=0)
        logo = STATIC_DIR / "mga-corner-logo.jfif"
        if logo.exists():
            c.drawImage(str(logo), 28, height - 54, width=68, height=28, preserveAspectRatio=True, mask="auto")
        draw_text(c, "MGA CONTRATISTA MINERA S.A. DE C.V.", width / 2, 25, 10, bold=True, align="center")
        draw_text(c, "REQUISICION DE INSUMOS", width / 2, 78, 12, bold=True, align="center")
        draw_text(c, "Fecha:", 590, 25, 6.4, bold=True)
        draw_text(c, row.request_date, 640, 25, 7)
        draw_text(c, "Elaboro:", 590, 42, 6.4, bold=True)
        draw_text(c, "Aux De Compras", 640, 42, 7)
        draw_text(c, "Reviso:", 590, 59, 6.4, bold=True)
        draw_text(c, "Coord De Compras", 640, 59, 7)
        draw_text(c, "Aprobo:", 590, 76, 6.4, bold=True)
        draw_text(c, "Director Administrativo", 640, 76, 7)
        draw_text(c, "Codigo:", 590, 93, 6.4, bold=True)
        draw_text(c, "MGA_0016    Rev:00", 640, 93, 7)
        draw_text(c, f"Pagina {page_number} de {total_pages}", 760, 93, 6.3, align="right")

        top = 118
        cell(28, top, 230, 24, "UNIDAD OPERATIVA QUE SOLICITA EL INSUMO:", 5.8, bold=True)
        cell(258, top, 118, 24, row.requesting_unit, 6.4, align="center")
        cell(376, top, 92, 24, "UNIDAD OPERATIVA", 5.8, bold=True)
        cell(468, top, 95, 24, row.operating_unit, 6.4, align="center")
        cell(563, top, 100, 24, "FECHA DE AUTORIZACION:", 5.4, bold=True)
        cell(663, top, 96, 24, row.authorization_date, 6.2, align="center")
        cell(28, top + 30, 160, 24, "CENTRO DE COSTOS:", 5.8, bold=True)
        cell(188, top + 30, 185, 24, row.cost_center, 6.4, align="center")
        cell(373, top + 30, 116, 24, "FECHA SOLICITADA:", 5.8, bold=True)
        cell(489, top + 30, 100, 24, row.request_date, 6.4, align="center")
        cell(28, top + 60, 262, 24, "EQUIPO QUE REQUIERE LA PARTE O PIEZA SOLICITADA:", 5.5, bold=True)
        cell(290, top + 60, 160, 24, row.equipment, 7.2, bold=True, align="center")
        cell(470, top + 60, 165, 24, "PRIORIDAD EN SU ADQUISICION:", 5.5, bold=True)
        check_option("URGENTE", row.priority, 635, top + 60, 70)
        check_option("ORDINARIA", row.priority, 723, top + 60, 48)
        cell(28, top + 90, 125, 24, "FOLIO CONSECUTIVO:", 5.8, bold=True)
        cell(153, top + 90, 100, 24, row.folio, 7.2, bold=True, align="center")
        cell(253, top + 90, 110, 24, "AREA QUE SOLICITA:", 5.8, bold=True)
        cell(363, top + 90, 150, 24, row.request_area, 6.4, align="center")
        cell(513, top + 90, 80, 24, "UBICACION:", 5.8, bold=True)
        cell(593, top + 90, 178, 24, row.location, 6.4, align="center")
        cell(28, top + 120, 345, 24, "RECOMENDACION PARA LA OBTENCION DE LA PARTE O PIEZA SOLICITADA:", 5.4, bold=True)
        check_option("ORIGINAL", row.recommendation, 373, top + 120, 78)
        check_option("FABRICACION LOCAL", row.recommendation, 469, top + 120, 116)

    def draw_table(chunk: list[CloudRequisitionItem]) -> None:
        table_top = 278
        widths = [64, 112, 142, 445]
        headers = ["Cant.", "Unidad o Medida", "No de Parte", "Material Requerido (refaccion o insumos)"]
        x = 28
        cx = x
        for head, col_w in zip(headers, widths):
            cell(cx, table_top, col_w, 24, head, 6.4, bold=True, fill=colors.HexColor("#E5E7EB"), align="center")
            cx += col_w
        row_h = 28
        for offset in range(8):
            item = chunk[offset] if offset < len(chunk) else None
            top = table_top + 24 + offset * row_h
            fill = colors.white if offset % 2 == 0 else colors.HexColor("#F8FAFC")
            values = ["", "", "", ""]
            if item:
                values = [fmt_qty(item.quantity), item.unit, item.part_number, item.description]
            cx = x
            for value, col_w in zip(values, widths):
                cell(cx, top, col_w, row_h, str(value), 6.1, fill=fill, align="center" if col_w < 150 else "left")
                cx += col_w
        if row.notes:
            cell(28, 532, 745, 28, f"Notas: {row.notes}", 6.3)
        footer_top = 570
        for sig_x, label in ((155, "SOLICITA"), (390, "REVISA"), (625, "AUTORIZA")):
            c.setStrokeColor(colors.black)
            c.line(sig_x - 80, height - footer_top, sig_x + 80, height - footer_top)
            draw_text(c, label, sig_x, footer_top + 8, 7, bold=True, align="center")

    total = len(chunks)
    for page_idx, chunk in enumerate(chunks, start=1):
        header(page_idx, total)
        draw_table(chunk)
        if page_idx < total:
            c.showPage()
    c.save()
    return stream.getvalue()


def kpi_format_key(group: str) -> str:
    text = normalize_text(group)
    if "ACEITE" in text:
        return "aceites"
    if "LLANTA" in text:
        return "llantas"
    if "BARRENACION" in text:
        return "barrenacion"
    if "REZAGADO" in text:
        return "rezagado"
    if "ACARREO" in text:
        return "acarreo"
    if "UTILITARIO" in text:
        return "utilitario"
    return ""


def kpi_format_pdf_path(group: str) -> tuple[str, Path]:
    key = kpi_format_key(group)
    if not key:
        raise HTTPException(status_code=400, detail="Formato KPI disponible solo para Barrenacion, Rezagado, Acarreo, Utilitario, Aceites y Llantas.")
    path = KPI_FORMAT_PDFS.get(key)
    if path is None or not path.exists():
        raise HTTPException(status_code=404, detail="Formato KPI no encontrado en el servidor.")
    return key, path


def kpi_format_image_bytes(path: Path) -> bytes:
    if fitz is None:
        raise HTTPException(status_code=500, detail="Renderizado de imagen PDF no disponible.")
    doc = fitz.open(path)
    try:
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(2.4, 2.4), alpha=False)
        return pix.tobytes("png")
    finally:
        doc.close()


MONTH_NAMES_ES_FULL = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
OIL_REPORT_COLUMNS = [
    {"label": "Motor 15W40", "key": "oil_motor_15w40"},
    {"label": "ISO 68", "key": "oil_hco_iso68"},
    {"label": "SAE 30", "key": "oil_trans_sae30"},
    {"label": "SAE 50", "key": "oil_sae50"},
    {"label": "85W140", "key": "oil_85w140"},
    {"label": "ALMO", "key": "almo_liters"},
    {"label": "Refrigerante", "key": "coolant_liters"},
    {"label": "VG100", "key": "oil_hyd_vg100"},
    {"label": "ATF", "key": "atf_liters"},
]
OIL_STOCK_FIELDS = [
    ("oil_motor_15w40", "15W40", "Aceite 15W40"),
    ("oil_hco_iso68", "HCO ISO 68", "Aceite HCO ISO 68"),
    ("oil_trans_sae30", "SAE 30", "Aceite SAE 30"),
    ("oil_85w140", "85W140", "Aceite 85W140"),
    ("almo_liters", "ALMO", "ALMO"),
    ("coolant_liters", "REFRIGERANTE", "Refrigerante"),
    ("oil_hyd_vg100", "VG100", "Hidraulico VG100"),
    ("atf_liters", "ATF", "ATF"),
]
PPT_EMU_PER_INCH = 914400
PPT_BLUE = "#0b2f6f"
PPT_TEAL = "#00a6a6"
PPT_RED = "#e11d48"
PPT_LIGHT = "#eef3f9"
PPT_LINE = "#cfd8e5"
PPT_TEXT = "#061a3b"
MGA_BLUE = PPT_BLUE
MGA_TEAL = PPT_TEAL
MGA_TEAL_DARK = "#008b8b"
MGA_RED = PPT_RED


def month_bounds(year: int, month: int) -> tuple[str, str]:
    if month < 1 or month > 12:
        raise HTTPException(status_code=400, detail="Mes invalido.")
    start = datetime(year, month, 1).date()
    if month == 12:
        end = datetime(year + 1, 1, 1).date() - timedelta(days=1)
    else:
        end = datetime(year, month + 1, 1).date() - timedelta(days=1)
    return start.isoformat(), end.isoformat()


def parse_report_date(value: str, field_name: str) -> date:
    text = str(value or "").strip()[:10]
    try:
        return date.fromisoformat(text)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"Fecha invalida en {field_name}.") from exc


def week_bounds(base: date) -> tuple[date, date]:
    start = base - timedelta(days=base.weekday())
    return start, start + timedelta(days=6)


def weekly_period_label(start: date, end: date) -> str:
    start_month = MONTH_NAMES_ES_FULL[start.month - 1]
    end_month = MONTH_NAMES_ES_FULL[end.month - 1]
    if start.year == end.year and start.month == end.month:
        return f"Semana del {start.day:02d} al {end.day:02d} de {start_month} {start.year}"
    if start.year == end.year:
        return f"Semana del {start.day:02d} de {start_month} al {end.day:02d} de {end_month} {start.year}"
    return f"Semana del {start.day:02d} de {start_month} {start.year} al {end.day:02d} de {end_month} {end.year}"


def normalized_ascii(value: Any) -> str:
    text = normalize_text(value)
    return "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")


def portal_date_in_range(value: Any, start: str, end: str) -> bool:
    text = str(value or "")[:10]
    return bool(text) and start <= text <= end


def portal_equipment_rows(portal: dict[str, Any]) -> list[dict[str, Any]]:
    rows = portal.get("equipment") if isinstance(portal, dict) else []
    return [row for row in rows if isinstance(row, dict) and (row.get("code") or row.get("equipment_code"))]


def equipment_keys_py(value: Any) -> list[str]:
    text = normalized_ascii(value)
    if not text:
        return []
    first = re.split(r"\s+-\s+|\s+\(|\s+", text)[0] if text else ""
    keys: set[str] = set()
    for candidate in {text, first}:
        raw = re.sub(r"[^A-Z0-9]", "", normalized_ascii(candidate))
        if not raw:
            continue
        keys.add(raw)
        keys.add(re.sub(r"([A-Z]+)0+(\d)", r"\1\2", raw))
    return [key for key in keys if key]


def kpi_unavailable_status(status: Any) -> bool:
    text = normalized_ascii(status)
    return any(token in text for token in ("NO DISPONIBLE", "FUERA", "NO DISP", "REPARACION", "MANTENIMIENTO"))


def kpi_status_from_condition_py(value: Any) -> str:
    text = normalized_ascii(value)
    if not text:
        return ""
    if "STAND" in text:
        return "Stand By"
    if kpi_unavailable_status(text):
        return "No Disponible"
    if "DISPONIBLE" in text:
        return "Disponible"
    if "OPERATIVA" in text:
        return "Operativa"
    return "No Disponible"


def availability_status_map_py(portal: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}
    rows = portal.get("availability") if isinstance(portal, dict) else []
    for row in rows if isinstance(rows, list) else []:
        if not isinstance(row, dict):
            continue
        status = kpi_status_from_condition_py(row.get("condition"))
        if not status:
            continue
        for key in [*equipment_keys_py(row.get("eco")), *equipment_keys_py(row.get("equipment"))]:
            result.setdefault(key, status)
    return result


def availability_status_for_equipment_py(eq: dict[str, Any], status_map: dict[str, str]) -> str:
    for key in [*equipment_keys_py(eq.get("code") or eq.get("equipment_code")), *equipment_keys_py(eq.get("description"))]:
        if key in status_map:
            return status_map[key]
    return ""


def group_matches_py(eq: dict[str, Any], group: str) -> bool:
    key = normalized_ascii(group or "Todos")
    code = normalized_ascii(eq.get("code") or eq.get("equipment_code"))
    text = normalized_ascii(f"{code} {eq.get('description') or ''} {eq.get('family') or ''}")
    if "TODOS" in key:
        return True
    if "BARRENACION" in key:
        return code.startswith(("JL", "JA")) or "JUMBO" in text or "BARREN" in text or "ANCLADOR" in text
    if "REZAGADO" in key:
        return code.startswith("ST") or any(token in text for token in ("SCOOP", "CATERPILLAR", "EPROC", "R1300", "R1600", "REZAG"))
    if "ACARREO" in key:
        return re.sub(r"[^A-Z0-9]", "", code) in {"CBP001", "MG044"}
    if "UTILITARIO" in key:
        return re.sub(r"[^A-Z0-9]", "", code) in {"RET009", "RET010"}
    return True


def kpi_total_label_py(group: str) -> str:
    text = normalized_ascii(group)
    if "BARRENACION" in text:
        return "Total Equipos de Barrenacion"
    if "REZAGADO" in text:
        return "Total Equipos de Rezagado"
    if "ACARREO" in text:
        return "Total Acarreo"
    if "UTILITARIO" in text:
        return "Total Equipo Utilitario"
    if "TODOS" in text:
        return "Total Todos los equipos"
    return "Total Equipos"


def kpi_required_component_py(code: Any) -> str:
    code_text = normalized_ascii(code)
    compact_code = re.sub(r"[^A-Z0-9]", "", code_text)
    if code_text.startswith(("JL", "JA")):
        return "ELECT"
    if code_text.startswith(("ST", "RET")) or compact_code in {"CBP001", "MG044"}:
        return "DIESEL"
    return ""


def kpi_capture_component_matches_py(equipment_code: Any, component_name: Any) -> bool:
    required = kpi_required_component_py(equipment_code)
    if not required:
        return True
    return required in normalized_ascii(component_name)


def monthly_kpi_metric(period: float, worked: float, mp: float, mc: float, stops: float, mission_hours: float = 12) -> dict[str, float]:
    available = max(period - mp - mc, 0)
    availability = max(min((available / period) * 100, 100), 0) if period > 0 else 0
    utilization = max(min((worked / available) * 100, 100), 0) if available > 0 else 0
    stop_count = max(stops, 0)
    if stop_count:
        tmef = worked / stop_count if worked else 0
        tmpr = mc / stop_count
        reliability = max(min(math.exp(-(mission_hours / tmef)) * 100, 100), 0) if tmef and mission_hours else 0
    else:
        tmef = worked
        tmpr = 0
        reliability = 100 if worked else 0
    return {
        "available": available,
        "availability": availability,
        "utilization": utilization,
        "tmef": tmef,
        "tmpr": tmpr,
        "reliability": reliability,
    }


def desktop_kpi_report(portal: dict[str, Any], group: str, start: str, end: str) -> dict[str, Any] | None:
    reports = portal.get("kpi_reports") if isinstance(portal.get("kpi_reports"), dict) else {}
    source = reports.get(group) if isinstance(reports.get(group), dict) else None
    if not source or source.get("start") != start or source.get("end") != end:
        return None

    rows: list[dict[str, Any]] = []
    for source_row in source.get("rows") if isinstance(source.get("rows"), list) else []:
        if not isinstance(source_row, dict):
            continue
        period = parse_float(source_row.get("period") if "period" in source_row else source_row.get("period_hours"), 0)
        mp = parse_float(source_row.get("mp") if "mp" in source_row else source_row.get("mp_hours"), 0)
        mc = parse_float(source_row.get("mc") if "mc" in source_row else source_row.get("mc_hours"), 0)
        worked = parse_float(source_row.get("worked") if "worked" in source_row else source_row.get("worked_hours"), 0)
        availability_text = str(source_row.get("availability_text") or source_row.get("availabilityText") or "").strip()
        utilization_text = str(source_row.get("utilization_text") or source_row.get("utilizationText") or "").strip()
        out = availability_text.upper() == "FUERA" or utilization_text.upper() == "FUERA"
        rows.append(
            {
                "code": source_row.get("code") or "",
                "description": source_row.get("description") or "",
                "family": source_row.get("family") or "",
                "status": source_row.get("status") or ("FUERA" if out else ""),
                "period": period,
                "worked": worked,
                "mp": mp,
                "mc": mc,
                "stops": parse_float(source_row.get("stops"), 0),
                "available": 0 if out else max(period - mp - mc, 0),
                "availability": parse_float(source_row.get("availability"), 0),
                "utilization": parse_float(source_row.get("utilization"), 0),
                "tmef": parse_float(source_row.get("tmef"), 0),
                "tmpr": parse_float(source_row.get("tmpr"), 0),
                "reliability": parse_float(source_row.get("reliability"), 0),
                "out": out,
                "availability_text": availability_text or ("FUERA" if out else f"{parse_float(source_row.get('availability'), 0):.1f}%"),
                "utilization_text": utilization_text or ("FUERA" if out else f"{parse_float(source_row.get('utilization'), 0):.1f}%"),
            }
        )

    return {
        "group": source.get("group") or group,
        "start": start,
        "end": end,
        "start_day": int(str(start)[-2:]) if start else 1,
        "end_day": int(str(end)[-2:]) if end else 31,
        "rows": rows,
        "totals": source.get("totals") if isinstance(source.get("totals"), dict) else {},
        "source": "desktop-kpi-report",
    }


def monthly_kpi_report(portal: dict[str, Any], group: str, start: str, end: str) -> dict[str, Any]:
    precomputed = desktop_kpi_report(portal, group, start, end)
    if precomputed is not None:
        return precomputed

    settings = portal.get("settings") if isinstance(portal.get("settings"), dict) else {}
    shift_hours = parse_float(settings.get("shift_hours"), 9) or 9
    daily_hours = shift_hours * (parse_float(settings.get("turns_per_day"), 2) or 2)
    mission_hours = parse_float(settings.get("reliability_mission_hours") or settings.get("mission_hours"), 12) or 12
    try:
        start_date = datetime.strptime(start, "%Y-%m-%d").date()
        end_date = datetime.strptime(end, "%Y-%m-%d").date()
        days = max((end_date - start_date).days + 1, 1)
    except ValueError:
        days = 1
    grouped: dict[str, dict[str, Any]] = {}
    status_map = availability_status_map_py(portal)
    for eq in portal_equipment_rows(portal):
        if not group_matches_py(eq, group):
            continue
        code = str(eq.get("code") or eq.get("equipment_code") or "").strip()
        availability_status = availability_status_for_equipment_py(eq, status_map)
        grouped[code] = {
            "code": code,
            "description": str(eq.get("description") or ""),
            "family": str(eq.get("family") or ""),
            "status": availability_status or str(eq.get("status") or "Disponible"),
            "availability_status": availability_status,
            "capture_status": "",
            "capture_status_order": "",
            "period": days * daily_hours,
            "worked": 0.0,
            "mp": 0.0,
            "mc": 0.0,
            "stops": 0.0,
            "unavailable_count": 0,
        }
    captures = portal.get("captures") if isinstance(portal.get("captures"), list) else []
    oil_sources = [capture for capture in captures if isinstance(capture, dict)]
    oil_sources.extend(
        preventive_execution_as_oil_capture(row)
        for row in preventive_execution_records(portal)
        if isinstance(row, dict)
        and normalize_text(row.get("status")) in PREVENTIVE_CLOSED_STATUSES
        and sum(parse_float(row.get(key), 0) for key, _part_number, _description in OIL_STOCK_FIELDS) > 0
    )
    for capture in oil_sources:
        if not isinstance(capture, dict) or not portal_date_in_range(capture.get("work_date"), start, end):
            continue
        code = str(capture.get("equipment_code") or capture.get("code") or "").strip()
        if code not in grouped:
            continue
        if not kpi_capture_component_matches_py(code, capture.get("component") or capture.get("component_name")):
            continue
        row = grouped[code]
        mp = parse_float(capture.get("mp_hours"), 0)
        mc = parse_float(capture.get("mc_hours"), 0)
        capture_status = str(capture.get("status") or "").strip()
        if capture_status:
            capture_order = f"{capture.get('work_date') or ''}-{str(capture.get('id') or '').zfill(10)}"
            if not row["capture_status_order"] or capture_order >= row["capture_status_order"]:
                row["capture_status"] = capture_status
                row["capture_status_order"] = capture_order
                if not row["availability_status"]:
                    row["status"] = capture_status
        if kpi_unavailable_status(capture_status):
            base = daily_hours if normalized_ascii(capture.get("shift")) == "GENERAL" else shift_hours
            mc += max(base - mp - mc, 0)
            row["unavailable_count"] += 1
            if not row["availability_status"]:
                row["status"] = capture_status or "FUERA"
        row["worked"] += parse_float(capture.get("worked_hours"), 0)
        row["mp"] += mp
        row["mc"] += mc
        row["stops"] += parse_float(capture.get("stops"), 0)

    rows: list[dict[str, Any]] = []
    for row in sorted(grouped.values(), key=lambda item: item["code"]):
        if row["availability_status"]:
            row["status"] = row["availability_status"]
        elif row["capture_status"]:
            row["status"] = row["capture_status"]
        out = row["worked"] <= 0 and (row["unavailable_count"] > 0 or kpi_unavailable_status(row["status"]))
        metric_values = {"available": 0, "availability": 0, "utilization": 0, "tmef": 0, "tmpr": 0, "reliability": 0} if out else monthly_kpi_metric(row["period"], row["worked"], row["mp"], row["mc"], row["stops"], mission_hours)
        row.update(metric_values)
        row["out"] = out
        row["availability_text"] = "FUERA" if out else f"{row['availability']:.1f}%"
        row["utilization_text"] = "FUERA" if out else f"{row['utilization']:.1f}%"
        rows.append(row)

    totals = {"period": 0.0, "worked": 0.0, "mp": 0.0, "mc": 0.0, "stops": 0.0, "available": 0.0}
    for row in rows:
        totals["period"] += row["period"]
        totals["worked"] += row["worked"]
        totals["mp"] += row["mp"]
        totals["mc"] += row["mc"]
        totals["stops"] += row["stops"]
        totals["available"] += row["available"]
    totals["availability"] = (totals["available"] / totals["period"] * 100) if totals["period"] else 0
    totals["utilization"] = (totals["worked"] / totals["available"] * 100) if totals["available"] else 0
    totals["tmef"] = (totals["worked"] / totals["stops"]) if totals["stops"] and totals["worked"] else (totals["worked"] if not totals["stops"] else 0)
    totals["tmpr"] = (totals["mc"] / totals["stops"]) if totals["stops"] else 0
    totals["reliability"] = (
        max(min(math.exp(-(mission_hours / totals["tmef"])) * 100, 100), 0)
        if totals["tmef"] and mission_hours
        else (100 if totals["worked"] and not totals["stops"] else 0)
    )
    return {
        "group": group,
        "start": start,
        "end": end,
        "start_day": int(str(start)[-2:]) if start else 1,
        "end_day": int(str(end)[-2:]) if end else 31,
        "rows": rows,
        "totals": totals,
    }


def kpi_sim_factor(value: Any, default: float = 100) -> float:
    return max(parse_float(value, default), 0) / 100


def simulate_monthly_kpi_report(report: dict[str, Any], scenario: dict[str, Any]) -> dict[str, Any]:
    mission_hours = parse_float(scenario.get("reliability_mission_hours"), 12) or 12
    period_factor = kpi_sim_factor(scenario.get("period_percent"), 100)
    worked_factor = kpi_sim_factor(scenario.get("worked_percent"), 100)
    mp_factor = kpi_sim_factor(scenario.get("mp_percent"), 100)
    mc_factor = kpi_sim_factor(scenario.get("mc_percent"), 100)
    stops_factor = kpi_sim_factor(scenario.get("stops_percent"), 100)
    rows: list[dict[str, Any]] = []
    for source in report.get("rows") if isinstance(report.get("rows"), list) else []:
        row = dict(source)
        row["period"] = max(parse_float(source.get("period"), 0) * period_factor, 0)
        row["worked"] = max(parse_float(source.get("worked"), 0) * worked_factor, 0)
        row["mp"] = max(parse_float(source.get("mp"), 0) * mp_factor, 0)
        row["mc"] = max(parse_float(source.get("mc"), 0) * mc_factor, 0)
        row["stops"] = max(round(parse_float(source.get("stops"), 0) * stops_factor), 0)
        out = row["worked"] <= 0 and kpi_unavailable_status(row.get("status"))
        metric_values = {"available": 0, "availability": 0, "utilization": 0, "tmef": 0, "tmpr": 0, "reliability": 0} if out else monthly_kpi_metric(row["period"], row["worked"], row["mp"], row["mc"], row["stops"], mission_hours)
        row.update(metric_values)
        row["out"] = out
        row["availability_text"] = "FUERA" if out else f"{row['availability']:.1f}%"
        row["utilization_text"] = "FUERA" if out else f"{row['utilization']:.1f}%"
        rows.append(row)
    totals = {"period": 0.0, "worked": 0.0, "mp": 0.0, "mc": 0.0, "stops": 0.0, "available": 0.0}
    for row in rows:
        totals["period"] += parse_float(row.get("period"), 0)
        totals["worked"] += parse_float(row.get("worked"), 0)
        totals["mp"] += parse_float(row.get("mp"), 0)
        totals["mc"] += parse_float(row.get("mc"), 0)
        totals["stops"] += parse_float(row.get("stops"), 0)
        totals["available"] += parse_float(row.get("available"), 0)
    totals["availability"] = (totals["available"] / totals["period"] * 100) if totals["period"] else 0
    totals["utilization"] = (totals["worked"] / totals["available"] * 100) if totals["available"] else 0
    totals["tmef"] = (totals["worked"] / totals["stops"]) if totals["stops"] and totals["worked"] else (totals["worked"] if not totals["stops"] else 0)
    totals["tmpr"] = (totals["mc"] / totals["stops"]) if totals["stops"] else 0
    totals["reliability"] = (
        max(min(math.exp(-(mission_hours / totals["tmef"])) * 100, 100), 0)
        if totals["tmef"] and mission_hours
        else (100 if totals["worked"] and not totals["stops"] else 0)
    )
    simulated = dict(report)
    simulated["rows"] = rows
    simulated["totals"] = totals
    simulated["simulation"] = {
        "enabled": True,
        "name": str(scenario.get("name") or "Escenario KPI").strip(),
        "reliability_mission_hours": mission_hours,
    }
    return simulated


def kpi_row_obj(row: dict[str, Any]):
    return type(
        "KpiMonthlyRow",
        (),
        {
            "code": row.get("code") or "",
            "description": row.get("description") or "",
            "period_hours": parse_float(row.get("period"), 0),
            "mp_hours": parse_float(row.get("mp"), 0),
            "mc_hours": parse_float(row.get("mc"), 0),
            "worked_hours": parse_float(row.get("worked"), 0),
            "stops": int(parse_float(row.get("stops"), 0)),
            "availability": parse_float(row.get("availability"), 0),
            "utilization": parse_float(row.get("utilization"), 0),
            "tmef": parse_float(row.get("tmef"), 0),
            "tmpr": parse_float(row.get("tmpr"), 0),
            "reliability": parse_float(row.get("reliability"), parse_float(row.get("availability"), 0)),
            "status": "FUERA" if row.get("out") else str(row.get("status") or ""),
        },
    )()


def monthly_kpi_row_objects(report: dict[str, Any]) -> list[Any]:
    return [kpi_row_obj(row) for row in report.get("rows", [])]


def portal_oil_group(eq: dict[str, Any]) -> str:
    if group_matches_py(eq, "Equipos de Barrenacion"):
        return "BARRENACION"
    if group_matches_py(eq, "Equipos de Rezagado"):
        return "REZAGADO"
    return "UTILITARIO"


def portal_oil_report_for_period(portal: dict[str, Any], start: str, end: str) -> dict[str, Any]:
    settings = portal.get("settings") if isinstance(portal.get("settings"), dict) else {}
    shift_hours = parse_float(settings.get("shift_hours"), 9) or 9
    turns_per_day = parse_float(settings.get("turns_per_day"), 2) or 2
    daily_hours = shift_hours * turns_per_day
    try:
        start_date = datetime.strptime(start, "%Y-%m-%d").date()
        end_date = datetime.strptime(end, "%Y-%m-%d").date()
        days = max((end_date - start_date).days + 1, 1)
    except ValueError:
        days = 1
    grouped: dict[str, dict[str, Any]] = {}
    descriptions: dict[str, str] = {}
    for eq in portal_equipment_rows(portal):
        code = str(eq.get("code") or eq.get("equipment_code") or "").strip()
        if not code:
            continue
        descriptions[code.upper()] = str(eq.get("description") or "")
        grouped.setdefault(
            code,
            {
                "code": code,
                "description": str(eq.get("description") or ""),
                "group": portal_oil_group(eq),
                "period_hours": days * daily_hours,
                "period": days * daily_hours,
                "worked_hours": 0.0,
                "worked": 0.0,
                "total_liters": 0.0,
                **{col["key"]: 0.0 for col in OIL_REPORT_COLUMNS},
            },
        )

    captures = portal.get("captures") if isinstance(portal.get("captures"), list) else []
    for capture in captures:
        if not isinstance(capture, dict) or not portal_date_in_range(capture.get("work_date"), start, end):
            continue
        code = str(capture.get("equipment_code") or capture.get("code") or "").strip()
        if not code:
            continue
        if code not in grouped:
            grouped[code] = {
                "code": code,
                "description": descriptions.get(code.upper(), str(capture.get("equipment_description") or "")),
                "group": "UTILITARIO",
                "period_hours": days * daily_hours,
                "period": days * daily_hours,
                "worked_hours": 0.0,
                "worked": 0.0,
                "total_liters": 0.0,
                **{col["key"]: 0.0 for col in OIL_REPORT_COLUMNS},
            }
        row = grouped[code]
        worked = parse_float(capture.get("worked_hours"), 0)
        row["worked_hours"] += worked
        row["worked"] += worked
        for col in OIL_REPORT_COLUMNS:
            key = col["key"]
            value = parse_float(capture.get(key), 0)
            row[key] += value
            row["total_liters"] += value

    def sort_key(row: dict[str, Any]) -> tuple[int, str]:
        group_order = {"BARRENACION": 1, "REZAGADO": 2, "UTILITARIO": 3}
        return group_order.get(str(row.get("group") or ""), 9), str(row.get("code") or "")

    rows = sorted(grouped.values(), key=sort_key)
    totals: dict[str, float] = {
        "period": sum(parse_float(row.get("period"), 0) for row in rows),
        "period_hours": sum(parse_float(row.get("period_hours"), 0) for row in rows),
        "worked": sum(parse_float(row.get("worked"), 0) for row in rows),
        "worked_hours": sum(parse_float(row.get("worked_hours"), 0) for row in rows),
        "total_liters": sum(parse_float(row.get("total_liters"), 0) for row in rows),
    }
    for col in OIL_REPORT_COLUMNS:
        totals[col["key"]] = sum(parse_float(row.get(col["key"]), 0) for row in rows)
    return {
        "start": start,
        "end": end,
        "start_day": int(str(start)[-2:]) if start else 1,
        "end_day": int(str(end)[-2:]) if end else 31,
        "days": days,
        "daily_hours": daily_hours,
        "columns": OIL_REPORT_COLUMNS,
        "rows": rows,
        "totals": totals,
    }


def portal_for_report_period(session: Session, portal: dict[str, Any], start: str, end: str) -> dict[str, Any]:
    report_portal = dict(portal)
    report_portal["period"] = {
        "start": start,
        "end": end,
        "year": int(str(start)[:4]) if start else utc_now().year,
        "month": int(str(start)[5:7]) if len(str(start)) >= 7 else utc_now().month,
    }
    captures = report_portal.get("captures")
    merged = [row for row in captures if isinstance(row, dict)] if isinstance(captures, list) else []
    index_by_key = {capture_merge_key(row): idx for idx, row in enumerate(merged) if any(capture_merge_key(row))}
    equipment_rows = report_portal.get("equipment") if isinstance(report_portal.get("equipment"), list) else []
    descriptions = {
        str(e.get("code") or e.get("equipment_code") or "").strip().upper(): str(e.get("description") or e.get("family") or "")
        for e in equipment_rows
        if isinstance(e, dict)
    }
    for row in mobile_capture_portal_rows(session, limit=0, start=start, end=end):
        key = capture_merge_key(row)
        if not any(key):
            continue
        code = str(row.get("equipment_code") or "").strip().upper()
        if code and descriptions.get(code):
            row["equipment_description"] = descriptions[code]
        existing_index = index_by_key.get(key)
        if existing_index is None:
            merged.append(row)
            index_by_key[key] = len(merged) - 1
        else:
            merged[existing_index] = row
    merged.sort(key=lambda row: (str(row.get("work_date") or ""), int(parse_float(row.get("id"), 0))), reverse=True)
    report_portal["captures"] = merged
    report_portal["oil_kpi"] = portal_oil_report_for_period(report_portal, start, end)
    return report_portal


def is_kpi_out_row_py(row: Any) -> bool:
    return (parse_float(getattr(row, "worked_hours", 0), 0) or 0) <= 0 and kpi_unavailable_status(getattr(row, "status", ""))


def kpi_availability_text_py(row: Any, suffix: str = " %") -> str:
    return "FUERA" if is_kpi_out_row_py(row) else f"{row.availability:.1f}{suffix}"


def kpi_utilization_text_py(row: Any, suffix: str = " %") -> str:
    return "FUERA" if is_kpi_out_row_py(row) else f"{row.utilization:.1f}{suffix}"


def ppt_rgb(hex_color: str) -> RGBColor:
    clean = str(hex_color or "000000").strip().lstrip("#")
    if len(clean) != 6:
        clean = "000000"
    return RGBColor(int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16))


def ppt_slide_size(prs: Presentation) -> tuple[float, float]:
    return prs.slide_width / PPT_EMU_PER_INCH, prs.slide_height / PPT_EMU_PER_INCH


def ppt_clear_slide(slide) -> None:
    for shape in list(slide.shapes):
        parent = shape._element.getparent()
        if parent is not None:
            parent.remove(shape._element)


def ppt_blank_layout(prs: Presentation):
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]


def ppt_ensure_slide(prs: Presentation, index: int):
    while len(prs.slides) <= index:
        prs.slides.add_slide(ppt_blank_layout(prs))
    return prs.slides[index]


def ppt_rect(slide, x: float, y: float, w: float, h: float, fill: str = "ffffff", line: str | None = PPT_LINE):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = ppt_rgb(fill)
    if line:
        shape.line.color.rgb = ppt_rgb(line)
        shape.line.width = Pt(0.6)
    else:
        shape.line.fill.background()
    return shape


def ppt_text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    text: Any,
    size: float = 12,
    bold: bool = False,
    color: str = PPT_TEXT,
    align=PP_ALIGN.LEFT,
) -> None:
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = shape.text_frame
    frame.clear()
    frame.margin_left = Inches(0.03)
    frame.margin_right = Inches(0.03)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    lines = ("" if text is None else str(text)).splitlines() or [""]
    for idx, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = ppt_rgb(color)


def ppt_add_header(slide, prs: Presentation, title: str, subtitle: str = "") -> None:
    sw, _ = ppt_slide_size(prs)
    ppt_rect(slide, 0, 0, sw, 0.58, PPT_BLUE, None)
    if DIESEL_LOGO_PATH.exists():
        try:
            slide.shapes.add_picture(str(DIESEL_LOGO_PATH), Inches(0.18), Inches(0.12), width=Inches(0.7), height=Inches(0.34))
        except Exception:
            ppt_text(slide, 0.2, 0.14, 0.65, 0.3, "MGA", 12, True, "ffffff", PP_ALIGN.CENTER)
    else:
        ppt_text(slide, 0.2, 0.14, 0.65, 0.3, "MGA", 12, True, "ffffff", PP_ALIGN.CENTER)
    ppt_text(slide, 1.05, 0.1, sw - 1.3, 0.25, title, 19, True, "ffffff")
    if subtitle:
        ppt_text(slide, 1.05, 0.35, sw - 1.3, 0.18, subtitle, 8.5, False, "dbeafe")


def ppt_metric_progress(value: float, target: float, inverse: bool = False) -> float:
    value = parse_float(value, 0)
    target = max(parse_float(target, 1), 1)
    if inverse:
        return max(min((target / max(value, 0.1)) * 100, 100), 0)
    return max(min((value / target) * 100, 100), 0)


def ppt_metric_card(slide, x: float, y: float, w: float, h: float, label: str, value: str, note: str, progress: float, bad: bool = False) -> None:
    color = PPT_RED if bad else PPT_TEAL
    ppt_rect(slide, x, y, w, h, "ffffff", PPT_LINE)
    ppt_text(slide, x + 0.13, y + 0.09, w - 0.26, 0.16, label.upper(), 8.5, True, "52627a")
    ppt_text(slide, x + 0.13, y + 0.31, w - 0.26, 0.26, value, 19, True, "5d626a")
    ppt_text(slide, x + 0.13, y + h - 0.27, w - 0.26, 0.16, note, 7.5, False, "52627a")
    bar_w = w - 0.26
    ppt_rect(slide, x + 0.13, y + h - 0.12, bar_w, 0.05, "e5e7eb", None)
    ppt_rect(slide, x + 0.13, y + h - 0.12, bar_w * max(min(progress, 100), 0) / 100, 0.05, color, None)


def ppt_cell_text(cell, value: Any, size: float = 7, bold: bool = False, color: str = PPT_TEXT, fill: str = "ffffff", align=PP_ALIGN.CENTER) -> None:
    cell.text = "" if value is None else str(value)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.fill.solid()
    cell.fill.fore_color.rgb = ppt_rgb(fill)
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = align
        for run in paragraph.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = ppt_rgb(color)


def ppt_add_table(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    headers: list[str],
    rows: list[list[Any]],
    weights: list[float] | None = None,
    font_size: float = 6.4,
    left_cols: set[int] | None = None,
) -> None:
    left_cols = left_cols or set()
    table_rows = max(len(rows) + 1, 2)
    shape = slide.shapes.add_table(table_rows, len(headers), Inches(x), Inches(y), Inches(w), Inches(h))
    table = shape.table
    weights = weights or [1 for _ in headers]
    total_weight = sum(weights) or 1
    total_width = Inches(w)
    for idx, weight in enumerate(weights[:len(headers)]):
        table.columns[idx].width = int(total_width * weight / total_weight)
    row_height = int(Inches(h) / table_rows)
    for ridx in range(table_rows):
        table.rows[ridx].height = row_height
    for cidx, header in enumerate(headers):
        ppt_cell_text(table.cell(0, cidx), header, font_size, True, "ffffff", PPT_BLUE, PP_ALIGN.CENTER)
    for ridx in range(1, table_rows):
        values = rows[ridx - 1] if ridx - 1 < len(rows) else ["" for _ in headers]
        fill = "ffffff" if ridx % 2 else PPT_LIGHT
        for cidx in range(len(headers)):
            value = values[cidx] if cidx < len(values) else ""
            ppt_cell_text(
                table.cell(ridx, cidx),
                value,
                font_size,
                False,
                PPT_TEXT,
                fill,
                PP_ALIGN.LEFT if cidx in left_cols else PP_ALIGN.CENTER,
            )


def ppt_format_number(value: Any, decimals: int = 1) -> str:
    return f"{parse_float(value, 0):.{decimals}f}"


def ppt_format_pct(value: Any) -> str:
    return f"{parse_float(value, 0):.1f}%"


def ppt_add_bar_chart(slide, x: float, y: float, w: float, h: float, rows: list[dict[str, Any]], metric: str = "availability", target: float = 85) -> None:
    ppt_rect(slide, x, y, w, h, "ffffff", PPT_LINE)
    metric_label = {"availability": "% Disponibilidad", "utilization": "% Utilizacion", "reliability": "Confiabilidad", "tmef": "TMEF", "tmpr": "TMPR"}.get(metric, "% Disponibilidad")
    ppt_text(slide, x + 0.12, y + 0.09, w - 0.24, 0.2, metric_label, 10, True, PPT_TEXT)
    display = rows[:8]
    values = [parse_float(row.get(metric), 0) for row in display]
    axis_max = 120 if metric in {"availability", "utilization", "reliability"} else max([target, *values, 1]) * 1.18
    plot_x, plot_y = x + 0.35, y + 0.5
    plot_w, plot_h = w - 0.65, h - 0.9
    for idx in range(6):
        yy = plot_y + plot_h * idx / 5
        ppt_rect(slide, plot_x, yy, plot_w, 0.006, "dbe3ef", None)
    if not display:
        ppt_text(slide, x + 0.2, y + h / 2 - 0.1, w - 0.4, 0.2, "Sin datos KPI", 11, False, "64748b", PP_ALIGN.CENTER)
        return
    gap = 0.16
    bar_w = max((plot_w - gap * (len(display) + 1)) / len(display), 0.24)
    for idx, row in enumerate(display):
        value = parse_float(row.get(metric), 0)
        label = "FUERA" if row.get("out") and metric in {"availability", "utilization", "reliability"} else ppt_format_number(value)
        bh = max(min(value / max(axis_max, 1), 1) * plot_h, 0.04)
        bx = plot_x + gap + idx * (bar_w + gap)
        by = plot_y + plot_h - bh
        ppt_rect(slide, bx, by, bar_w, bh, PPT_TEAL, None)
        ppt_text(slide, bx - 0.05, by - 0.22, bar_w + 0.1, 0.15, label, 6.3, False, PPT_TEXT, PP_ALIGN.CENTER)
        ppt_text(slide, bx - 0.08, plot_y + plot_h + 0.07, bar_w + 0.16, 0.16, row.get("code") or "-", 6.2, True, PPT_TEXT, PP_ALIGN.CENTER)
    if metric in {"availability", "utilization", "reliability"}:
        target_y = plot_y + plot_h - min(target / 120, 1) * plot_h
        ppt_rect(slide, plot_x, target_y, plot_w, 0.012, PPT_RED, None)


def monthly_machine_slide(slide, prs: Presentation, portal: dict[str, Any], group: str, start: str, end: str, month_name: str, year: int) -> None:
    ppt_clear_slide(slide)
    report = monthly_kpi_report(portal, group, start, end)
    settings = portal.get("settings") if isinstance(portal.get("settings"), dict) else {}
    target_availability = parse_float(settings.get("meta_availability"), 85) or 85
    target_utilization = parse_float(settings.get("meta_utilization"), 75) or 75
    target_reliability = parse_float(settings.get("meta_reliability"), 80) or 80
    target_tmef = parse_float(settings.get("meta_tmef"), 8) or 8
    target_tmpr = parse_float(settings.get("meta_tmpr"), 4) or 4
    totals = report["totals"]
    sw, _ = ppt_slide_size(prs)
    ppt_add_header(slide, prs, group, f"{month_name} {year} | {start} a {end}")
    card_w = (sw - 0.7 - 0.18 * 4) / 5
    card_y = 0.82
    cards = [
        ("% Disponibilidad", ppt_format_pct(totals["availability"]), f"Meta {ppt_format_pct(target_availability)}", ppt_metric_progress(totals["availability"], target_availability), totals["availability"] < target_availability),
        ("% Utilizacion", ppt_format_pct(totals["utilization"]), f"Meta {ppt_format_pct(target_utilization)}", ppt_metric_progress(totals["utilization"], target_utilization), totals["utilization"] < target_utilization),
        ("Confiabilidad", ppt_format_pct(totals["reliability"]), f"Meta {ppt_format_pct(target_reliability)}", ppt_metric_progress(totals["reliability"], target_reliability), totals["reliability"] < target_reliability),
        ("TMEF", f"{ppt_format_number(totals['tmef'])} h", f"Meta {ppt_format_number(target_tmef)} h", ppt_metric_progress(totals["tmef"], target_tmef), totals["tmef"] < target_tmef),
        ("TMPR", f"{ppt_format_number(totals['tmpr'])} h", f"Meta {ppt_format_number(target_tmpr)} h", ppt_metric_progress(totals["tmpr"], target_tmpr, True), totals["tmpr"] > target_tmpr),
    ]
    for idx, (label, value, note, progress, bad) in enumerate(cards):
        ppt_metric_card(slide, 0.35 + idx * (card_w + 0.18), card_y, card_w, 0.78, label, value, note, progress, bad)
    ppt_add_bar_chart(slide, 0.35, 1.82, sw - 0.7, 2.18, report["rows"], "availability", target_availability)
    display_rows = [
        [
            row.get("code"),
            row.get("description"),
            ppt_format_number(row.get("period")),
            ppt_format_number(row.get("mp")),
            ppt_format_number(row.get("mc")),
            ppt_format_number(row.get("worked")),
            ppt_format_number(row.get("stops"), 0),
            row.get("availability_text"),
            row.get("utilization_text"),
            ppt_format_pct(row.get("reliability")),
            ppt_format_number(row.get("tmef")),
            ppt_format_number(row.get("tmpr")),
            "FUERA" if row.get("out") else row.get("status"),
        ]
        for row in report["rows"][:10]
    ]
    display_rows.append([
        "",
        f"Total {group}",
        ppt_format_number(totals["period"]),
        ppt_format_number(totals["mp"]),
        ppt_format_number(totals["mc"]),
        ppt_format_number(totals["worked"]),
        ppt_format_number(totals["stops"], 0),
        ppt_format_pct(totals["availability"]),
        ppt_format_pct(totals["utilization"]),
        ppt_format_pct(totals["reliability"]),
        ppt_format_number(totals["tmef"]),
        ppt_format_number(totals["tmpr"]),
        "",
    ])
    ppt_text(slide, 0.35, 4.18, sw - 0.7, 0.18, "REPORTE MENSUAL DE INDICADORES", 10, True, PPT_TEXT, PP_ALIGN.CENTER)
    headers = ["# Eco", "Equipo", "Hrs Periodo", "Hrs MP", "Hrs MC", "Hrs Trab", "# Paradas", "% Disp", "% Util", "Confiabilidad", "TMEF", "TMPR", "Estatus"]
    weights = [0.5, 1.65, 0.76, 0.62, 0.62, 0.68, 0.62, 0.62, 0.62, 0.82, 0.58, 0.58, 0.9]
    ppt_add_table(slide, 0.32, 4.45, sw - 0.64, 2.55, headers, display_rows, weights, 5.4, {1})


def monthly_tires_slide(slide, prs: Presentation, portal: dict[str, Any], month_name: str, year: int) -> None:
    ppt_clear_slide(slide)
    tire = portal.get("tire_kpi") if isinstance(portal.get("tire_kpi"), dict) else {}
    rows = tire.get("rows") if isinstance(tire.get("rows"), list) else []
    summary = tire.get("summary") if isinstance(tire.get("summary"), dict) else {}
    sw, _ = ppt_slide_size(prs)
    ppt_add_header(slide, prs, "Vida util de llantas", f"{month_name} {year}")
    cards = [
        ("Llantas", str(int(parse_float(summary.get("total"), len(rows)))), ""),
        ("Vida prom.", ppt_format_pct(summary.get("avg_life")), ""),
        ("Criticas", str(int(parse_float(summary.get("critical"), 0))), ""),
        ("Proximas", str(int(parse_float(summary.get("soon"), 0))), ""),
        ("Hrs rest. prom.", ppt_format_number(summary.get("avg_remaining_hours")), ""),
    ]
    card_w = (sw - 0.7 - 0.16 * 4) / 5
    for idx, (label, value, note) in enumerate(cards):
        ppt_metric_card(slide, 0.35 + idx * (card_w + 0.16), 0.82, card_w, 0.72, label, value, note, 100, False)
    table_rows = []
    for row in rows[:18]:
        life = parse_float(row.get("life_percent") if row.get("life_percent") is not None else row.get("tread_remaining_percent"), 0)
        table_rows.append([
            row.get("equipment_code"),
            row.get("tire_code"),
            row.get("position"),
            ppt_format_number(row.get("hours_used")),
            ppt_format_number(row.get("life_remaining_hours")),
            ppt_format_pct(life),
            ppt_format_pct(life),
            row.get("control_status") or "S/D",
        ])
    headers = ["Equipo", "Llanta", "Pos.", "Hrs uso", "Hrs rest.", "% vida", "% piso", "KPI"]
    weights = [0.9, 1.5, 0.5, 0.8, 0.8, 0.7, 0.7, 0.8]
    ppt_add_table(slide, 0.35, 1.78, sw - 0.7, 5.35, headers, table_rows, weights, 7.0, {1})


def monthly_oil_slide(slide, prs: Presentation, portal: dict[str, Any], month_name: str, year: int) -> None:
    ppt_clear_slide(slide)
    oil = portal.get("oil_kpi") if isinstance(portal.get("oil_kpi"), dict) else {}
    rows = oil.get("rows") if isinstance(oil.get("rows"), list) else []
    totals = oil.get("totals") if isinstance(oil.get("totals"), dict) else {}
    sw, _ = ppt_slide_size(prs)
    ppt_add_header(slide, prs, "KPI aceites", f"{month_name} {year}")
    metric_keys = [
        ("Motor 15W40", "oil_motor_15w40"),
        ("HCO ISO 68", "oil_hco_iso68"),
        ("Trans. SAE 30", "oil_trans_sae30"),
        ("SAE 50", "oil_sae50"),
        ("Total", "total_liters"),
    ]
    if "total_liters" not in totals:
        totals["total_liters"] = sum(parse_float(totals.get(key), 0) for _, key in metric_keys[:-1])
    card_w = (sw - 0.7 - 0.16 * 4) / 5
    for idx, (label, key) in enumerate(metric_keys):
        ppt_metric_card(slide, 0.35 + idx * (card_w + 0.16), 0.82, card_w, 0.72, label, f"{ppt_format_number(totals.get(key))} L", "Consumo", 100, False)
    sorted_rows = sorted(rows, key=lambda row: parse_float(row.get("total_liters"), 0), reverse=True)
    table_rows = []
    for row in sorted_rows[:15]:
        total = parse_float(row.get("total_liters"), 0)
        if not total:
            total = sum(parse_float(row.get(key), 0) for _, key in metric_keys[:-1])
        table_rows.append([
            row.get("code") or row.get("equipment_code"),
            row.get("description"),
            ppt_format_number(row.get("worked_hours") or row.get("worked")),
            ppt_format_number(row.get("oil_motor_15w40")),
            ppt_format_number(row.get("oil_hco_iso68")),
            ppt_format_number(row.get("oil_trans_sae30")),
            ppt_format_number(row.get("oil_sae50")),
            ppt_format_number(total),
        ])
    headers = ["Equipo", "Descripcion", "Hrs", "15W40", "ISO 68", "SAE 30", "SAE 50", "Total L"]
    weights = [0.8, 2.2, 0.65, 0.65, 0.65, 0.65, 0.65, 0.75]
    ppt_add_table(slide, 0.35, 1.78, sw - 0.7, 5.35, headers, table_rows, weights, 6.7, {1})


def monthly_diesel_slide(slide, prs: Presentation, portal: dict[str, Any], month_name: str, year: int) -> None:
    ppt_clear_slide(slide)
    diesel = portal.get("diesel") if isinstance(portal.get("diesel"), dict) else {}
    rows = diesel.get("rows") if isinstance(diesel.get("rows"), list) else []
    totals = diesel.get("totals") if isinstance(diesel.get("totals"), dict) else {}
    sw, _ = ppt_slide_size(prs)
    ppt_add_header(slide, prs, "KPI diesel", f"{month_name} {year}")
    cards = [
        ("Consumo", f"{ppt_format_number(totals.get('diesel_liters'))} L", ""),
        ("Horas", f"{ppt_format_number(totals.get('worked_hours'))} h", ""),
        ("Rendimiento", f"{ppt_format_number(totals.get('rendimiento_lh'))} L/H", ""),
        ("MGA disponible", f"{ppt_format_number(totals.get('mga_stock'))} L", ""),
        ("PROSERMIN", f"{ppt_format_number(totals.get('prosermin_stock'))} L", ""),
    ]
    card_w = (sw - 0.7 - 0.16 * 4) / 5
    for idx, (label, value, note) in enumerate(cards):
        ppt_metric_card(slide, 0.35 + idx * (card_w + 0.16), 0.82, card_w, 0.72, label, value, note, 100, False)
    table_rows = []
    source_rows = rows if rows else diesel.get("records", [])
    for row in source_rows[:16]:
        hours = parse_float(row.get("worked_hours"), 0)
        liters = parse_float(row.get("diesel_liters"), 0)
        rendimiento = parse_float(row.get("rendimiento_lh"), 0) or (liters / hours if hours else 0)
        table_rows.append([
            row.get("equipment") or row.get("equipment_code"),
            row.get("condition") or "",
            ppt_format_number(row.get("horometer_initial")),
            ppt_format_number(row.get("horometer_final")),
            ppt_format_number(hours),
            ppt_format_number(liters),
            ppt_format_number(rendimiento),
        ])
    headers = ["Equipo", "Condicion", "HI", "HF", "Hrs", "Diesel L", "L/H"]
    weights = [1.2, 1.1, 0.7, 0.7, 0.7, 0.8, 0.7]
    ppt_add_table(slide, 0.35, 1.78, sw - 0.7, 5.35, headers, table_rows, weights, 7.0, {0})


def iter_pptx_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_pptx_shapes(shape.shapes)


def update_monthly_ppt_text(prs: Presentation, month_name: str, year: int) -> None:
    month_pattern = r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+"
    for slide in prs.slides:
        for shape in iter_pptx_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    text = re.sub(
                        rf"Reporte Mensual\s+{month_pattern}(?:\s+\d{{4}})?",
                        f"Reporte Mensual {month_name} {year}",
                        text,
                        flags=re.IGNORECASE,
                    )
                    text = re.sub(
                        rf"Mes de\s+{month_pattern}(?:\s+\d{{4}})?",
                        f"Mes de {month_name} {year}",
                        text,
                        flags=re.IGNORECASE,
                    )
                    for source in MONTH_NAMES_ES_FULL:
                        text = text.replace(source.upper(), month_name.upper()).replace(source, month_name)
                    run.text = text

def update_weekly_ppt_text(prs: Presentation, label: str, month_name: str, year: int) -> None:
    month_pattern = r"[A-Za-zÃÃ‰ÃÃ“ÃšÃœÃ‘Ã¡Ã©Ã­Ã³ÃºÃ¼Ã±]+"
    for slide in prs.slides:
        for shape in iter_pptx_shapes(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    text = re.sub(
                        rf"Reporte Mensual\s+{month_pattern}(?:\s+\d{{4}})?",
                        f"Reporte Semanal {label}",
                        text,
                        flags=re.IGNORECASE,
                    )
                    text = re.sub(
                        rf"Mes de\s+{month_pattern}(?:\s+\d{{4}})?",
                        label,
                        text,
                        flags=re.IGNORECASE,
                    )
                    text = re.sub(r"Reporte\s+Mensual", "Reporte Semanal", text, flags=re.IGNORECASE)
                    for source in MONTH_NAMES_ES_FULL:
                        text = text.replace(source.upper(), month_name.upper()).replace(source, month_name)
                    run.text = text


def pil_font(size: int, bold: bool = False):
    names = (
        ["arialbd.ttf", "DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
        if bold
        else ["arial.ttf", "DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
    )
    for name in names:
        try:
            return ImageFont.truetype(name, max(int(size), 6))
        except Exception:
            continue
    return ImageFont.load_default()


def pil_text_size(draw: ImageDraw.ImageDraw, text: Any, font) -> tuple[int, int]:
    box = draw.textbbox((0, 0), str(text or ""), font=font)
    return max(box[2] - box[0], 1), max(box[3] - box[1], 1)


def pil_center(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: Any, font, fill: str) -> None:
    w, h = pil_text_size(draw, text, font)
    draw.text((xy[0] - w / 2, xy[1] - h / 2), str(text or ""), font=font, fill=fill)


def pil_right(draw: ImageDraw.ImageDraw, xy: tuple[float, float], text: Any, font, fill: str) -> None:
    w, h = pil_text_size(draw, text, font)
    draw.text((xy[0] - w, xy[1]), str(text or ""), font=font, fill=fill)


def pil_truncate(draw: ImageDraw.ImageDraw, text: Any, font, max_width: int) -> str:
    clean = " ".join(str(text or "").split())
    if draw.textlength(clean, font=font) <= max_width:
        return clean
    while clean and draw.textlength(clean + "...", font=font) > max_width:
        clean = clean[:-1]
    return f"{clean}..." if clean else ""


def pil_metric_card(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    title: str,
    value: str,
    note: str,
    color: str,
    progress: float,
    value_font,
    title_font,
    note_font,
) -> None:
    x1, y1, x2, y2 = box
    draw.rectangle(box, fill="white", outline="#e4e7ec")
    pil_center(draw, ((x1 + x2) / 2, y1 + (y2 - y1) * 0.22), title.upper(), title_font, "#5d6676")
    pil_center(draw, ((x1 + x2) / 2, y1 + (y2 - y1) * 0.48), value, value_font, "#5d626a")
    bar_x1 = x1 + max(6, int((x2 - x1) * 0.07))
    bar_x2 = x2 - max(6, int((x2 - x1) * 0.07))
    bar_y = y1 + int((y2 - y1) * 0.70)
    bar_h = max(5, int((y2 - y1) * 0.08))
    draw.rectangle((bar_x1, bar_y, bar_x2, bar_y + bar_h), fill="#e8ecf2")
    fill_w = int((bar_x2 - bar_x1) * max(min(progress, 100), 0) / 100)
    draw.rectangle((bar_x1, bar_y, bar_x1 + fill_w, bar_y + bar_h), fill=color)
    draw.text((bar_x1, y2 - max(18, int((y2 - y1) * 0.14))), note, font=note_font, fill="#52627a")


def pil_monthly_metric_card(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    value: str,
    bar_color: str,
    target_text: str,
    font_value,
    font_small,
    fill_ratio: float = 0.74,
) -> None:
    x1, y1, x2, y2 = box
    draw.rectangle(box, fill="white", outline="#e4e7ec")
    pil_center(draw, ((x1 + x2) / 2, y1 + (y2 - y1) * 0.38), value, font_value, "#5d626a")
    bar_x1 = x1 + 8
    bar_x2 = x2 - 8
    bar_y = y1 + int((y2 - y1) * 0.63)
    draw.rectangle((bar_x1, bar_y, bar_x2, bar_y + 11), fill="#eef1f4")
    draw.rectangle((bar_x1, bar_y, bar_x1 + int((bar_x2 - bar_x1) * max(min(fill_ratio, 1), 0)), bar_y + 11), fill=bar_color)
    draw.text((x1 + 8, y2 - 22), "vs Meta", font=font_small, fill="#6b7280")
    pil_right(draw, (x2 - 8, y2 - 22), target_text, font_small, "#111111")


def monthly_image_metric_progress(value: float, target: float, inverse: bool = False) -> float:
    value = parse_float(value, 0)
    target = max(parse_float(target, 1), 1)
    if inverse:
        return min((target / max(value, 0.1)) * 100, 100)
    return min((value / target) * 100, 100)


def create_monthly_kpi_dashboard_image(portal: dict[str, Any], group: str, start: str, end: str, path: Path, size: tuple[int, int]) -> None:
    report = monthly_kpi_report(portal, group, start, end)
    rows = monthly_kpi_row_objects(report)
    totals = report["totals"]
    settings = portal.get("settings") if isinstance(portal.get("settings"), dict) else {}
    target_availability = parse_float(settings.get("meta_availability"), 85) or 85
    target_utilization = parse_float(settings.get("meta_utilization"), 75) or 75
    target_reliability = parse_float(settings.get("meta_reliability"), 80) or 80
    target_tmef = parse_float(settings.get("meta_tmef"), 8) or 8
    target_tmpr = parse_float(settings.get("meta_tmpr"), 4) or 4
    w, h = size
    img = Image.new("RGB", size, "#eeeeee")
    draw = ImageDraw.Draw(img)
    title_font = pil_font(max(16, int(h * 0.045)), True)
    section_font = pil_font(max(13, int(h * 0.035)), True)
    value_font = pil_font(max(21, int(h * 0.065)), True)
    small_font = pil_font(max(8, int(h * 0.018)))
    axis_font = pil_font(max(8, int(h * 0.022)), True)
    tab_font = pil_font(max(8, int(h * 0.022)))

    draw.rectangle((0, 0, w, max(28, int(h * 0.08))), fill="white")
    pil_center(draw, (w / 2, int(h * 0.045)), group, title_font, "#333333")
    side_w = int(w * 0.265)
    center_x = side_w + int(w * 0.025)
    right_x = w - side_w
    center_w = max(w - side_w * 2 - int(w * 0.05), int(w * 0.42))
    y_top = int(h * 0.095)
    section_gap = int(h * 0.28)
    card_h = int(h * 0.215)
    gap = 5
    card_w = (side_w - gap) // 2

    pil_center(draw, (side_w / 2, y_top + 11), "% Disponibilidad", section_font, "#7a7d82")
    pil_monthly_metric_card(draw, (0, y_top + 28, card_w, y_top + 28 + card_h), f"{totals['availability']:.1f} %", MGA_TEAL, f"Meta {target_availability:.1f}%", value_font, small_font, totals["availability"] / 100)
    pil_monthly_metric_card(draw, (card_w + gap, y_top + 28, side_w, y_top + 28 + card_h), f"{target_availability:.1f} %", MGA_TEAL, f"{totals['availability'] - target_availability:+.1f}%", value_font, small_font, target_availability / 100)

    y2 = y_top + section_gap
    pil_center(draw, (side_w / 2, y2 + 11), "% Utilizacion", section_font, "#7a7d82")
    pil_monthly_metric_card(draw, (0, y2 + 28, card_w, y2 + 28 + card_h), f"{totals['utilization']:.1f} %", "#d36b73", f"Meta {target_utilization:.1f}%", value_font, small_font, totals["utilization"] / 100)
    pil_monthly_metric_card(draw, (card_w + gap, y2 + 28, side_w, y2 + 28 + card_h), f"{target_utilization:.1f} %", "#d36b73", f"{totals['utilization'] - target_utilization:+.1f}%", value_font, small_font, target_utilization / 100)

    right_card_h = int(h * 0.155)
    right_gap_y = int(h * 0.205)
    pil_center(draw, (right_x + side_w / 2, y_top + 11), "Confiabilidad", section_font, "#7a7d82")
    pil_monthly_metric_card(draw, (right_x, y_top + 28, right_x + card_w, y_top + 28 + right_card_h), f"{totals['reliability']:.1f} %", MGA_TEAL, f"Meta {target_reliability:.1f}%", value_font, small_font, totals["reliability"] / 100)
    pil_monthly_metric_card(draw, (right_x + card_w + gap, y_top + 28, w, y_top + 28 + right_card_h), f"{target_reliability:.1f} %", MGA_TEAL, f"{totals['reliability'] - target_reliability:+.1f}%", value_font, small_font, target_reliability / 100)

    y_tmef = y_top + right_gap_y
    pil_center(draw, (right_x + side_w / 2, y_tmef + 11), "TMEF", section_font, "#7a7d82")
    pil_monthly_metric_card(draw, (right_x, y_tmef + 28, right_x + card_w, y_tmef + 28 + right_card_h), f"{totals['tmef']:.1f} hrs", MGA_TEAL, f"Meta {target_tmef:.1f} h", value_font, small_font, min(totals["tmef"] / max(target_tmef, 1), 1))
    pil_monthly_metric_card(draw, (right_x + card_w + gap, y_tmef + 28, w, y_tmef + 28 + right_card_h), f"{target_tmef:.1f} hrs", "#c00000", f"{totals['tmef'] - target_tmef:+.1f} h", value_font, small_font, 0.36)

    y_tmpr = y_top + right_gap_y * 2
    pil_center(draw, (right_x + side_w / 2, y_tmpr + 11), "TMPR", section_font, "#7a7d82")
    pil_monthly_metric_card(draw, (right_x, y_tmpr + 28, right_x + card_w, y_tmpr + 28 + right_card_h), f"{totals['tmpr']:.1f} hrs", "#d36b73", f"Meta {target_tmpr:.1f} h", value_font, small_font, min(totals["tmpr"] / max(target_tmpr, 1), 1))
    pil_monthly_metric_card(draw, (right_x + card_w + gap, y_tmpr + 28, w, y_tmpr + 28 + right_card_h), f"{target_tmpr:.1f} hrs", "#d36b73", f"{target_tmpr - totals['tmpr']:+.1f} h", value_font, small_font, 0.38)

    chart_x = center_x
    chart_y = int(h * 0.12)
    chart_w = min(center_w, right_x - center_x - int(w * 0.02))
    chart_h = int(h * 0.75)
    draw.rectangle((chart_x, chart_y, chart_x + chart_w, chart_y + chart_h), fill="white")
    draw.text((chart_x + 5, chart_y + 5), "KPI", font=small_font, fill="#333333")
    labels = ["% Disponibilidad", "% Utilizacion", "Confiabilidad", "TMEF", "TMPR"]
    tab_y = chart_y + 24
    tab_w = (chart_w - 20) // 5
    for idx, label in enumerate(labels):
        tx = chart_x + 5 + idx * (tab_w + 4)
        fill = MGA_TEAL if idx == 0 else "white"
        draw.rounded_rectangle((tx, tab_y, tx + tab_w, tab_y + 24), radius=3, fill=fill, outline="#222222", width=1)
        draw.text((tx + 7, tab_y + 6), label, font=tab_font, fill="#111111")
    plot_x = chart_x + int(chart_w * 0.08)
    plot_y = tab_y + 50
    plot_w = chart_w - int(chart_w * 0.14)
    plot_h = chart_h - 100
    for pct in range(0, 121, 20):
        yy = plot_y + plot_h - int(plot_h * pct / 120)
        draw.line((plot_x, yy, plot_x + plot_w, yy), fill="#d8dde3", width=1)
        draw.text((chart_x + 8, yy - 6), str(pct), font=small_font, fill="#4b5563")
    display = rows[:7]
    if display:
        bar_gap = max(8, int(plot_w * 0.035))
        bar_w = max(12, int((plot_w - bar_gap * (len(display) + 1)) / max(len(display), 1)))
        for idx, row in enumerate(display):
            value = max(min(row.availability, 120), 0)
            x = plot_x + bar_gap + idx * (bar_w + bar_gap)
            y = plot_y + plot_h - int(plot_h * value / 120)
            draw.rectangle((x, y, x + bar_w, plot_y + plot_h), fill=PPT_TEAL)
            pil_center(draw, (x + bar_w / 2, y - 12), "FUERA" if is_kpi_out_row_py(row) else f"{row.availability:.0f}", axis_font, "#5d626a")
            pil_center(draw, (x + bar_w / 2, plot_y + plot_h + 13), row.code or "-", small_font, "#5d626a")
            pil_center(draw, (x + bar_w / 2, plot_y + plot_h - 8), "0", axis_font, "#5d626a")
    target_y = plot_y + plot_h - int(plot_h * target_availability / 120)
    draw.line((plot_x, target_y, plot_x + plot_w, target_y), fill="#7f858c", width=1)
    img.save(path, quality=95)


def pil_table_cell(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], text: Any, font, fill: str, outline: str = "#111111", text_fill: str = "#555555", align: str = "center") -> None:
    draw.rectangle(box, fill=fill, outline=outline)
    x1, y1, x2, y2 = box
    clean = pil_truncate(draw, text, font, max(x2 - x1 - 6, 1))
    if align == "left":
        draw.text((x1 + 4, y1 + max(2, (y2 - y1 - pil_text_size(draw, clean, font)[1]) / 2)), clean, font=font, fill=text_fill)
    else:
        pil_center(draw, ((x1 + x2) / 2, (y1 + y2) / 2), clean, font, text_fill)


def create_monthly_kpi_table_image(portal: dict[str, Any], group: str, start: str, end: str, path: Path, size: tuple[int, int]) -> None:
    report = monthly_kpi_report(portal, group, start, end)
    rows = monthly_kpi_row_objects(report)
    totals = report["totals"]
    w, h = size
    img = Image.new("RGB", size, "#eeeeee")
    draw = ImageDraw.Draw(img)
    title_font = pil_font(max(13, int(h * 0.055)), True)
    label_font = pil_font(max(9, int(h * 0.037)), True)
    table_font = pil_font(max(8, int(h * 0.031)), True)
    small_font = pil_font(max(7, int(h * 0.026)))

    pil_center(draw, (w / 2, int(h * 0.07)), "REPORTE MENSUAL DE INDICADORES", title_font, "#111111")
    draw.text((int(w * 0.36), int(h * 0.13)), "Dia Inicial:", font=label_font, fill="#777777")
    draw.rectangle((int(w * 0.44), int(h * 0.115), int(w * 0.50), int(h * 0.175)), fill="white")
    pil_center(draw, (int(w * 0.47), int(h * 0.145)), str(report["start_day"]), title_font, "#111111")
    draw.text((int(w * 0.62), int(h * 0.13)), "Dia Final:", font=label_font, fill="#777777")
    draw.rectangle((int(w * 0.70), int(h * 0.115), int(w * 0.76), int(h * 0.175)), fill="white")
    pil_center(draw, (int(w * 0.73), int(h * 0.145)), str(report["end_day"]), title_font, "#111111")

    table_top = int(h * 0.22)
    headers = ["# Eco", "Equipo", "Hrs Periodo", "Hrs MP", "Hrs MC", "Hrs Trab", "# Paradas", "% Disp", "% Util", "Confiabilidad", "TMEF", "TMPR", "Estatus"]
    weights = [0.55, 1.95, 0.8, 0.75, 0.75, 0.75, 0.8, 0.75, 0.75, 0.95, 0.75, 0.75, 1.25]
    total_weight = sum(weights)
    widths = [int(w * weight / total_weight) for weight in weights]
    widths[-1] += w - sum(widths)
    display_rows = rows[:8]
    total_label = kpi_total_label_py(group)
    total_row = type(
        "KpiMonthlyRow",
        (),
        {
            "code": "",
            "description": total_label,
            "period_hours": totals["period"],
            "mp_hours": totals["mp"],
            "mc_hours": totals["mc"],
            "worked_hours": totals["worked"],
            "stops": int(totals["stops"]),
            "availability": totals["availability"],
            "utilization": totals["utilization"],
            "tmef": totals["tmef"],
            "tmpr": totals["tmpr"],
            "reliability": totals.get("reliability", totals["availability"]),
            "status": "",
        },
    )()
    table_rows = display_rows + [total_row]
    row_count = len(table_rows) + 1
    row_h = max(16, int((h - table_top - 6) / max(row_count, 1)))
    x = 0
    for header, col_w in zip(headers, widths):
        pil_table_cell(draw, (x, table_top, x + col_w, table_top + row_h), header, table_font, "white", "#222222", "#666666")
        x += col_w
    for ridx, row in enumerate(table_rows, 1):
        is_total = ridx == len(table_rows)
        y = table_top + ridx * row_h
        fill = "#d9d9d9" if is_total else "#efefef"
        values = [
            row.code,
            row.description,
            f"{row.period_hours:.1f}" if not is_total else f"{row.period_hours:.2f}",
            f"{row.mp_hours:.1f}" if not is_total else f"{row.mp_hours:.2f}",
            f"{row.mc_hours:.1f}" if not is_total else f"{row.mc_hours:.2f}",
            f"{row.worked_hours:.1f}",
            f"{row.stops}",
            kpi_availability_text_py(row),
            kpi_utilization_text_py(row),
            f"{row.reliability:.0f}%",
            f"{row.tmef:.1f}",
            f"{row.tmpr:.1f}",
            row.status,
        ]
        x = 0
        for cidx, (value, col_w) in enumerate(zip(values, widths)):
            cell_fill = fill
            text_fill = "#666666"
            if value == "FUERA":
                text_fill = "#c76870"
            elif cidx in (7, 9):
                text_fill = MGA_TEAL
            elif cidx in (8, 11) and (is_total or row.utilization < 50):
                text_fill = "#c76870"
            elif cidx == 12 and value:
                text_fill = "#c76870" if kpi_unavailable_status(value) else MGA_TEAL
            pil_table_cell(draw, (x, y, x + col_w, y + row_h), value, table_font if cidx != 1 else small_font, cell_fill, "#222222", text_fill)
            x += col_w
    img.save(path, quality=95)


def create_monthly_tire_image(portal: dict[str, Any], path: Path, size: tuple[int, int], month_name: str, year: int) -> None:
    tire = portal.get("tire_kpi") if isinstance(portal.get("tire_kpi"), dict) else {}
    rows = tire.get("rows") if isinstance(tire.get("rows"), list) else []
    summary = tire.get("summary") if isinstance(tire.get("summary"), dict) else {}
    w, h = size
    img = Image.new("RGB", size, "#f5f8fc")
    draw = ImageDraw.Draw(img)
    title_text = f"VIDA UTIL DE LLANTAS - {month_name.upper()} {year}"
    title_size = max(12, int(h * 0.040))
    title_font = pil_font(title_size, True)
    stat_font = pil_font(max(12, int(h * 0.030)), True)
    label_font = pil_font(max(6, int(h * 0.016)), True)
    table_font = pil_font(max(7, int(h * 0.020)), True)
    small_font = pil_font(max(7, int(h * 0.020)))
    while title_size > 12 and draw.textlength(title_text, font=title_font) > w - max(16, int(w * 0.04)):
        title_size -= 1
        title_font = pil_font(title_size, True)
    header_h = int(h * 0.10)
    draw.rectangle((0, 0, w, header_h), fill=PPT_BLUE)
    pil_center(draw, (w / 2, int(header_h * 0.50)), title_text, title_font, "white")
    stats = [
        ("Llantas", summary.get("total", len(rows))),
        ("Vida prom.", f"{parse_float(summary.get('avg_life'), 0):.0f}%"),
        ("Criticas", summary.get("critical", 0)),
        ("Proximas", summary.get("soon", 0)),
        ("Hrs rest. prom.", f"{parse_float(summary.get('avg_remaining_hours'), 0):.0f}"),
    ]
    stat_y = header_h + max(8, int(h * 0.02))
    stat_gap = max(4, int(w * 0.01))
    stat_w = (w - stat_gap * 6) // 5
    stat_h = max(44, int(h * 0.11))
    for idx, (label, value) in enumerate(stats):
        x = stat_gap + idx * (stat_w + stat_gap)
        draw.rectangle((x, stat_y, x + stat_w, stat_y + stat_h), fill="white", outline="#d8dee8")
        pil_center(draw, (x + stat_w / 2, stat_y + stat_h * 0.42), value, stat_font, PPT_TEAL if idx == 1 else "#0b2f6f")
        pil_center(draw, (x + stat_w / 2, stat_y + stat_h * 0.72), label, label_font, "#667085")
    headers = ["Equipo", "Llanta", "Pos.", "Hrs uso", "Hrs rest.", "% vida", "% piso", "KPI"]
    weights = [0.85, 1.2, 0.45, 0.7, 0.8, 0.95, 0.65, 0.6]
    total_weight = sum(weights)
    widths = [int(w * weight / total_weight) for weight in weights]
    widths[-1] += w - sum(widths)
    table_top = stat_y + stat_h + max(8, int(h * 0.02))
    row_h = max(18, int((h - table_top - 4) / max(min(len(rows), 14) + 1, 2)))
    x = 0
    for header, col_w in zip(headers, widths):
        pil_table_cell(draw, (x, table_top, x + col_w, table_top + row_h), header, table_font, "#e8eef7", "#cdd5df", "#516173")
        x += col_w
    y = table_top + row_h
    for row in rows[:14]:
        life = max(min(parse_float(row.get("life_percent") if row.get("life_percent") is not None else row.get("tread_remaining_percent"), 0), 100), 0)
        values = [
            row.get("equipment_code"),
            row.get("tire_code"),
            row.get("position"),
            f"{parse_float(row.get('hours_used'), 0):.0f}",
            f"{parse_float(row.get('life_remaining_hours'), 0):.0f}",
            f"{life:.0f}%",
            f"{life:.0f}%",
            row.get("control_status") or "S/D",
        ]
        x = 0
        for idx, (value, col_w) in enumerate(zip(values, widths)):
            if idx == 5:
                draw.rectangle((x, y, x + col_w, y + row_h), fill="white", outline="#d8dee8")
                bx1 = x + 6
                by1 = y + max(5, (row_h - 8) // 2)
                bx2 = x + col_w - max(42, int(col_w * 0.34))
                bar_h = max(4, int(row_h * 0.16))
                draw.rounded_rectangle((bx1, by1, bx2, by1 + bar_h), radius=3, fill="#e5e7eb")
                fill_w = int((bx2 - bx1) * life / 100)
                if fill_w > 0:
                    draw.rounded_rectangle((bx1, by1, bx1 + fill_w, by1 + bar_h), radius=3, fill=PPT_TEAL)
                draw.rectangle((bx1, by1, bx2, by1 + bar_h), outline="#d1d5db")
                pil_right(draw, (x + col_w - 6, y + max(2, (row_h - pil_text_size(draw, values[5], small_font)[1]) / 2)), values[5], small_font, "#344054")
            else:
                pil_table_cell(draw, (x, y, x + col_w, y + row_h), value, small_font, "white", "#d8dee8", PPT_TEAL if idx in (6, 7) else "#5c6673")
            x += col_w
        y += row_h
    img.save(path, quality=95)


def create_monthly_diesel_image(portal: dict[str, Any], path: Path, size: tuple[int, int], start: str, end: str, month_name: str, year: int) -> None:
    diesel = portal.get("diesel") if isinstance(portal.get("diesel"), dict) else {}
    rows = diesel.get("rows") if isinstance(diesel.get("rows"), list) else []
    records = diesel.get("records") if isinstance(diesel.get("records"), list) else []
    totals = diesel.get("totals") if isinstance(diesel.get("totals"), dict) else {}
    meta = parse_float(diesel.get("meta_lh") or (portal.get("settings") or {}).get("meta_diesel_lh"), 25) or 25
    w, h = size
    img = Image.new("RGB", size, "#f4f7fb")
    draw = ImageDraw.Draw(img)
    subtitle_font = pil_font(max(13, int(w * 0.011)), True)
    card_label_font = pil_font(max(13, int(w * 0.010)), True)
    card_value_font = pil_font(max(26, int(w * 0.024)), True)
    small_font = pil_font(max(10, int(w * 0.008)))
    table_font = pil_font(max(10, int(w * 0.008)))
    table_bold = pil_font(max(10, int(w * 0.008)), True)

    def rendimiento_text(value: float | None) -> str:
        return "S/H" if value is None else f"{value:.1f}"

    def metric_card(box: tuple[int, int, int, int], label: str, value: str, note: str, progress: float, bad: bool = False) -> None:
        x1, y1, x2, y2 = box
        draw.rounded_rectangle(box, radius=8, fill="white", outline="#d5dde8")
        draw.text((x1 + 18, y1 + 15), label.upper(), font=card_label_font, fill="#52627a")
        draw.text((x1 + 18, y1 + 43), value, font=card_value_font, fill=MGA_BLUE)
        draw.text((x1 + 18, y2 - 32), note, font=small_font, fill="#65758b")
        bar_x1, bar_y = x1 + 18, y2 - 15
        bar_x2 = x2 - 18
        draw.rounded_rectangle((bar_x1, bar_y, bar_x2, bar_y + 8), radius=4, fill="#e5e9ef")
        fill_w = int((bar_x2 - bar_x1) * max(min(progress, 100), 0) / 100)
        draw.rounded_rectangle((bar_x1, bar_y, bar_x1 + fill_w, bar_y + 8), radius=4, fill=MGA_RED if bad else MGA_TEAL)

    def diesel_status(row: dict[str, Any], rendimiento: float | None) -> str:
        hours = parse_float(row.get("worked_hours"), 0)
        liters = parse_float(row.get("diesel_liters"), 0)
        if liters > 0 and hours <= 0:
            return "SIN HORAS"
        if rendimiento is not None and rendimiento > meta:
            return "ALTO"
        return "OK"

    source_rows = rows if rows else records
    normalized_rows: list[dict[str, Any]] = []
    for row in source_rows:
        hours = parse_float(row.get("worked_hours"), 0)
        liters = parse_float(row.get("diesel_liters"), 0)
        rendimiento = parse_float(row.get("rendimiento_lh"), 0) or (liters / hours if hours else None)
        normalized_rows.append({
            "equipment": row.get("equipment") or row.get("equipment_code") or "",
            "condition": row.get("condition") or "",
            "horometer_initial": parse_float(row.get("horometer_initial"), 0),
            "horometer_final": parse_float(row.get("horometer_final"), 0),
            "worked_hours": hours,
            "diesel_liters": liters,
            "rendimiento_lh": rendimiento,
            "status": diesel_status(row, rendimiento),
        })
    total_liters = parse_float(totals.get("diesel_liters"), 0) or 0
    mga_liters = parse_float(totals.get("mga_stock"), 0) or 0
    prosermin_liters = parse_float(totals.get("prosermin_stock"), 0) or 0
    total_hours = parse_float(totals.get("worked_hours"), 0) or 0
    avg_value = parse_float(totals.get("rendimiento_lh"), 0) or (total_liters / total_hours if total_hours else None)
    critical = sum(1 for row in normalized_rows if row["status"] in {"ALTO", "SIN HORAS"})
    active_rows = [row for row in normalized_rows if row["diesel_liters"] > 0 or row["worked_hours"] > 0]

    sx = w / 1056
    sy = h / 1374

    def rx(value: float) -> int:
        return int(value * sx)

    def ry(value: float) -> int:
        return int(value * sy)

    def rbox(box: tuple[float, float, float, float]) -> tuple[int, int, int, int]:
        return (rx(box[0]), ry(box[1]), rx(box[2]), ry(box[3]))

    def diesel_bad(row: dict[str, Any]) -> bool:
        return row["status"] in {"ALTO", "SIN HORAS"}

    def status_pill(box: tuple[int, int, int, int], text: str, bad: bool) -> None:
        x1, y1, x2, y2 = box
        fill = "#fde1e4" if bad else "#dff7ef"
        text_color = MGA_RED if bad else MGA_TEAL_DARK
        draw.rounded_rectangle((x1, y1, x2, y2), radius=max(5, ry(8)), fill=fill)
        pil_center(draw, ((x1 + x2) / 2, (y1 + y2) / 2), text, small_font, text_color)

    def card(box: tuple[float, float, float, float], label: str, value: str, note: str, progress: float, bad: bool = False) -> None:
        x1, y1, x2, y2 = rbox(box)
        draw.rounded_rectangle((x1, y1, x2, y2), radius=rx(7), fill="white", outline="#d5dde8")
        draw.rectangle((x1, y1 + rx(6), x1 + rx(4), y2 - rx(6)), fill=MGA_TEAL)
        draw.text((x1 + rx(14), y1 + ry(15)), label.upper(), font=card_label_font, fill="#52627a")
        draw.text((x1 + rx(14), y1 + ry(42)), value, font=card_value_font, fill=MGA_BLUE)
        draw.text((x1 + rx(14), y2 - ry(34)), note, font=small_font, fill="#65758b")
        bar_x1 = x1 + rx(14)
        bar_x2 = x2 - rx(14)
        bar_y = y2 - ry(16)
        draw.rounded_rectangle((bar_x1, bar_y, bar_x2, bar_y + ry(7)), radius=ry(4), fill="#e8edf3")
        fill_w = int((bar_x2 - bar_x1) * max(min(progress, 100), 0) / 100)
        if fill_w > 0:
            draw.rounded_rectangle((bar_x1, bar_y, bar_x1 + fill_w, bar_y + ry(7)), radius=ry(4), fill=MGA_RED if bad else MGA_TEAL)

    img.paste("#f4f7fb", (0, 0, w, h))
    draw.rounded_rectangle(rbox((16, 16, 1040, 88)), radius=rx(5), fill=MGA_BLUE)
    draw.text((rx(31), ry(24)), "MGA MANTENIMIENTO", font=subtitle_font, fill="#dbeafe")
    draw.text((rx(31), ry(40)), "KPI DIESEL", font=pil_font(max(18, int(w * 0.026)), True), fill="white")
    draw.text((rx(31), ry(70)), f"{start} a {end}", font=small_font, fill="#dbeafe")
    pil_right(draw, (rx(1024), ry(26)), "RENDIMIENTO PROMEDIO", subtitle_font, "#dbeafe")
    pil_right(draw, (rx(1024), ry(44)), rendimiento_text(avg_value), pil_font(max(22, int(w * 0.028)), True), "white")
    pil_right(draw, (rx(1024), ry(76)), f"Meta {meta:.1f} L/H", small_font, "#dbeafe")

    analyzed_days = max(1, (date.fromisoformat(end) - date.fromisoformat(start)).days + 1)
    total_stock = max(mga_liters + prosermin_liters, 1)
    cards = [
        ((16, 100, 179, 196), "Equipos", f"{len(active_rows)}", "con captura diesel", 100 if active_rows else 0, False),
        ((190, 100, 352, 196), "Consumo total", f"{total_liters:.1f} L", f"{analyzed_days} dias analizados", min(total_liters / 40000 * 100, 100), False),
        ((363, 100, 525, 196), "Diesel MGA", f"{mga_liters:.1f} L", "disponible MGA", mga_liters / total_stock * 100, False),
        ((536, 100, 698, 196), "Diesel PROSERMIN", f"{prosermin_liters:.1f} L", "disponible PROSERMIN", prosermin_liters / total_stock * 100, False),
        ((709, 100, 871, 196), "Horas trabajadas", f"{total_hours:.1f} h", "horas del periodo", min(total_hours / 1000 * 100, 100), False),
        ((882, 100, 1040, 196), "Rendimiento", f"{rendimiento_text(avg_value)}", f"Meta {meta:.1f} L/H", ((avg_value or 0) / max(meta, 1)) * 100, avg_value is not None and avg_value > meta),
    ]
    for item in cards:
        card(*item)

    chart_box = rbox((16, 209, 628, 622))
    draw.rounded_rectangle(chart_box, radius=rx(7), fill="white", outline="#d5dde8")
    draw.text((chart_box[0] + rx(12), chart_box[1] + ry(14)), "Consumo por equipo", font=subtitle_font, fill=MGA_BLUE)
    pil_right(draw, (chart_box[2] - rx(16), chart_box[1] + ry(15)), "TOP 12", small_font, "#667085")
    chart_rows = sorted(active_rows, key=lambda row: row["diesel_liters"], reverse=True)[:12]
    max_liters = max([row["diesel_liters"] for row in chart_rows] or [1])
    list_y = chart_box[1] + ry(45)
    row_gap = max(20, int((chart_box[3] - list_y - ry(14)) / max(len(chart_rows), 1)))
    for idx, row in enumerate(chart_rows):
        y = list_y + idx * row_gap
        name = pil_truncate(draw, row["equipment"], table_bold, rx(112))
        draw.text((chart_box[0] + rx(12), y), name, font=table_bold, fill=MGA_BLUE)
        draw.text((chart_box[0] + rx(12), y + ry(14)), row["status"], font=small_font, fill="#6b7280")
        bar_x1 = chart_box[0] + rx(132)
        bar_x2 = chart_box[2] - rx(150)
        bar_y = y + ry(7)
        draw.rounded_rectangle((bar_x1, bar_y, bar_x2, bar_y + ry(9)), radius=ry(5), fill="#e8edf3")
        fill_w = int((bar_x2 - bar_x1) * row["diesel_liters"] / max_liters)
        if fill_w > 0:
            draw.rounded_rectangle((bar_x1, bar_y, bar_x1 + fill_w, bar_y + ry(9)), radius=ry(5), fill=MGA_RED if diesel_bad(row) else MGA_TEAL)
        pil_right(draw, (chart_box[2] - rx(76), y + ry(4)), f"{row['diesel_liters']:.1f} L", table_bold, "#344054")
        pil_right(draw, (chart_box[2] - rx(12), y + ry(4)), rendimiento_text(row["rendimiento_lh"]), small_font, "#64748b")

    right_x1, right_x2 = rx(640), rx(1040)
    panel1 = (right_x1, ry(219), right_x2, ry(319))
    draw.rounded_rectangle(panel1, radius=rx(8), fill="white", outline="#d5dde8")
    draw.text((panel1[0] + rx(12), panel1[1] + ry(28)), "RENDIMIENTO PROMEDIO", font=small_font, fill="#667085")
    pil_center(draw, ((panel1[0] + panel1[2]) / 2, panel1[1] + ry(31)), rendimiento_text(avg_value), pil_font(max(18, int(w * 0.030)), True), MGA_BLUE)
    pil_right(draw, (panel1[2] - rx(16), panel1[1] + ry(32)), f"Meta {meta:.1f} L/H", small_font, "#667085")
    bar_x1, bar_x2 = panel1[0] + rx(22), panel1[2] - rx(22)
    bar_y = panel1[1] + ry(52)
    draw.rounded_rectangle((bar_x1, bar_y, bar_x2, bar_y + ry(13)), radius=ry(7), fill="#e8edf3")
    progress = min(((avg_value or 0) / max(meta, 1)) * 100, 100)
    if progress:
        draw.rounded_rectangle((bar_x1, bar_y, bar_x1 + int((bar_x2 - bar_x1) * progress / 100), bar_y + ry(13)), radius=ry(7), fill=MGA_RED if avg_value and avg_value > meta else MGA_TEAL)
    marker_x = bar_x1 + int((bar_x2 - bar_x1) * min(meta / max(meta * 1.2, 1), 1))
    draw.line((marker_x, bar_y - ry(6), marker_x, bar_y + ry(19)), fill=MGA_BLUE, width=max(1, rx(2)))

    panel2 = (right_x1, ry(328), right_x2, ry(456))
    draw.rounded_rectangle(panel2, radius=rx(8), fill="white", outline="#d5dde8")
    draw.text((panel2[0] + rx(12), panel2[1] + ry(18)), "Existencia disponible", font=subtitle_font, fill=MGA_BLUE)
    stock_x1, stock_x2 = panel2[0] + rx(22), panel2[2] - rx(22)
    stock_y = panel2[1] + ry(52)
    draw.rounded_rectangle((stock_x1, stock_y, stock_x2, stock_y + ry(16)), radius=ry(8), fill="#e8edf3")
    mga_w = int((stock_x2 - stock_x1) * mga_liters / total_stock)
    pro_w = int((stock_x2 - stock_x1) * prosermin_liters / total_stock)
    if mga_w > 0:
        draw.rounded_rectangle((stock_x1, stock_y, stock_x1 + mga_w, stock_y + ry(16)), radius=ry(8), fill=MGA_TEAL)
    if pro_w > 0:
        draw.rounded_rectangle((stock_x1 + mga_w, stock_y, stock_x1 + mga_w + pro_w, stock_y + ry(16)), radius=ry(8), fill="#f59e0b")
    for idx, (label, value, color) in enumerate([("MGA", mga_liters, MGA_TEAL), ("PROSERMIN", prosermin_liters, "#f59e0b")]):
        yy = panel2[1] + ry(85 + idx * 23)
        draw.rectangle((panel2[0] + rx(24), yy - ry(5), panel2[0] + rx(32), yy + ry(3)), fill=color)
        draw.text((panel2[0] + rx(40), yy - ry(8)), label, font=small_font, fill="#64748b")
        pil_right(draw, (panel2[2] - rx(18), yy - ry(8)), f"{value:.1f} L", table_bold, "#64748b")

    list_panel = (right_x1, ry(466), right_x2, ry(622))
    top_side_rows = chart_rows[:4]
    sub_h = int((list_panel[3] - list_panel[1]) / max(len(top_side_rows), 1))
    for idx, row in enumerate(top_side_rows):
        y1 = list_panel[1] + idx * sub_h
        draw.rounded_rectangle((list_panel[0], y1, list_panel[2], y1 + sub_h - ry(5)), radius=rx(5), fill="white", outline="#edf1f6")
        draw.text((list_panel[0] + rx(14), y1 + ry(15)), pil_truncate(draw, row["equipment"], table_bold, rx(180)), font=table_bold, fill=MGA_BLUE)
        pil_right(draw, (list_panel[2] - rx(16), y1 + ry(16)), f"{row['status']} | {rendimiento_text(row['rendimiento_lh'])}", small_font, "#64748b")

    table_left = rx(16)
    table_top = ry(633)
    table_w = w - rx(32)
    headers = ["EQUIPO", "CONDICION", "HI", "HF", "HRS TRAB", "DIESEL L", "REND. L/H", "META", "KPI"]
    weights = [0.18, 0.20, 0.08, 0.08, 0.10, 0.12, 0.10, 0.07, 0.09]
    col_w = [int(table_w * value / sum(weights)) for value in weights]
    col_w[-1] += table_w - sum(col_w)
    table_rows = sorted(normalized_rows, key=lambda row: row["diesel_liters"], reverse=True)
    max_table_rows = max(1, min(len(table_rows), int((h - table_top - ry(44)) / max(20, ry(26)))))
    table_rows = table_rows[:max_table_rows]
    row_h = max(20, int((h - table_top - ry(22)) / max(len(table_rows) + 2, 1)))
    x = table_left
    for header, cw in zip(headers, col_w):
        pil_table_cell(draw, (x, table_top, x + cw, table_top + row_h), header, table_bold, "#dfe8f3", "#d5dde8", "#0f274f")
        x += cw
    for idx, row in enumerate(table_rows, 1):
        y = table_top + idx * row_h
        fill = "#ffffff" if idx % 2 else "#f3f7fb"
        values = [
            row["equipment"],
            row["condition"],
            f"{row['horometer_initial']:.1f}" if row["horometer_initial"] else "",
            f"{row['horometer_final']:.1f}" if row["horometer_final"] else "",
            f"{row['worked_hours']:.1f}",
            f"{row['diesel_liters']:.1f}",
            rendimiento_text(row["rendimiento_lh"]),
            f"{meta:.1f}",
            row["status"],
        ]
        x = table_left
        for col_idx, (value, cw) in enumerate(zip(values, col_w)):
            if col_idx == 8:
                draw.rectangle((x, y, x + cw, y + row_h), fill=fill, outline="#d5dde8")
                status_pill((x + rx(8), y + ry(6), x + cw - rx(8), y + row_h - ry(6)), str(value), diesel_bad(row))
            else:
                text_fill = MGA_BLUE if col_idx == 0 else "#344054"
                pil_table_cell(draw, (x, y, x + cw, y + row_h), value, table_font if col_idx != 0 else table_bold, fill, "#d5dde8", text_fill, "left" if col_idx in (0, 1) else "center")
            x += cw
    total_y = table_top + (len(table_rows) + 1) * row_h
    x = table_left
    total_values = ["Total", "", "", "", f"{total_hours:.1f}", f"{total_liters:.1f}", rendimiento_text(avg_value), f"{meta:.1f}", f"{critical} revision"]
    for col_idx, (value, cw) in enumerate(zip(total_values, col_w)):
        pil_table_cell(draw, (x, total_y, x + cw, total_y + row_h), value, table_bold, "#eafaf5", "#d5dde8", MGA_BLUE if col_idx == 0 else "#0f274f", "center")
        x += cw
    img.save(path, quality=95)


def create_monthly_oil_image(portal: dict[str, Any], path: Path, size: tuple[int, int], start: str, end: str, month_name: str, year: int) -> None:
    oil = portal.get("oil_kpi") if isinstance(portal.get("oil_kpi"), dict) else {}
    rows = oil.get("rows") if isinstance(oil.get("rows"), list) else []
    totals = dict(oil.get("totals") if isinstance(oil.get("totals"), dict) else {})
    w, h = size
    img = Image.new("RGB", size, "#eeeeee")
    draw = ImageDraw.Draw(img)
    title_font = pil_font(max(18, int(h * 0.022)), True)
    section_font = pil_font(max(14, int(h * 0.018)), True)
    value_font = pil_font(max(28, int(h * 0.035)), True)
    small_font = pil_font(max(9, int(h * 0.012)))
    axis_font = pil_font(max(10, int(h * 0.014)))
    table_title_font = pil_font(max(13, int(h * 0.015)), True)
    table_header_font = pil_font(max(9, int(h * 0.012)), True)
    table_cell_font = pil_font(max(8, int(h * 0.011)), True)
    oil_columns = [
        ("Motor 15W40", "oil_motor_15w40"),
        ("ISO 68", "oil_hco_iso68"),
        ("SAE 30", "oil_trans_sae30"),
        ("SAE 50", "oil_sae50"),
        ("85W140", "oil_85w140"),
    ]
    for _label, key in oil_columns:
        totals.setdefault(key, sum(parse_float(row.get(key), 0) for row in rows))
    month_idx = MONTH_NAMES_ES_FULL.index(month_name) + 1 if month_name in MONTH_NAMES_ES_FULL else 1
    month_short = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"][max(min(month_idx, 12), 1) - 1]
    month_label = f"{month_short}-{str(year)[-2:]}"

    def row_total(row: dict[str, Any]) -> float:
        return sum(parse_float(row.get(key), 0) for _label, key in oil_columns)

    def oil_group(row: dict[str, Any]) -> str:
        group = normalized_ascii(row.get("group"))
        if group in {"BARRENACION", "REZAGADO", "UTILITARIO"}:
            return group
        code = normalized_ascii(row.get("code") or row.get("equipment_code"))
        text = normalized_ascii(f"{code} {row.get('description') or ''}")
        if code.startswith(("JL", "JA")) or "JUMBO" in text:
            return "BARRENACION"
        if code.startswith("ST") or "SCOOP" in text or "R1300" in text or "R1600" in text:
            return "REZAGADO"
        return "UTILITARIO"

    def metric_section(box, title, period_value, accumulated_value, bar_color):
        x1, y1, x2, y2 = box
        title_h = max(24, int((y2 - y1) * 0.24))
        gap = 6
        card_y1 = y1 + title_h
        card_w = (x2 - x1 - gap) // 2
        draw.rectangle((x1, y1, x2, y1 + title_h), fill="white", outline="#e5e7eb")
        pil_center(draw, ((x1 + x2) / 2, y1 + title_h / 2), title, section_font, "#707780")
        for cx1, cy1, cx2, cy2, value, caption, color in [
            (x1, card_y1, x1 + card_w, y2, period_value, "Consumo semanal", bar_color),
            (x1 + card_w + gap, card_y1, x2, y2, accumulated_value, "Consumo total acumulado", "#a40000" if bar_color != MGA_TEAL else bar_color),
        ]:
            draw.rectangle((cx1, cy1, cx2, cy2), fill="white")
            pil_center(draw, ((cx1 + cx2) / 2, cy1 + (cy2 - cy1) * 0.32), f"{value:.2f}", value_font, "#777d86")
            bar_y = cy1 + int((cy2 - cy1) * 0.56)
            bar_h = max(7, int((cy2 - cy1) * 0.10))
            draw.rectangle((cx1 + 4, bar_y, cx2 - 4, bar_y + bar_h), fill=color)
            pil_center(draw, ((cx1 + cx2) / 2, cy2 - max(15, int((cy2 - cy1) * 0.14))), caption, small_font, "#4b5563")

    top_h = max(30, int(h * 0.036))
    draw.rectangle((0, 0, w, top_h), fill="white")
    pil_center(draw, (w / 2, top_h / 2), 'Consumo de aceite de equipos "Providencia"', title_font, "#333333")
    pil_center(draw, (int(w * 0.23), top_h / 2), month_label, section_font, "#111111")
    pil_center(draw, (int(w * 0.94), top_h / 2), month_label, section_font, "#111111")

    side_w = int(w * 0.265)
    center_gap = int(w * 0.015)
    center_x = side_w + center_gap
    right_x = w - side_w
    center_w = right_x - center_x - center_gap
    section_h = int(h * 0.125)
    section_1_y = top_h + int(h * 0.005)
    section_2_y = section_1_y + section_h + int(h * 0.052)
    for box, title, key, color in [
        ((0, section_1_y, side_w - 6, section_1_y + section_h), "Consumo de Aceite HCO", "oil_hco_iso68", MGA_TEAL),
        ((0, section_2_y, side_w - 6, section_2_y + section_h), "Consumo de Aceite SAE 30", "oil_trans_sae30", "#d76f75"),
        ((right_x + 6, section_1_y, w, section_1_y + section_h), "Consumo de Aceite de Motor", "oil_motor_15w40", "#d76f75"),
        ((right_x + 6, section_2_y, w, section_2_y + section_h), "Consumo de Aceite SAE 50", "oil_sae50", "#d76f75"),
    ]:
        value = parse_float(totals.get(key), 0)
        metric_section(box, title, value, value, color)

    chart_box = (center_x, top_h + int(h * 0.020), center_x + center_w, int(h * 0.425))
    x1, y1, x2, y2 = chart_box
    draw.rectangle(chart_box, fill="white", outline="#d1d5db")
    chart_rows = sorted([row for row in rows if row_total(row) > 0], key=row_total, reverse=True)[:6] or rows[:6]
    plot_x = x1 + int((x2 - x1) * 0.07)
    plot_y = y1 + int((y2 - y1) * 0.11)
    plot_w = int((x2 - x1) * 0.88)
    plot_h = int((y2 - y1) * 0.64)
    max_value = max([max([parse_float(row.get(key), 0) for _label, key in oil_columns] or [0]) for row in chart_rows] or [1])
    axis_max = max(100, int(((max_value * 1.25) + 9) // 10 * 10))
    colors_by_column = {
        "oil_motor_15w40": "#4472c4",
        "oil_hco_iso68": "#ed7d31",
        "oil_trans_sae30": "#a5a5a5",
        "oil_sae50": "#ffc000",
        "oil_85w140": "#5b9bd5",
    }
    for value in range(0, axis_max + 1, max(axis_max // 5, 10)):
        yy = plot_y + plot_h - int(plot_h * value / axis_max)
        draw.line((plot_x, yy, plot_x + plot_w, yy), fill="#d9d9d9", width=1)
        pil_right(draw, (plot_x - 10, yy - 8), str(value), axis_font, "#111111")
    if chart_rows:
        cluster_w = plot_w / max(len(chart_rows), 1)
        bar_w = max(2, int(cluster_w / (len(oil_columns) + 2)))
        for idx, row in enumerate(chart_rows):
            cluster_x = plot_x + idx * cluster_w + max(2, int(cluster_w * 0.13))
            for series_idx, (_label, key) in enumerate(oil_columns):
                value = parse_float(row.get(key), 0)
                bar_h = int(plot_h * value / axis_max) if axis_max else 0
                bx1 = int(cluster_x + series_idx * bar_w)
                bx2 = bx1 + max(1, bar_w - 1)
                by1 = plot_y + plot_h - bar_h
                if value:
                    draw.rectangle((bx1, by1, bx2, plot_y + plot_h), fill=colors_by_column[key])
                pil_center(draw, ((bx1 + bx2) / 2, by1 - 10 if value else plot_y + plot_h - 10), f"{value:.0f}", small_font, "#111111")
            pil_center(draw, (cluster_x + (len(oil_columns) * bar_w) / 2, plot_y + plot_h + 18), row.get("code") or row.get("equipment_code") or "", axis_font, "#111111")
    legend_y = y2 - int((y2 - y1) * 0.08)
    legend_x = x1 + int((x2 - x1) * 0.25)
    for label, key in oil_columns:
        draw.rectangle((legend_x, legend_y - 5, legend_x + 10, legend_y + 5), fill=colors_by_column[key])
        draw.text((legend_x + 15, legend_y - 9), label, font=axis_font, fill="#111111")
        legend_x += int((x2 - x1) * 0.14)

    grouped = {"BARRENACION": [], "REZAGADO": [], "UTILITARIO": []}
    for row in rows:
        grouped.setdefault(oil_group(row), []).append(row)
    group_totals = {}
    for group_name, group_rows in grouped.items():
        group_totals[group_name] = {
            "period": sum(parse_float(row.get("period_hours") or row.get("period"), 0) for row in group_rows),
            "worked": sum(parse_float(row.get("worked_hours") or row.get("worked"), 0) for row in group_rows),
            **{key: sum(parse_float(row.get(key), 0) for row in group_rows) for _label, key in oil_columns},
        }
    total_values = {
        "period": sum(group_totals[g]["period"] for g in group_totals),
        "worked": sum(group_totals[g]["worked"] for g in group_totals),
        **{key: sum(group_totals[g][key] for g in group_totals) for _label, key in oil_columns},
    }
    table_x = int(w * 0.07)
    table_top = int(h * 0.525)
    table_w = int(w * 0.64)
    table_box = (table_x, table_top, table_x + table_w, h - int(h * 0.025))
    tx1, ty1, tx2, ty2 = table_box
    draw.rectangle((tx1, ty1 - 68, tx2, ty1 - 6), fill="#f3f3f3")
    pil_center(draw, ((tx1 + tx2) / 2, ty1 - 52), "REPORTE SEMANAL CONSUMO DE ACEITES", table_title_font, "#111111")
    draw.text((tx1 + int(table_w * 0.40), ty1 - 28), "Dia Inicial:", font=table_header_font, fill="#707780")
    draw.rectangle((tx1 + int(table_w * 0.52), ty1 - 42, tx1 + int(table_w * 0.61), ty1 - 18), fill="white")
    pil_center(draw, (tx1 + int(table_w * 0.565), ty1 - 30), str(int(str(start)[-2:])), table_header_font, "#111111")
    draw.text((tx1 + int(table_w * 0.74), ty1 - 28), "Dia Final:", font=table_header_font, fill="#707780")
    draw.rectangle((tx1 + int(table_w * 0.84), ty1 - 42, tx1 + int(table_w * 0.93), ty1 - 18), fill="white")
    pil_center(draw, (tx1 + int(table_w * 0.885), ty1 - 30), str(int(str(end)[-2:])), table_header_font, "#111111")
    headers = ["# Eco", "Equipo", "Hrs\nPeriodo", "Hrs\nTrab", "Consumo\nMotor\n15W40", "Consumo\nISO 68", "SAE30", "SAE 50", "85W140"]
    weights = [1.05, 2.65, 0.95, 0.95, 1.05, 0.95, 0.95, 0.95, 0.95]
    widths = [int(table_w * weight / sum(weights)) for weight in weights]
    widths[-1] += table_w - sum(widths)
    table_rows = []
    subtotal_labels = {"BARRENACION": "ACUMULADO EQ'S DE\nBARRENACION", "REZAGADO": "EQUIPO REZAGADO", "UTILITARIO": "EQUIPO UTILITARIO"}
    for group_name in ("BARRENACION", "REZAGADO", "UTILITARIO"):
        for row in grouped.get(group_name, [])[:5]:
            table_rows.append(("row", row, group_name))
        if grouped.get(group_name):
            table_rows.append(("subtotal", group_totals[group_name], subtotal_labels[group_name]))
    table_rows.append(("total", total_values, "Total de Aceite Utilizado"))
    row_h = max(14, min(28, int((ty2 - ty1) / max(len(table_rows) + 1, 1))))
    header_h = max(row_h + 4, 30)
    x = tx1
    for header, col_w in zip(headers, widths):
        draw.rectangle((x, ty1, x + col_w, ty1 + header_h), fill="white", outline="#111111")
        draw.text((x + 3, ty1 + 3), header.replace("\n", " "), font=table_header_font, fill="#666666")
        x += col_w
    current_y = ty1 + header_h
    for kind, source, label in table_rows:
        fill = "#fff200" if kind == "total" else "#ffd966" if kind == "subtotal" else "#efefef"
        if kind == "row":
            values = [
                source.get("code") or source.get("equipment_code"),
                source.get("description"),
                f"{parse_float(source.get('period_hours') or source.get('period'), 0):.1f}",
                f"{parse_float(source.get('worked_hours') or source.get('worked'), 0):.1f}" if parse_float(source.get("worked_hours") or source.get("worked"), 0) else "",
                f"{parse_float(source.get('oil_motor_15w40'), 0):.2f}",
                f"{parse_float(source.get('oil_hco_iso68'), 0):.1f}",
                f"{parse_float(source.get('oil_trans_sae30'), 0):.2f}",
                f"{parse_float(source.get('oil_sae50'), 0):.2f}",
                f"{parse_float(source.get('oil_85w140'), 0):.2f}",
            ]
        else:
            values = ["", label, f"{source.get('period', 0):.1f}" if kind != "total" else "", f"{source.get('worked', 0):.1f}", *[f"{source.get(key, 0):.1f}" for _label, key in oil_columns]]
        x = tx1
        for cidx, (value, col_w) in enumerate(zip(values, widths)):
            text_fill = "#f05b5b" if kind == "row" and cidx >= 4 and parse_float(value, 0) == 0 else "#666666"
            pil_table_cell(draw, (x, current_y, x + col_w, current_y + row_h), value, small_font if cidx == 1 else table_cell_font, fill, "#111111", text_fill)
            x += col_w
        current_y += row_h
        if current_y > ty2:
            break

    stock_box = (table_x + table_w + int(w * 0.018), table_top, w - int(w * 0.025), h - int(h * 0.025))
    sx1, sy1, sx2, sy2 = stock_box
    draw.rectangle(stock_box, fill="white", outline="#cbd5e1")
    header_h = max(34, int((sy2 - sy1) * 0.09))
    draw.rectangle((sx1, sy1, sx2, sy1 + header_h), fill=MGA_BLUE)
    pil_center(draw, ((sx1 + sx2) / 2, sy1 + header_h / 2), "PEDIDO DE LUBRICANTES", table_title_font, "white")
    for idx, (label, value) in enumerate([("Pedido 7d", 0), ("Pedido 15d", 0), ("Pedido 30d", 0)]):
        card_w = (sx2 - sx1) / 3
        cx1 = int(sx1 + idx * card_w)
        cx2 = int(sx1 + (idx + 1) * card_w)
        draw.rectangle((cx1, sy1 + header_h, cx2, sy1 + header_h + 52), fill="#f8fafc", outline="#e2e8f0")
        pil_center(draw, ((cx1 + cx2) / 2, sy1 + header_h + 20), f"{value:.1f} L", table_header_font, MGA_TEAL)
        pil_center(draw, ((cx1 + cx2) / 2, sy1 + header_h + 38), label, small_font, "#475569")
    img.save(path, quality=95)


def picture_pixel_size(shape) -> tuple[int, int]:
    try:
        image = Image.open(BytesIO(shape.image.blob))
        return max(image.size[0], 1), max(image.size[1], 1)
    except Exception:
        return max(int(shape.width / 9525), 1), max(int(shape.height / 9525), 1)


def replace_ppt_picture(slide, shape, image_path: Path) -> None:
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    parent = shape._element.getparent()
    index = parent.index(shape._element)
    parent.remove(shape._element)
    new_picture = slide.shapes.add_picture(str(image_path), left, top, width, height)
    new_element = new_picture._element
    parent.remove(new_element)
    parent.insert(index, new_element)


def replace_single_picture(slide, image_path: Path) -> None:
    pictures = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if not pictures:
        return
    replace_ppt_picture(slide, max(pictures, key=lambda shape: shape.width * shape.height), image_path)


def add_full_slide_picture(slide, prs: Presentation, image_path: Path) -> None:
    slide.shapes.add_picture(str(image_path), 0, 0, width=prs.slide_width, height=prs.slide_height)


def slide_text(slide) -> str:
    values: list[str] = []
    for shape in iter_pptx_shapes(slide.shapes):
        if getattr(shape, "has_text_frame", False):
            values.append(shape.text)
    return " ".join(values).upper()


def replace_slide_report_pictures(slide, tmp_dir: Path, portal: dict[str, Any], group: str, start: str, end: str) -> None:
    pictures = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if len(pictures) < 2:
        return
    pictures.sort(key=lambda item: item.top)
    top_shape, table_shape = pictures[0], pictures[1]
    dashboard_path = tmp_dir / f"{kpi_format_key(group) or 'kpi'}_dashboard.jpg"
    table_path = tmp_dir / f"{kpi_format_key(group) or 'kpi'}_table.jpg"
    create_monthly_kpi_dashboard_image(portal, group, start, end, dashboard_path, picture_pixel_size(top_shape))
    create_monthly_kpi_table_image(portal, group, start, end, table_path, picture_pixel_size(table_shape))
    replace_ppt_picture(slide, top_shape, dashboard_path)
    replace_ppt_picture(slide, table_shape, table_path)


def replace_or_add_oil_report_slide(prs: Presentation, tmp_dir: Path, portal: dict[str, Any], start: str, end: str, month_name: str, year: int) -> None:
    target_slide = None
    for slide in prs.slides:
        if "ACEITE" in slide_text(slide) or "ACEITES" in slide_text(slide):
            target_slide = slide
            break
    if target_slide is None:
        target_slide = prs.slides.add_slide(ppt_blank_layout(prs))
        marker = target_slide.shapes.add_textbox(0, 0, 1, 1)
        marker.text = "KPI ACEITES"
    pictures = [shape for shape in target_slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
    size = picture_pixel_size(max(pictures, key=lambda shape: shape.width * shape.height)) if pictures else (1600, 1200)
    oil_path = tmp_dir / "kpi_aceites.jpg"
    create_monthly_oil_image(portal, oil_path, size, start, end, month_name, year)
    if pictures:
        replace_single_picture(target_slide, oil_path)
    else:
        add_full_slide_picture(target_slide, prs, oil_path)


def period_report_pptx_bytes(
    portal: dict[str, Any],
    start: str,
    end: str,
    month_name: str,
    year: int,
    diesel_data: dict[str, Any] | None = None,
    weekly_label: str = "",
) -> bytes:
    if isinstance(diesel_data, dict):
        portal = dict(portal)
        portal["diesel"] = diesel_data
    prs = Presentation(str(MONTHLY_REPORT_TEMPLATE_PATH)) if MONTHLY_REPORT_TEMPLATE_PATH.exists() else Presentation()
    if len(prs.slides) == 0:
        prs.slides.add_slide(ppt_blank_layout(prs))
    if weekly_label:
        update_weekly_ppt_text(prs, weekly_label, month_name, year)
    else:
        update_monthly_ppt_text(prs, month_name, year)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        if len(prs.slides) >= 3:
            replace_slide_report_pictures(prs.slides[2], tmp_dir, portal, "Equipos de Barrenacion", start, end)
        if len(prs.slides) >= 4:
            replace_slide_report_pictures(prs.slides[3], tmp_dir, portal, "Equipos de Rezagado", start, end)
        if len(prs.slides) >= 5:
            pictures = [shape for shape in prs.slides[4].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
            size = picture_pixel_size(max(pictures, key=lambda shape: shape.width * shape.height)) if pictures else (440, 534)
            tire_path = tmp_dir / "vida_util_llantas.jpg"
            create_monthly_tire_image(portal, tire_path, size, month_name, year)
            replace_single_picture(prs.slides[4], tire_path)
        if len(prs.slides) >= 6:
            pictures = [shape for shape in prs.slides[5].shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
            size = picture_pixel_size(max(pictures, key=lambda shape: shape.width * shape.height)) if pictures else (1056, 1374)
            diesel_path = tmp_dir / "diesel_acumulado.jpg"
            create_monthly_diesel_image(portal, diesel_path, size, start, end, month_name, year)
            replace_single_picture(prs.slides[5], diesel_path)
        replace_or_add_oil_report_slide(prs, tmp_dir, portal, start, end, month_name, year)
    stream = BytesIO()
    prs.save(stream)
    return stream.getvalue()


def monthly_report_pptx_bytes(portal: dict[str, Any], year: int, month: int, diesel_data: dict[str, Any] | None = None) -> bytes:
    start, end = month_bounds(year, month)
    month_name = MONTH_NAMES_ES_FULL[month - 1]
    return period_report_pptx_bytes(portal, start, end, month_name, year, diesel_data)


def weekly_report_pptx_bytes(portal: dict[str, Any], start: date, end: date, diesel_data: dict[str, Any] | None = None) -> bytes:
    month_name = MONTH_NAMES_ES_FULL[start.month - 1]
    return period_report_pptx_bytes(
        portal,
        start.isoformat(),
        end.isoformat(),
        month_name,
        start.year,
        diesel_data,
        weekly_period_label(start, end),
    )


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "ok": True,
        "service": SERVICE_NAME,
        "version": app.version,
        "database": database_status(),
        "generated_at": utc_now().isoformat(timespec="seconds"),
    }


@app.get("/favicon.ico")
def favicon():
    logo = STATIC_DIR / "mga-corner-logo.jfif"
    if logo.exists():
        return FileResponse(logo, media_type="image/jpeg")
    return Response(status_code=204)


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
        return catalog_with_inventory(session)


@app.get("/api/portal")
def get_portal(response: Response) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        return latest_portal_payload(session)


@app.get("/api/products")
def get_products(response: Response, q: str = Query(default=""), limit: int = Query(default=120, ge=1, le=25000)) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        rows = product_rows(session, q, limit)
        return {"ok": True, "products": rows, "count": len(rows)}


@app.get("/api/requisitions")
def get_requisitions(response: Response) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        rows = session.scalars(select(CloudRequisition).order_by(CloudRequisition.request_date.desc(), CloudRequisition.id.desc()).limit(300)).all()
        return {
            "ok": True,
            "next_folio": next_requisition_folio(session),
            "requisitions": [requisition_payload(row) for row in rows],
        }


@app.get("/api/requisitions/{requisition_id}")
def get_requisition(requisition_id: int) -> dict[str, Any]:
    with SessionLocal() as session:
        row = session.get(CloudRequisition, requisition_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Requisicion no encontrada.")
        return {"ok": True, "requisition": requisition_payload(row, include_items=True)}


@app.post("/api/requisition-tracking/import")
async def import_requisition_tracking(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Carga invalida.")
    data = str(payload.get("data") or "")
    if "," in data and data.startswith("data:"):
        data = data.split(",", 1)[1]
    try:
        raw = base64.b64decode(data)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"No se pudo leer el archivo: {exc}")
    file_name = str(payload.get("file_name") or "seguimiento_requisiciones.xlsx")
    try:
        wb = load_workbook(BytesIO(raw), read_only=True, data_only=True)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Excel invalido: {exc}")

    parsed_rows: list[dict[str, Any]] = []
    skipped = 0
    try:
        sheet_layouts: list[tuple[str, int, dict[int, str], bool]] = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            header_row = None
            fields: dict[int, str] = {}
            for idx, values in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 30), values_only=True), 1):
                mapped = {col_idx: requisition_tracking_field_for(value) for col_idx, value in enumerate(values)}
                mapped = {col_idx: field for col_idx, field in mapped.items() if field}
                if "folio" in mapped.values() and (
                    "purchase_status" in mapped.values()
                    or "purchase_order" in mapped.values()
                    or "supplier" in mapped.values()
                    or "expected_date" in mapped.values()
                ):
                    header_row = idx
                    fields = mapped
                    break
            if header_row is None:
                continue
            mtto_layout = fields.get(2) == "folio" and fields.get(3) == "purchase_order"
            sheet_layouts.append((sheet_name, header_row, fields, mtto_layout))
        if any(layout[3] for layout in sheet_layouts):
            sheet_layouts = [layout for layout in sheet_layouts if layout[3]][-1:]
        for sheet_name, header_row, fields, mtto_layout in sheet_layouts:
            ws = wb[sheet_name]
            current_category = ""
            current_equipment = ""
            for values in ws.iter_rows(min_row=header_row + 1, values_only=True):
                if mtto_layout:
                    category = excel_cell_text(values[0] if len(values) > 0 else None)
                    equipment = excel_cell_text(values[1] if len(values) > 1 else None)
                    if category:
                        current_category = category
                    if equipment:
                        current_equipment = equipment
                item: dict[str, Any] = {"sheet": sheet_name}
                for col_idx, field in fields.items():
                    item[field] = values[col_idx] if col_idx < len(values) else None
                if mtto_layout and current_equipment and "equipment" not in item:
                    item["equipment"] = current_equipment
                if mtto_layout and current_category:
                    item["category"] = current_category
                raw_folio = item.get("folio")
                if mtto_layout and not REQUISITION_REFERENCE_RE.search(str(raw_folio or "")):
                    skipped += 1
                    continue
                folio, parsed_description = split_requisition_reference(raw_folio)
                if not folio:
                    skipped += 1
                    continue
                item["folio"] = folio
                if parsed_description and not excel_cell_text(item.get("description")):
                    item["description"] = parsed_description
                parsed_rows.append(item)
    finally:
        wb.close()
    if not parsed_rows:
        raise HTTPException(status_code=400, detail="No encontre columnas de folio y seguimiento/OC en el Excel.")

    create_missing = bool(payload.get("create_missing", False))
    unmatched: list[str] = []
    updated = 0
    created = 0
    now = utc_now()
    date_fields = {"purchase_order_date", "expected_date", "received_date"}
    text_fields = {"purchase_status", "purchase_order", "supplier", "buyer", "tracking_notes"}
    with SessionLocal() as session:
        for item in parsed_rows:
            row = requisition_lookup(session, str(item.get("folio") or ""))
            if row is None:
                if not create_missing:
                    unmatched.append(str(item.get("folio") or ""))
                    continue
                row = CloudRequisition(
                    folio=str(item.get("folio") or ""),
                    request_date=utc_now().date().isoformat(),
                    authorization_date=utc_now().date().isoformat(),
                    equipment=normalize_text(item.get("equipment") or "SIN RELACIONAR"),
                    notes=normalize_text(item.get("description")),
                    created_at=now,
                )
                session.add(row)
                created += 1
            changed = False
            equipment = normalize_text(item.get("equipment"))
            if equipment and normalize_text(row.equipment) in {"", "SIN RELACIONAR"}:
                row.equipment = equipment
                changed = True
            description = normalize_text(item.get("description"))
            if description and not str(row.notes or "").strip():
                row.notes = description
                changed = True
            for field in text_fields:
                value = excel_cell_text(item.get(field))
                if value:
                    setattr(row, field, normalize_text(value) if field != "tracking_notes" else value)
                    changed = True
            for field in date_fields:
                value = excel_cell_text(item.get(field))
                if value:
                    setattr(row, field, iso_date(value))
                    changed = True
            inferred_status = infer_purchase_status(row.purchase_order, row.tracking_notes)
            if inferred_status and not row.purchase_status:
                row.purchase_status = inferred_status
                changed = True
            if changed:
                row.tracking_source_file = file_name
                row.tracking_updated_at = now
                row.updated_at = now
                updated += 1
        session.commit()
        rows = session.scalars(select(CloudRequisition).order_by(CloudRequisition.request_date.desc(), CloudRequisition.id.desc()).limit(300)).all()
        return {
            "ok": True,
            "imported_rows": len(parsed_rows),
            "updated": updated,
            "created": created,
            "skipped": skipped,
            "unmatched": sorted(set(unmatched))[:80],
            "requisitions": [requisition_payload(row) for row in rows],
        }


@app.post("/api/requisitions")
async def save_requisition(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Requisicion invalida.")
    with SessionLocal() as session:
        requisition_id = int(parse_float(payload.get("id"), 0) or 0)
        row = session.get(CloudRequisition, requisition_id) if requisition_id else None
        folio = normalize_text(payload.get("folio")) or next_requisition_folio(session)
        existing = session.scalar(select(CloudRequisition).where(CloudRequisition.folio == folio))
        if existing is not None and (row is None or existing.id != row.id):
            row = existing
        if row is None:
            row = CloudRequisition(folio=folio, created_at=utc_now())
            session.add(row)
        row.folio = folio
        row.request_date = iso_date(payload.get("request_date"))
        row.authorization_date = iso_date(payload.get("authorization_date") or payload.get("request_date"))
        row.equipment = normalize_text(payload.get("equipment") or "PARA STOCK")
        row.cost_center = normalize_text(payload.get("cost_center"))
        row.request_area = normalize_text(payload.get("request_area") or "MTTO")
        row.location = normalize_text(payload.get("location") or "PROVIDENCIA")
        row.requesting_unit = normalize_text(payload.get("requesting_unit") or "TALLER CENTRAL")
        row.operating_unit = normalize_text(payload.get("operating_unit") or "PROVIDENCIA")
        row.priority = normalize_text(payload.get("priority") or "URGENTE")
        row.recommendation = normalize_text(payload.get("recommendation") or "ORIGINAL")
        row.status = str(payload.get("status") or "Abierta").strip() or "Abierta"
        row.notes = str(payload.get("notes") or "").strip()
        if any(key in payload for key in ("purchase_status", "purchase_order", "purchase_order_date", "supplier", "buyer", "expected_date", "received_date", "tracking_notes")):
            row.purchase_status = str(payload.get("purchase_status") or "").strip()
            row.purchase_order = normalize_text(payload.get("purchase_order"))
            row.purchase_order_date = iso_date(payload.get("purchase_order_date")) if payload.get("purchase_order_date") else ""
            row.supplier = normalize_text(payload.get("supplier"))
            row.buyer = normalize_text(payload.get("buyer"))
            row.expected_date = iso_date(payload.get("expected_date")) if payload.get("expected_date") else ""
            row.received_date = iso_date(payload.get("received_date")) if payload.get("received_date") else ""
            row.tracking_notes = str(payload.get("tracking_notes") or "").strip()
            row.tracking_updated_at = utc_now()
        row.updated_at = utc_now()
        items = payload.get("items")
        if isinstance(items, list):
            row.items.clear()
            for idx, item in enumerate(items, start=1):
                if not isinstance(item, dict):
                    continue
                part_number = normalize_text(item.get("part_number"))
                description = normalize_text(item.get("description"))
                if not part_number and not description:
                    continue
                row.items.append(
                    CloudRequisitionItem(
                        quantity=parse_float(item.get("quantity"), 1) or 1,
                        unit=normalize_text(item.get("unit") or "PZA"),
                        part_number=part_number,
                        description=description,
                        sort_order=idx,
                        active=1,
                    )
                )
        session.commit()
        session.refresh(row)
        return {"ok": True, "requisition": requisition_payload(row, include_items=True), "next_folio": next_requisition_folio(session)}


@app.delete("/api/requisitions/{requisition_id}")
def delete_requisition(requisition_id: int, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        row = session.get(CloudRequisition, requisition_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Requisicion no encontrada.")
        session.delete(row)
        session.commit()
        return {"ok": True}


@app.get("/api/requisition-tracking/export")
def export_requisition_tracking(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> StreamingResponse:
    require_api_key(_auth)
    with SessionLocal() as session:
        workbook = build_requisition_tracking_workbook(session)
    filename = f"MTTO_PROVIDENCIA_SEGUIMIENTO_REQ_{utc_now().date().isoformat()}.xlsx"
    return StreamingResponse(
        BytesIO(workbook),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/api/requisitions/{requisition_id}/pdf")
def get_requisition_pdf(requisition_id: int) -> StreamingResponse:
    with SessionLocal() as session:
        row = session.get(CloudRequisition, requisition_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Requisicion no encontrada.")
        items = [item for item in sorted(row.items, key=lambda item: (item.sort_order, item.id)) if item.active]
        if not items:
            raise HTTPException(status_code=400, detail="Agrega al menos una partida antes de generar PDF.")
        pdf = requisition_pdf_bytes(row, items)
        filename = f"Requisicion_{row.folio}.pdf".replace(" ", "_")
        return StreamingResponse(
            BytesIO(pdf),
            media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="{filename}"'},
        )


@app.get("/api/kpi-format/pdf")
def get_kpi_format_pdf(group: str = Query(default="")) -> FileResponse:
    key, path = kpi_format_pdf_path(group)
    filename = KPI_FORMAT_FILENAMES.get(key, path.name)
    return FileResponse(
        path,
        media_type="application/pdf",
        filename=filename,
        headers={"Content-Disposition": f'inline; filename="{filename}"'},
    )


@app.get("/api/kpi-format/image")
def get_kpi_format_image(group: str = Query(default="")) -> StreamingResponse:
    key, path = kpi_format_pdf_path(group)
    image = kpi_format_image_bytes(path)
    filename = KPI_FORMAT_FILENAMES.get(key, path.name).replace(".pdf", ".png")
    return StreamingResponse(
        BytesIO(image),
        media_type="image/png",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/api/kpi-format/excel")
def get_kpi_format_excel(
    group: str = Query(default="Todos los equipos"),
    start: str = Query(default=""),
    end: str = Query(default=""),
    simulation: bool = Query(default=False),
    sim_name: str = Query(default="Escenario KPI"),
    sim_period_percent: float = Query(default=100),
    sim_worked_percent: float = Query(default=100),
    sim_mp_percent: float = Query(default=100),
    sim_mc_percent: float = Query(default=100),
    sim_stops_percent: float = Query(default=100),
    sim_mission_hours: float = Query(default=24),
    sim_meta_availability: float = Query(default=85),
    sim_meta_utilization: float = Query(default=75),
    sim_meta_reliability: float = Query(default=80),
    sim_meta_tmef: float = Query(default=8),
    sim_meta_tmpr: float = Query(default=4),
) -> StreamingResponse:
    with SessionLocal() as session:
        portal = latest_portal_payload(session)
    period = portal.get("period") if isinstance(portal.get("period"), dict) else {}
    today = utc_now().date().isoformat()
    start_text = start or str(period.get("start") or today)
    end_text = end or str(period.get("end") or start_text)
    start_date = parse_report_date(start_text, "start")
    end_date = parse_report_date(end_text, "end")
    if end_date < start_date:
        raise HTTPException(status_code=400, detail="La fecha final no puede ser menor que la fecha inicial.")

    key = kpi_format_key(group)
    normalized_group = normalize_text(group)
    if key in {"aceites", "llantas"} or "DIESEL" in normalized_group:
        raise HTTPException(status_code=400, detail="Excel editable disponible para Barrenacion, Rezagado, Acarreo y Utilitario.")
    if key == "barrenacion":
        groups = ["Equipos de Barrenacion"]
    elif key == "rezagado":
        groups = ["Equipos de Rezagado"]
    elif key == "acarreo":
        groups = ["Acarreo"]
    elif key == "utilitario":
        groups = ["Equipo Utilitario"]
    elif normalized_group and "TODO" not in normalized_group:
        raise HTTPException(status_code=400, detail="Selecciona Barrenacion, Rezagado, Acarreo, Utilitario o Todos los equipos.")
    else:
        groups = ["Equipos de Rezagado", "Equipos de Barrenacion", "Acarreo", "Equipo Utilitario"]
    settings = portal.get("settings") if isinstance(portal.get("settings"), dict) else {}
    reports = [monthly_kpi_report(portal, item, start_date.isoformat(), end_date.isoformat()) for item in groups]
    if simulation:
        scenario = {
            "name": sim_name,
            "period_percent": sim_period_percent,
            "worked_percent": sim_worked_percent,
            "mp_percent": sim_mp_percent,
            "mc_percent": sim_mc_percent,
            "stops_percent": sim_stops_percent,
            "reliability_mission_hours": sim_mission_hours,
        }
        reports = [simulate_monthly_kpi_report(report, scenario) for report in reports]
        settings = {
            **settings,
            "meta_availability": sim_meta_availability,
            "meta_utilization": sim_meta_utilization,
            "meta_reliability": sim_meta_reliability,
            "meta_tmef": sim_meta_tmef,
            "meta_tmpr": sim_meta_tmpr,
            "reliability_mission_hours": sim_mission_hours,
        }
    data = build_editable_kpi_excel(reports, settings)
    group_label = "Barrenacion_Rezagado_Acarreo_Utilitario" if len(groups) > 1 else re.sub(r"[^A-Za-z0-9_.-]+", "_", groups[0].replace("Equipos de ", ""))
    prefix = "KPI_SIMULACION" if simulation else "KPI"
    filename = f"{prefix}_{group_label}_{start_date.isoformat()}_{end_date.isoformat()}.xlsx"
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store, max-age=0",
        },
    )


@app.get("/api/monthly-report/powerpoint")
def get_monthly_report_powerpoint(
    year: int = Query(default=0),
    month: int = Query(default=0),
) -> StreamingResponse:
    now = utc_now()
    year = year or now.year
    month = month or now.month
    if year < 2000 or year > 2100:
        raise HTTPException(status_code=400, detail="Anio invalido.")
    if month < 1 or month > 12:
        raise HTTPException(status_code=400, detail="Mes invalido.")
    with SessionLocal() as session:
        start, end = month_bounds(year, month)
        portal = portal_for_report_period(session, latest_portal_payload(session), start, end)
        diesel_data = diesel_payload(session, start, end)
    data = monthly_report_pptx_bytes(portal, year, month, diesel_data)
    month_name = MONTH_NAMES_ES_FULL[month - 1]
    filename = re.sub(r"[^A-Za-z0-9_.-]+", "_", f"Reporte_Mensual_{month_name}_{year}.pptx")
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store, max-age=0",
        },
    )


@app.get("/api/weekly-report/powerpoint")
def get_weekly_report_powerpoint(
    start: str = Query(default=""),
    end: str = Query(default=""),
    base: str = Query(default=""),
) -> StreamingResponse:
    if start and end:
        start_date = parse_report_date(start, "start")
        end_date = parse_report_date(end, "end")
    else:
        base_date = parse_report_date(base, "base") if base else utc_now().date()
        start_date, end_date = week_bounds(base_date)
    if end_date < start_date:
        raise HTTPException(status_code=400, detail="La fecha final no puede ser menor que la fecha inicial.")
    if (end_date - start_date).days > 13:
        raise HTTPException(status_code=400, detail="El reporte semanal permite maximo 14 dias.")
    with SessionLocal() as session:
        start_text, end_text = start_date.isoformat(), end_date.isoformat()
        portal = portal_for_report_period(session, latest_portal_payload(session), start_text, end_text)
        diesel_data = diesel_payload(session, start_text, end_text)
    data = weekly_report_pptx_bytes(portal, start_date, end_date, diesel_data)
    filename = re.sub(
        r"[^A-Za-z0-9_.-]+",
        "_",
        f"Reporte_Semanal_{start_date.isoformat()}_{end_date.isoformat()}.pptx",
    )
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store, max-age=0",
        },
    )


@app.get("/api/hose-changes")
def get_hose_changes(
    response: Response,
    start: str = Query(default=""),
    end: str = Query(default=""),
    equipment: str = Query(default=""),
) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        return hose_report(session, start, end, equipment)


@app.post("/api/hose-changes")
async def save_hose_change_cloud(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Cambio de manguera invalido.")
    change_date = hose_iso_or_none(payload.get("change_date")) or utc_now().date().isoformat()
    equipment = normalize_text(payload.get("equipment"))
    if not equipment:
        raise HTTPException(status_code=400, detail="Equipo requerido.")
    record_id = int(parse_float(payload.get("id"), 0) or 0)
    with SessionLocal() as session:
        row = session.get(HoseChange, record_id) if record_id else None
        if row is None:
            row = HoseChange(change_date=change_date, equipment=equipment, created_at=utc_now())
            session.add(row)
        payload = {**payload, "change_date": change_date, "equipment": equipment, "source": "web"}
        row.external_id = str(payload.get("external_id") or row.external_id or "").strip()
        row.change_date = change_date
        row.equipment = equipment
        row.system = normalize_text(payload.get("system"))
        row.part_type = normalize_text(payload.get("part_type") or "MANGUERA")[:80] or "MANGUERA"
        row.diameter = normalize_text(payload.get("diameter"))[:80]
        row.length_m = max(parse_float(payload.get("length_m"), 0), 0)
        row.quantity = max(parse_float(payload.get("quantity"), 1), 0)
        row.unit_cost = max(parse_float(payload.get("unit_cost"), 0), 0)
        row.estimated_life_days = max(parse_float(payload.get("estimated_life_days"), 30), 0)
        row.estimated_weekly_qty = max(parse_float(payload.get("estimated_weekly_qty"), 0), 0)
        row.failure_reason = normalize_text(payload.get("failure_reason"))[:220]
        row.technician = normalize_text(payload.get("technician"))[:180]
        row.notes = str(payload.get("notes") or "").strip()
        row.source = "web"
        row.updated_at = utc_now()
        session.commit()
        session.refresh(row)
        return {"ok": True, "record": hose_change_payload(row), "hoses": hose_report(session, change_date, change_date)}


@app.post("/api/hose-changes/delete")
async def delete_hose_change_cloud(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Cambio de manguera invalido.")
    record_id = int(parse_float(payload.get("id"), 0) or 0)
    if not record_id:
        raise HTTPException(status_code=400, detail="Selecciona un cambio para eliminar.")
    with SessionLocal() as session:
        row = session.get(HoseChange, record_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Cambio no encontrado.")
        session.delete(row)
        session.commit()
        return {"ok": True}


@app.get("/api/diesel")
def get_diesel(
    response: Response,
    start: str = Query(default=""),
    end: str = Query(default=""),
    meta_lh: float = Query(default=0),
) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        return diesel_payload(session, start, end, meta_lh or None)


@app.get("/api/diesel/export")
def get_diesel_export(
    start: str = Query(default=""),
    end: str = Query(default=""),
    meta_lh: float = Query(default=0),
) -> StreamingResponse:
    with SessionLocal() as session:
        payload = diesel_payload(session, start, end, meta_lh or None)
    data = diesel_workbook_bytes(payload)
    filename = f"Control_Diesel_{payload.get('start','')}_{payload.get('end','')}.xlsx".replace(" ", "_")
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/api/diesel/records")
async def save_diesel_record_cloud(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Captura diesel invalida.")
    work_date = iso_date(payload.get("work_date"))
    equipment = normalize_text(payload.get("equipment"))
    if not equipment:
        raise HTTPException(status_code=400, detail="Equipo requerido.")
    shift = normalize_text(payload.get("shift") or "1")
    hi = parse_float(payload.get("horometer_initial"), 0)
    hf = parse_float(payload.get("horometer_final"), 0)
    worked = parse_float(payload.get("worked_hours"), -1)
    if worked < 0:
        worked = max(hf - hi, 0) if hi and hf and hf >= hi else 0
    with SessionLocal() as session:
        equipment = diesel_canonical_equipment(equipment, diesel_equipment_alias_map(latest_portal_payload(session))) or equipment
        record_id = int(parse_float(payload.get("id"), 0) or 0)
        row = session.get(DieselRecord, record_id) if record_id else None
        if row is None:
            row = session.scalar(
                select(DieselRecord).where(
                    DieselRecord.work_date == work_date,
                    DieselRecord.equipment == equipment,
                    DieselRecord.shift == shift,
                )
            )
        if row is None:
            row = DieselRecord(work_date=work_date, equipment=equipment, shift=shift, created_at=utc_now())
            session.add(row)
        row.work_date = work_date
        row.equipment = equipment
        row.condition = normalize_text(payload.get("condition") or "DISPONIBLE")
        row.shift = shift
        row.horometer_initial = hi
        row.horometer_final = hf
        row.worked_hours = max(worked, 0)
        row.diesel_liters = max(parse_float(payload.get("diesel_liters"), 0), 0)
        row.operator = normalize_text(payload.get("operator"))
        row.dispatcher = normalize_text(payload.get("dispatcher"))
        row.supervisor = normalize_text(payload.get("supervisor"))
        row.notes = str(payload.get("notes") or "").strip()
        row.source = "web"
        row.updated_at = utc_now()
        unmark_diesel_record_deleted(
            session,
            {"work_date": row.work_date, "equipment": row.equipment, "shift": row.shift},
        )
        session.commit()
        session.refresh(row)
        return {"ok": True, "record": diesel_record_payload(row), "diesel": diesel_payload(session, work_date, work_date)}


@app.delete("/api/diesel/records/{record_id}")
def delete_diesel_record_cloud(record_id: int, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        row = session.get(DieselRecord, record_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Captura diesel no encontrada.")
        mark_diesel_record_deleted(
            session,
            {"work_date": row.work_date, "equipment": row.equipment, "shift": row.shift},
        )
        session.delete(row)
        session.commit()
        return {"ok": True}


@app.post("/api/diesel/records/delete")
async def delete_diesel_record_by_key_cloud(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Captura diesel invalida.")
    work_date = iso_date(payload.get("work_date"))
    equipment = normalize_text(payload.get("equipment"))
    shift = normalize_text(payload.get("shift") or "1")
    if not equipment:
        raise HTTPException(status_code=400, detail="Equipo requerido.")
    with SessionLocal() as session:
        equipment = diesel_canonical_equipment(equipment, diesel_equipment_alias_map(latest_portal_payload(session))) or equipment
        row = None
        record_id = int(parse_float(payload.get("id"), 0) or 0)
        if record_id:
            row = session.get(DieselRecord, record_id)
        if row is None:
            row = session.scalar(
                select(DieselRecord).where(
                    DieselRecord.work_date == work_date,
                    DieselRecord.equipment == equipment,
                    DieselRecord.shift == shift,
                )
            )
        mark_diesel_record_deleted(session, {"work_date": work_date, "equipment": equipment, "shift": shift})
        if row is not None:
            session.delete(row)
        session.commit()
        return {"ok": True}


@app.post("/api/diesel/days")
async def save_diesel_day_cloud(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Control diario invalido.")
    work_date = iso_date(payload.get("work_date"))
    with SessionLocal() as session:
        row = session.scalar(select(DieselDay).where(DieselDay.work_date == work_date))
        if row is None:
            row = DieselDay(work_date=work_date, created_at=utc_now())
            session.add(row)
        row.work_date = work_date
        row.diesel_received = max(parse_float(payload.get("diesel_received"), 0), 0)
        row.initial_stock = max(parse_float(payload.get("initial_stock"), 0), 0)
        row.final_stock = max(parse_float(payload.get("final_stock"), 0), 0)
        row.prosermin_stock = max(parse_float(payload.get("prosermin_stock"), 0), 0)
        row.supplier = normalize_text(payload.get("supplier"))
        row.notes = str(payload.get("notes") or "").strip()
        row.source = "web"
        row.updated_at = utc_now()
        unmark_diesel_day_deleted(session, work_date)
        session.commit()
        session.refresh(row)
        return {"ok": True, "day": diesel_day_payload(row), "diesel": diesel_payload(session, work_date, work_date)}


@app.post("/api/diesel/days/delete")
async def delete_diesel_day_cloud(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Control diario invalido.")
    work_date = iso_date(payload.get("work_date"))
    with SessionLocal() as session:
        row = session.scalar(select(DieselDay).where(DieselDay.work_date == work_date))
        mark_diesel_day_deleted(session, work_date)
        if row is not None:
            session.delete(row)
        session.commit()
        return {"ok": True}


WAREHOUSE_HTML = r"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MGA Almacen de filtros</title>
  <style>
    :root { --blue:#2563eb; --blue2:#0ea5e9; --navy:#0f172a; --teal:#14b8a6; --green:#22c55e; --amber:#f59e0b; --red:#ef4444; --muted:#64748b; --line:#d7e0ea; --bg:#f8fafc; --panel:#ffffff; --soft:#f8fafc; --shadow:0 18px 42px rgba(15,23,42,.10); --deep:#07162f; --steel:#334155; --cyan:#22d3ee; }
    * { box-sizing:border-box; }
    body { margin:0; font-family:Segoe UI, Arial, sans-serif; color:#1f2937; background:
      radial-gradient(circle at 9% 8%, rgba(34,211,238,.18), transparent 26%),
      radial-gradient(circle at 90% 0%, rgba(20,184,166,.18), transparent 25%),
      linear-gradient(180deg,#f8fafc 0%,#eef5f8 100%); }
    body::before { content:""; position:fixed; inset:0; z-index:-1; opacity:.32; pointer-events:none; background:
      linear-gradient(90deg, rgba(15,23,42,.045) 1px, transparent 1px),
      linear-gradient(180deg, rgba(15,23,42,.045) 1px, transparent 1px);
      background-size:34px 34px; mask-image:linear-gradient(180deg,#000 0%,transparent 74%); }
    .hero { position:relative; overflow:hidden; color:var(--navy); padding:18px 28px 22px; display:grid; grid-template-columns:minmax(330px,1fr) minmax(360px,.9fr) minmax(280px,.34fr); gap:22px; align-items:center; background:
      linear-gradient(135deg,rgba(255,255,255,.97) 0%,rgba(248,251,255,.94) 42%,rgba(237,253,248,.92) 100%),
      radial-gradient(circle at 75% 25%, rgba(34,211,238,.22), transparent 36%);
      border-bottom:1px solid var(--line); box-shadow:0 18px 44px rgba(15,23,42,.12); }
    .hero::before { content:""; position:absolute; inset:0 0 auto; height:5px; background:linear-gradient(90deg,var(--blue2),var(--teal),var(--green),var(--amber)); }
    .hero::after { content:""; position:absolute; right:0; top:5px; width:46%; height:100%; opacity:.38; background:linear-gradient(90deg, transparent, rgba(20,184,166,.08)), repeating-linear-gradient(135deg, rgba(37,99,235,.14) 0 1px, transparent 1px 18px); pointer-events:none; }
    .brand { position:relative; z-index:1; display:flex; align-items:center; gap:18px; min-width:0; }
    .corner-logo { width:118px; height:78px; object-fit:contain; flex:0 0 auto; padding:8px 10px; border-radius:8px; background:white; border:1px solid var(--line); box-shadow:0 14px 28px rgba(15,23,42,.12); }
    header h1 { margin:0; font-size:28px; line-height:1.05; letter-spacing:0; }
    header p { margin:7px 0 0; color:var(--muted); font-size:14px; }
    .hero-visual { position:relative; z-index:1; min-height:142px; display:grid; grid-template-columns:1fr 1.2fr; gap:10px; align-items:stretch; }
    .ops-card, .ops-graph { border:1px solid rgba(215,224,234,.96); border-radius:8px; background:rgba(255,255,255,.82); box-shadow:0 12px 28px rgba(15,23,42,.08); }
    .ops-card { display:grid; align-content:center; gap:8px; padding:12px; }
    .ops-card span, .ops-graph span { color:var(--muted); font-size:11px; font-weight:800; text-transform:uppercase; }
    .ops-card b { color:var(--navy); font-size:26px; line-height:1; }
    .ops-card i { display:block; height:8px; overflow:hidden; border-radius:999px; background:#e8edf5; }
    .ops-card i::before { content:""; display:block; width:86%; height:100%; border-radius:999px; background:linear-gradient(90deg,var(--teal),var(--green)); animation:flowPulse 2.6s ease-in-out infinite; }
    .ops-graph { display:grid; gap:9px; padding:12px; }
    .ops-lines { display:grid; grid-template-columns:repeat(8,1fr); align-items:end; gap:5px; min-height:54px; padding-top:4px; }
    .ops-lines i { display:block; min-height:8px; border-radius:5px 5px 2px 2px; background:linear-gradient(180deg,var(--blue2),var(--teal)); opacity:.94; transform-origin:bottom; animation:barBreathe 2.8s ease-in-out infinite; }
    .ops-lines i:nth-child(3n) { animation-delay:.25s; }
    .ops-lines i:nth-child(4n) { animation-delay:.5s; }
    .ops-lines i:nth-child(2n) { background:linear-gradient(180deg,var(--green),var(--teal)); }
    .ops-legend { display:flex; gap:8px; color:var(--muted); font-size:11px; }
    .ops-legend b { display:inline-block; width:8px; height:8px; border-radius:2px; margin-right:4px; background:var(--teal); }
    .ops-legend span:last-child b { background:var(--amber); }
    .maintenance-video { grid-column:1 / -1; position:relative; min-height:58px; overflow:hidden; border:1px solid rgba(215,224,234,.96); border-radius:8px; background:
      linear-gradient(180deg,rgba(7,22,47,.94),rgba(15,23,42,.92)),
      repeating-linear-gradient(90deg,rgba(34,211,238,.15) 0 1px,transparent 1px 28px); box-shadow:0 16px 32px rgba(7,22,47,.16); }
    .maintenance-video::before { content:""; position:absolute; inset:12px 14px auto; height:3px; border-radius:999px; background:linear-gradient(90deg,transparent,var(--cyan),var(--teal),transparent); animation:scanLine 3.2s linear infinite; }
    .maintenance-video::after { content:""; position:absolute; inset:auto 0 0; height:22px; background:
      repeating-linear-gradient(90deg, rgba(245,158,11,.88) 0 18px, rgba(15,23,42,.4) 18px 32px); opacity:.38; transform:skewX(-16deg); }
    .machine-track { position:absolute; left:18px; right:18px; bottom:17px; height:4px; border-radius:999px; background:rgba(148,163,184,.45); }
    .machine { position:absolute; left:7%; bottom:20px; width:66px; height:28px; border-radius:5px 12px 5px 5px; background:linear-gradient(135deg,#f59e0b,#facc15); box-shadow:0 10px 18px rgba(0,0,0,.28); animation:machineMove 7s ease-in-out infinite; }
    .machine::before { content:""; position:absolute; right:8px; top:-11px; width:24px; height:15px; border-radius:6px 8px 0 0; background:linear-gradient(135deg,#38bdf8,#0ea5e9); }
    .machine::after { content:""; position:absolute; left:8px; right:8px; bottom:-8px; height:10px; border-radius:999px; background:radial-gradient(circle at 12px 5px,#0f172a 0 5px,transparent 6px), radial-gradient(circle at 46px 5px,#0f172a 0 5px,transparent 6px); }
    .maintenance-pulse { position:absolute; right:18px; top:12px; display:flex; gap:5px; color:#cffafe; font-size:10px; font-weight:900; letter-spacing:.08em; text-transform:uppercase; }
    .maintenance-pulse i { width:8px; height:8px; border-radius:50%; background:var(--green); box-shadow:0 0 0 0 rgba(34,197,94,.52); animation:statusPulse 1.6s infinite; }
    .key-card { position:relative; z-index:1; min-width:280px; padding:12px; border:1px solid var(--line); border-radius:8px; background:rgba(248,250,252,.88); box-shadow:inset 0 1px 0 rgba(255,255,255,.9), 0 12px 28px rgba(15,23,42,.07); }
    .key-card label { color:#334155; }
    header input { min-width:260px; padding:10px 11px; border:1px solid #cbd5e1; border-radius:6px; color:#172033; background:white; outline:none; }
    header input::placeholder { color:#94a3b8; }
    main { width:min(1480px, 100%); margin:0 auto; padding:18px; display:grid; gap:14px; }
    .tabs { position:sticky; top:0; z-index:10; display:flex; gap:6px; flex-wrap:wrap; padding:7px; border:1px solid rgba(215,224,234,.94); border-radius:12px; background:rgba(255,255,255,.88); box-shadow:0 16px 36px rgba(15,23,42,.10); backdrop-filter:blur(14px); }
    .tabs button, .btn { border:0; background:var(--blue); color:white; padding:10px 14px; border-radius:6px; font-weight:700; cursor:pointer; transition:transform .15s ease, box-shadow .15s ease, background .15s ease; }
    .tabs button:hover, .btn:hover { transform:translateY(-1px); box-shadow:0 10px 20px rgba(39,58,92,.14); }
    .tabs button { background:linear-gradient(180deg,#f4f7fb,#e9f0f7); color:#263447; border:1px solid transparent; }
    .tabs button.active { background:#ffffff; color:var(--blue); border-color:#b7c8e8; box-shadow:inset 0 -3px 0 var(--teal), 0 10px 20px rgba(39,58,92,.10); }
    .btn.secondary { background:white; color:var(--blue); border:1px solid var(--line); }
    .btn.danger { background:linear-gradient(135deg,#a91d2c,var(--red)); }
    .btn.small { padding:5px 8px; border-radius:5px; font-size:11px; white-space:nowrap; }
    .panel { position:relative; overflow:hidden; background:rgba(255,255,255,.97); border:1px solid rgba(215,224,234,.96); border-radius:12px; padding:16px; box-shadow:var(--shadow); }
    .panel::before { content:""; position:absolute; inset:0 0 auto; height:3px; background:linear-gradient(90deg,var(--blue2),var(--teal)); opacity:.86; }
    .toolbar { display:grid; grid-template-columns:repeat(5, minmax(140px, 1fr)); gap:10px; align-items:end; }
    .kpi-sim-toolbar { grid-template-columns:repeat(auto-fit, minmax(120px, 1fr)); align-items:end; border-color:#f6d365; background:#fffdf4; }
    .kpi-sim-toolbar::before { background:linear-gradient(90deg,#f59e0b,#facc15); }
    .kpi-sim-toolbar input[type="number"] { min-width:0; }
    .kpi-sim-note { align-self:center; font-size:12px; font-weight:700; }
    #kpiPrintArea.simulation::before { background:linear-gradient(90deg,#f59e0b,#facc15); }
    .kpi-simulation-badge { display:inline-block; margin-left:10px; padding:3px 8px; border:1px solid #d97706; border-radius:6px; background:#fff3cd; color:#7a4d00; font-size:12px; font-weight:800; vertical-align:middle; }
    .kpi-main-strip { display:grid; grid-template-columns:repeat(6,minmax(150px,1fr)); gap:10px; }
    .kpi-main-card { position:relative; overflow:hidden; min-height:120px; border:1px solid var(--line); border-radius:14px; padding:14px; background:linear-gradient(135deg,#fff,#f8fbff); box-shadow:0 12px 28px rgba(15,23,42,.08); }
    .kpi-main-card::before { content:""; position:absolute; inset:0 0 auto; height:4px; background:var(--teal); }
    .kpi-main-card.warn::before { background:var(--amber); }
    .kpi-main-card.bad::before { background:var(--red); }
    .kpi-main-card span { display:block; color:#64748b; font-size:12px; font-weight:900; text-transform:uppercase; }
    .kpi-main-card strong { display:block; margin-top:8px; color:var(--navy); font-size:32px; line-height:1; }
    .kpi-main-card small { display:block; margin-top:8px; color:#475569; font-size:12px; line-height:1.25; }
    .kpi-main-meter { height:7px; overflow:hidden; border-radius:999px; background:#e5e7eb; margin-top:11px; }
    .kpi-main-meter i { display:block; height:100%; border-radius:999px; background:linear-gradient(90deg,var(--teal),var(--green)); }
    .kpi-main-card.warn .kpi-main-meter i { background:linear-gradient(90deg,#f59e0b,#facc15); }
    .kpi-main-card.bad .kpi-main-meter i { background:linear-gradient(90deg,#ef4444,#b91c1c); }
    label { display:grid; gap:4px; color:#344054; font-size:12px; font-weight:700; }
    input, select, textarea { width:100%; padding:9px 10px; border:1px solid #cbd5e1; border-radius:6px; font:inherit; background:white; outline:none; transition:border .15s ease, box-shadow .15s ease; }
    .inline-check { display:flex; align-items:center; gap:8px; min-height:38px; }
    .inline-check input { width:auto; }
    input:focus, select:focus, textarea:focus { border-color:var(--teal); box-shadow:0 0 0 3px rgba(0,156,154,.14); }
    .stats { display:grid; grid-template-columns:repeat(auto-fit, minmax(150px, 1fr)); gap:10px; }
    .stat { position:relative; overflow:hidden; min-height:92px; display:grid; grid-template-columns:auto 1fr; align-items:start; gap:8px 10px; background:linear-gradient(180deg,#fff,#f8fbff); border:1px solid var(--line); padding:14px 15px; border-radius:12px; box-shadow:0 12px 28px rgba(15,23,42,.07); transition:transform .18s ease, box-shadow .18s ease; }
    .stat:hover { transform:translateY(-2px); box-shadow:0 18px 34px rgba(15,23,42,.11); }
    .stat::before { content:""; position:absolute; left:0; top:0; bottom:0; width:4px; background:linear-gradient(180deg,var(--teal),var(--green)); }
    .stat-icon { position:relative; width:38px; height:38px; border-radius:8px; background:linear-gradient(135deg,rgba(37,99,235,.16),rgba(20,184,166,.18)); border:1px solid rgba(37,99,235,.16); }
    .stat-icon::after { content:""; position:absolute; left:9px; right:9px; top:18px; height:9px; border-left:3px solid var(--blue); border-right:3px solid var(--teal); border-bottom:3px solid var(--green); border-radius:0 0 5px 5px; }
    .stat strong { display:block; color:var(--blue); font-size:30px; line-height:1; margin-bottom:5px; }
    .stat span:not(.stat-icon) { color:var(--muted); font-weight:700; }
    .stat-spark { grid-column:1 / -1; display:flex; align-items:end; gap:4px; height:24px; padding-left:5px; }
    .stat-spark i { flex:1; min-width:3px; border-radius:4px 4px 1px 1px; background:linear-gradient(180deg,var(--blue2),var(--teal)); opacity:.82; }
    .executive-board { background:linear-gradient(135deg,rgba(255,255,255,.98),rgba(241,245,249,.96)); }
    .exec-alert-grid { display:grid; grid-template-columns:repeat(6,minmax(120px,1fr)); gap:10px; }
    .exec-card { position:relative; overflow:hidden; min-height:104px; padding:13px; border:1px solid var(--line); border-radius:12px; background:white; box-shadow:0 10px 24px rgba(15,23,42,.07); cursor:pointer; transition:transform .16s ease, box-shadow .16s ease; }
    .exec-card:hover { transform:translateY(-2px); box-shadow:0 16px 30px rgba(15,23,42,.11); }
    .exec-card::before { content:""; position:absolute; inset:0 auto 0 0; width:5px; background:var(--teal); }
    .exec-card.bad::before { background:var(--red); }
    .exec-card.warn::before { background:var(--amber); }
    .exec-card strong { display:block; color:var(--navy); font-size:30px; line-height:1; }
    .exec-card span { display:block; margin-top:6px; color:#475569; font-size:12px; font-weight:900; text-transform:uppercase; }
    .exec-card small { display:block; margin-top:7px; color:var(--muted); font-size:12px; line-height:1.25; }
    .exec-alert-list { display:grid; gap:7px; margin-top:12px; }
    .exec-alert { display:flex; gap:9px; align-items:center; padding:9px 11px; border:1px solid var(--line); border-radius:10px; background:#f8fafc; color:#334155; font-size:13px; }
    .exec-alert b { flex:0 0 auto; min-width:82px; color:white; text-align:center; padding:3px 8px; border-radius:999px; font-size:11px; background:var(--teal); }
    .exec-alert.bad b { background:var(--red); }
    .exec-alert.warn b { background:var(--amber); color:#4a2b00; }
    .dashboard-command { background:linear-gradient(135deg,rgba(255,255,255,.98),rgba(239,246,255,.96)); }
    .kpi-command-grid { display:grid; grid-template-columns:minmax(360px,.9fr) minmax(520px,1.1fr); gap:14px; align-items:start; }
    .priority-list { display:grid; gap:8px; }
    .priority-row { display:grid; grid-template-columns:42px 1fr auto; gap:10px; align-items:center; padding:10px 11px; border:1px solid var(--line); border-radius:12px; background:white; cursor:pointer; box-shadow:0 8px 18px rgba(15,23,42,.05); }
    .priority-row:hover { background:#f0fdfa; transform:translateY(-1px); }
    .priority-rank { display:grid; place-items:center; width:32px; height:32px; border-radius:10px; color:white; font-weight:900; background:var(--teal); }
    .priority-row.bad .priority-rank { background:var(--red); }
    .priority-row.warn .priority-rank { background:var(--amber); color:#4a2b00; }
    .priority-row strong { display:block; color:var(--navy); font-size:14px; }
    .priority-row span { display:block; color:#64748b; font-size:12px; margin-top:2px; }
    .priority-score { font-weight:900; color:#0f172a; }
    .semaphore-dot { display:inline-block; width:12px; height:12px; border-radius:50%; box-shadow:0 0 0 3px rgba(15,23,42,.06); vertical-align:middle; }
    .sem-green { background:#22c55e; }
    .sem-yellow { background:#f59e0b; }
    .sem-red { background:#ef4444; }
    .sem-gray { background:#94a3b8; }
    .meeting-mode header, .meeting-mode .tabs, .meeting-mode #stats, .meeting-mode .dashboard-controls, .meeting-mode .kpi-sim-toolbar { display:none !important; }
    .meeting-mode main { width:100%; max-width:1600px; padding:12px; }
    .meeting-mode .no-print { display:grid; }
    .meeting-mode #dashboard { gap:10px; }
    .meeting-mode #kpiPrintArea { box-shadow:none; }
    .profile-hero { display:grid; grid-template-columns:minmax(260px,.82fr) minmax(420px,1.18fr); gap:14px; align-items:stretch; }
    .profile-card { border:1px solid var(--line); border-radius:12px; background:linear-gradient(135deg,#ffffff,#f8fbff); padding:15px; box-shadow:0 12px 26px rgba(15,23,42,.07); }
    .profile-card h3 { margin:0 0 8px; color:var(--navy); font-size:24px; }
    .profile-card p { margin:4px 0; color:#475569; }
    .profile-badges { display:flex; flex-wrap:wrap; gap:7px; margin-top:10px; }
    .profile-badge { display:inline-flex; align-items:center; gap:6px; padding:5px 9px; border-radius:999px; background:#e0f2fe; color:#075985; font-size:12px; font-weight:900; }
    .profile-badge.bad { background:#fee2e2; color:#991b1b; }
    .profile-badge.warn { background:#fef3c7; color:#92400e; }
    .profile-grid { display:grid; grid-template-columns:repeat(6,minmax(120px,1fr)); gap:10px; }
    .profile-metric { min-height:92px; border:1px solid var(--line); border-radius:12px; padding:12px; background:white; box-shadow:0 10px 22px rgba(15,23,42,.06); }
    .profile-metric strong { display:block; color:var(--blue); font-size:26px; line-height:1; }
    .profile-metric span { display:block; margin-top:7px; color:#475569; font-size:12px; font-weight:900; text-transform:uppercase; }
    .profile-section-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; align-items:start; }
    .profile-actions { display:flex; flex-wrap:wrap; gap:8px; margin-top:12px; }
    .view { display:none; }
    .view.active { display:grid; gap:14px; }
    table { width:100%; border-collapse:separate; border-spacing:0; background:white; }
    th, td { border-bottom:1px solid var(--line); padding:9px 10px; font-size:13px; vertical-align:top; }
    th { background:linear-gradient(180deg,#edf3f9,#e4edf7); color:#243042; position:sticky; top:0; z-index:1; text-transform:uppercase; font-size:12px; }
    tbody tr:nth-child(even) { background:#fbfdff; }
    tbody tr:hover { background:#edf8f7; }
    .table-wrap { max-height:620px; overflow:auto; border:1px solid var(--line); border-radius:8px; background:white; box-shadow:inset 0 1px 0 rgba(255,255,255,.8); }
    .pill { display:inline-block; padding:2px 7px; border-radius:999px; font-weight:700; font-size:12px; }
    .ok { color:#047857; background:#d1fae5; }
    .bad { color:#b91c1c; background:#fee2e2; }
    .warn { color:#92400e; background:#fef3c7; }
    .muted { color:var(--muted); }
    .grid2 { display:grid; grid-template-columns:1.1fr .9fr; gap:14px; align-items:start; }
    .capture-layout { grid-template-columns:minmax(0,1fr) minmax(560px,.9fr); }
    .capture-form-grid { display:grid; grid-template-columns:repeat(4,minmax(120px,1fr)); gap:10px; }
    .capture-section-title { grid-column:1 / -1; margin:4px 0 -2px; padding:8px 10px; border-left:4px solid var(--teal); border-radius:4px; background:#f0fdfa; color:#0f766e; font-size:12px; font-weight:900; text-transform:uppercase; letter-spacing:.04em; }
    .capture-fluid { background:#f8fafc; border-radius:6px; padding:8px; margin:-2px; }
    .capture-actions { position:sticky; bottom:0; z-index:3; padding:10px 0 2px; background:linear-gradient(180deg,rgba(255,255,255,.82),#fff 30%); }
    .capture-actions .btn { min-height:40px; }
    .quick-alerts { display:grid; gap:7px; margin-top:10px; }
    .quick-alert { padding:8px 10px; border-radius:9px; border:1px solid #fed7aa; background:#fff7ed; color:#9a3412; font-size:12px; font-weight:800; }
    .input-warning { border-color:#f59e0b !important; box-shadow:0 0 0 3px rgba(245,158,11,.16) !important; }
    .latest-captures-panel { min-width:0; position:sticky; top:12px; }
    .latest-captures-wrap { max-height:clamp(320px, calc(100vh - 395px), 620px); overflow:auto; }
    #capRecentTable { min-width:900px; }
    #capRecentTable th, #capRecentTable td { padding:7px 8px; font-size:12px; line-height:1.2; }
    #capRecentTable th { font-size:11px; white-space:nowrap; }
    #capRecentTable td { white-space:nowrap; }
    #capRecentTable td:nth-child(1), #capRecentTable td:nth-child(2), #capRecentTable td:nth-child(11) { white-space:normal; }
    #capRecentTable td:nth-child(12) { width:72px; }
    #capRecentTable td:nth-child(12) .btn { display:block; width:100%; margin:0 0 5px; padding:5px 7px; font-size:11px; }
    #capRecentTable td:nth-child(12) .btn:last-child { margin-bottom:0; }
    .tire-track-layout { grid-template-columns:minmax(430px,.75fr) minmax(620px,1.25fr); }
    .tire-track-form { display:grid; grid-template-columns:repeat(3,minmax(110px,1fr)); gap:10px; }
    .tire-track-form .wide { grid-column:1 / -1; }
    .tire-kpi-short { display:grid; grid-template-columns:repeat(4,minmax(110px,1fr)); gap:8px; margin-bottom:10px; }
    .tire-kpi-short span { display:grid; gap:3px; padding:9px 10px; border:1px solid var(--line); border-radius:7px; background:#f8fafc; color:#475569; font-size:11px; font-weight:800; text-transform:uppercase; }
    .tire-kpi-short b { color:var(--blue); font-size:22px; line-height:1; }
    .movement-grid { display:grid; grid-template-columns:repeat(4, 1fr); gap:10px; }
    .req-header-grid { display:grid; grid-template-columns:repeat(3, 1fr); gap:10px; }
    .req-item-grid { display:grid; grid-template-columns:110px 150px 1fr 1.6fr; gap:10px; align-items:end; }
    .req-actions { display:flex; gap:8px; flex-wrap:wrap; margin-top:10px; }
    .wide { grid-column:1 / -1; }
    .dashboard-grid { display:grid; grid-template-columns:repeat(4, 1fr); gap:10px; }
    .metric-card { display:grid; grid-template-columns:minmax(0,1fr) 62px; align-items:center; gap:10px; border:1px solid var(--line); border-radius:12px; padding:13px; background:linear-gradient(180deg,#fff,#f8fbff); box-shadow:0 10px 24px rgba(15,23,42,.06); transition:transform .18s ease, box-shadow .18s ease; }
    .metric-card:hover { transform:translateY(-2px); box-shadow:0 16px 30px rgba(15,23,42,.10); }
    .metric-card span { display:block; color:var(--muted); font-size:12px; font-weight:800; text-transform:uppercase; }
    .metric-card strong { display:block; color:var(--blue); font-size:30px; margin-top:5px; }
    .metric-card .bar-track { height:8px; border-radius:999px; background:#e5e7eb; margin-top:10px; overflow:hidden; }
    .metric-card .bar-fill { display:block; height:100%; background:linear-gradient(90deg,var(--teal),var(--green)); }
    .metric-card.bad .bar-fill { background:var(--red); }
    .metric-ring { position:relative; width:58px; height:58px; border-radius:50%; display:grid; place-items:center; background:conic-gradient(var(--teal) var(--ring), #e8edf5 0deg); box-shadow:inset 0 0 0 1px rgba(15,23,42,.05); }
    .metric-ring b { display:grid; place-items:center; width:42px; height:42px; border-radius:50%; background:white; color:var(--navy); font-size:11px; box-shadow:0 2px 6px rgba(15,23,42,.08); }
    .metric-card.bad .metric-ring { background:conic-gradient(var(--red) var(--ring), #e8edf5 0deg); }
    .metric-body { min-width:0; }
    .kpi-format-board { display:grid; grid-template-columns:minmax(260px,.82fr) minmax(430px,1.36fr) minmax(260px,.82fr); gap:12px; align-items:stretch; }
    .kpi-side { display:grid; grid-template-columns:1fr 1fr; gap:0; align-self:stretch; border:1px solid var(--line); background:white; }
    .kpi-side .metric-card { min-height:126px; border-radius:0; border:0; border-right:1px solid var(--line); border-bottom:1px solid var(--line); box-shadow:none; background:#fff; }
    .kpi-side .metric-card:nth-child(2n) { border-right:0; }
    .kpi-side .metric-card:nth-last-child(-n+2) { border-bottom:0; }
    .kpi-side .metric-card strong { color:#5f6671; font-size:30px; text-align:center; }
    .kpi-side .metric-card span { text-align:center; color:#667085; font-size:13px; }
    .kpi-side .metric-card small { display:flex; justify-content:space-between; gap:8px; margin-top:10px; color:#5f6671; }
    .kpi-special-mode .kpi-format-board { display:block; }
    .kpi-special-mode .kpi-side { border:0; background:transparent; }
    .kpi-special-mode #kpiCards { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; }
    .kpi-special-mode #kpiSideCards { display:none; }
    .kpi-diesel-mode { background:#f4f7fb; border-color:#dbe3ef; color:#172033; }
    .kpi-diesel-mode > .subtle-title { align-items:flex-start; margin-bottom:14px; padding:0 2px; }
    .kpi-diesel-mode #kpiTitle { color:#071f49; font-size:22px; line-height:1.15; }
    .kpi-diesel-mode #portalUpdated { color:#667085; font-size:12px; padding-top:4px; }
    .kpi-diesel-mode .kpi-format-board { display:grid; grid-template-columns:1fr; gap:12px; }
    .kpi-diesel-mode .kpi-side { display:block; border:0; background:transparent; }
    .kpi-diesel-mode #kpiCards, .diesel-card-grid { display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:10px; }
    .kpi-diesel-mode #kpiSideCards { display:none; }
    .kpi-diesel-mode .chart { display:block; min-height:0; padding:0; border:0; border-radius:0; background:transparent; overflow:visible; }
    .kpi-diesel-mode .kpi-report-table { margin-top:12px; max-height:none; overflow:visible; border:0; background:transparent; }
    .diesel-card { position:relative; min-height:122px; overflow:hidden; border:1px solid #dbe3ef; border-radius:8px; padding:14px 14px 12px; background:linear-gradient(180deg,#fff,#f8fbff); box-shadow:0 12px 26px rgba(7,31,73,.07); }
    .diesel-card::before { content:""; position:absolute; inset:0 auto 0 0; width:4px; background:#009c9a; }
    .diesel-card.is-bad::before { background:#c81e1e; }
    .diesel-card span { display:block; color:#667085; font-size:11px; font-weight:800; letter-spacing:0; text-transform:uppercase; }
    .diesel-card strong { display:block; margin-top:8px; color:#071f49; font-size:25px; line-height:1; }
    .diesel-card small { display:block; min-height:28px; margin-top:7px; color:#475569; font-size:12px; line-height:1.2; }
    .diesel-meter { height:7px; margin-top:10px; overflow:hidden; border-radius:999px; background:#e8eef6; }
    .diesel-meter i { display:block; height:100%; border-radius:999px; background:linear-gradient(90deg,#009c9a,#18b7a6); }
    .diesel-card.is-bad .diesel-meter i { background:linear-gradient(90deg,#e11d48,#c81e1e); }
    .diesel-visual-grid { display:grid; grid-template-columns:1.45fr .95fr; gap:12px; align-items:stretch; }
    .diesel-panel { min-width:0; border:1px solid #dbe3ef; border-radius:8px; padding:14px; background:white; box-shadow:0 12px 26px rgba(7,31,73,.06); }
    .diesel-panel-head { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; margin-bottom:12px; }
    .diesel-panel-head span { color:#071f49; font-size:15px; font-weight:800; }
    .diesel-panel-head b { color:#667085; font-size:11px; font-weight:800; text-transform:uppercase; }
    .diesel-bars-list { display:grid; gap:8px; }
    .diesel-bar-row { display:grid; grid-template-columns:132px minmax(110px,1fr) 76px 70px; gap:10px; align-items:center; min-height:32px; }
    .diesel-bar-label { min-width:0; }
    .diesel-bar-label b { display:block; overflow:hidden; color:#071f49; font-size:13px; text-overflow:ellipsis; white-space:nowrap; }
    .diesel-bar-label span { display:block; overflow:hidden; color:#667085; font-size:10px; font-weight:800; text-overflow:ellipsis; text-transform:uppercase; white-space:nowrap; }
    .diesel-bar-track { height:12px; overflow:hidden; border-radius:999px; background:#edf2f7; }
    .diesel-bar-track i { display:block; height:100%; min-width:3px; border-radius:999px; background:linear-gradient(90deg,#009c9a,#18b7a6); }
    .diesel-bar-row.is-bad .diesel-bar-track i { background:linear-gradient(90deg,#e11d48,#c81e1e); }
    .diesel-bar-row strong { color:#172033; font-size:12px; text-align:right; white-space:nowrap; }
    .diesel-bar-row em { color:#475569; font-size:11px; font-style:normal; text-align:right; white-space:nowrap; }
    .diesel-performance-grid { display:grid; gap:12px; }
    .diesel-target { border:1px solid #e2e8f0; border-radius:8px; padding:13px; background:#f8fafc; }
    .diesel-target-top { display:flex; align-items:end; justify-content:space-between; gap:10px; margin-bottom:11px; }
    .diesel-target-top span { color:#667085; font-size:11px; font-weight:800; text-transform:uppercase; }
    .diesel-target-top strong { color:#071f49; font-size:32px; line-height:1; white-space:nowrap; }
    .diesel-target-top small { color:#475569; font-size:12px; text-align:right; }
    .diesel-target-meter { position:relative; height:16px; overflow:hidden; border-radius:999px; background:#e8eef6; }
    .diesel-target-meter i { display:block; height:100%; min-width:3px; border-radius:999px; background:linear-gradient(90deg,#009c9a,#f59e0b,#e11d48); }
    .diesel-target-meter .diesel-target-marker { position:absolute; top:-5px; bottom:-5px; left:var(--target); width:2px; background:#071f49; box-shadow:0 0 0 2px rgba(255,255,255,.85); }
    .diesel-split { border:1px solid #e2e8f0; border-radius:8px; padding:13px; background:white; }
    .diesel-split h4 { margin:0 0 10px; color:#071f49; font-size:14px; }
    .diesel-split-track { display:flex; height:20px; overflow:hidden; border-radius:999px; background:#e8eef6; }
    .diesel-split-track i { display:block; height:100%; min-width:0; }
    .diesel-split-track .mga { background:#009c9a; }
    .diesel-split-track .pro { background:#f59e0b; }
    .diesel-split-legend { display:grid; gap:7px; margin-top:10px; }
    .diesel-split-legend > span { display:flex; align-items:center; justify-content:space-between; gap:10px; color:#475569; font-size:12px; }
    .diesel-split-legend em { font-style:normal; }
    .diesel-split-legend i { display:inline-block; width:10px; height:10px; border-radius:2px; margin-right:6px; vertical-align:-1px; }
    .diesel-split-legend .mga { background:#009c9a; }
    .diesel-split-legend .pro { background:#f59e0b; }
    .diesel-watch-list { display:grid; gap:7px; }
    .diesel-watch-item { display:flex; justify-content:space-between; gap:10px; padding:8px 10px; border-radius:7px; background:#f8fafc; color:#475569; font-size:12px; }
    .diesel-watch-item b { color:#071f49; }
    .diesel-empty { display:grid; min-height:210px; place-items:center; color:#667085; font-weight:700; }
    .diesel-table { width:100%; border-collapse:separate; border-spacing:0; overflow:hidden; border:1px solid #dbe3ef; border-radius:8px; background:white; color:#172033; }
    .diesel-table th { position:static; padding:9px 8px; background:#e8eef7; color:#243042; font-size:11px; text-align:center; text-transform:uppercase; }
    .diesel-table td { padding:8px 8px; border-bottom:1px solid #e5ebf3; color:#334155; font-size:12px; text-align:center; vertical-align:middle; }
    .diesel-table tbody tr:last-child td { border-bottom:0; }
    .diesel-table .diesel-eq { color:#071f49; font-weight:800; text-align:left; }
    .diesel-table .diesel-number { font-variant-numeric:tabular-nums; text-align:right; }
    .diesel-table .diesel-total td { background:#f0fdfa; color:#071f49; font-weight:800; }
    .kpi-oil-mode { background:#eeeeee; box-shadow:none; border-color:#d8d8d8; color:#333; }
    .kpi-oil-mode > .subtle-title { display:block; height:34px; margin:-2px -2px 12px; background:white; }
    .kpi-oil-mode #kpiTitle { display:grid; grid-template-columns:1fr 2fr 1fr; align-items:center; margin:0; height:34px; color:#333; text-align:center; font-size:20px; }
    .kpi-oil-mode #portalUpdated { display:none; }
    .kpi-oil-mode .kpi-format-board { grid-template-columns:minmax(260px,.7fr) minmax(430px,1.2fr) minmax(260px,.7fr); gap:22px; align-items:start; }
    .kpi-oil-mode .kpi-side { display:grid; grid-template-columns:1fr; gap:40px; border:0; background:transparent; align-self:start; }
    .oil-head { display:grid; grid-template-columns:1fr 2fr 1fr; align-items:center; height:34px; margin:-2px -2px 12px; background:white; color:#111; font-weight:800; text-align:center; }
    .oil-head h2 { margin:0; font-size:20px; color:#333; }
    .oil-month { font-size:14px; }
    .oil-metric-section { background:white; border:1px solid #e5e7eb; }
    .oil-metric-title { height:34px; display:flex; align-items:center; justify-content:center; color:#707780; font-weight:800; font-size:18px; }
    .oil-metric-grid { display:grid; grid-template-columns:1fr 1fr; gap:6px; background:#eeeeee; }
    .oil-metric-cell { min-height:92px; background:white; display:grid; align-content:center; justify-items:center; gap:8px; padding:8px 6px; }
    .oil-metric-cell strong { color:#777d86; font-size:31px; line-height:1; }
    .oil-metric-line { height:10px; width:100%; background:#eef1f4; }
    .oil-metric-line i { display:block; height:100%; width:100%; background:var(--teal); }
    .oil-metric-line.oil-red i { background:#d76f75; }
    .oil-metric-line.oil-darkred i { background:#a40000; }
    .oil-metric-cell span { color:#4b5563; font-size:12px; }
    .kpi-oil-mode .chart { min-height:330px; padding:14px 18px 10px; border-radius:0; background:white; display:block; overflow:hidden; }
    .oil-chart-grid { display:grid; grid-template-columns:42px 1fr; grid-template-rows:250px 38px; column-gap:8px; }
    .oil-axis { grid-row:1; display:flex; flex-direction:column; justify-content:space-between; align-items:end; padding:0 2px 0 0; color:#111; font-size:12px; }
    .oil-plot { position:relative; grid-column:2; grid-row:1; display:flex; align-items:stretch; gap:14px; padding:0 8px; border-bottom:1px solid #d9d9d9; background:repeating-linear-gradient(to top, transparent 0, transparent 49px, #d9d9d9 50px); }
    .oil-cluster { flex:1 1 62px; min-width:54px; display:grid; grid-template-rows:1fr auto; justify-items:center; gap:8px; }
    .oil-bars-stack { height:100%; display:flex; align-items:flex-end; gap:2px; }
    .oil-series-bar { width:9px; min-height:1px; position:relative; }
    .oil-series-bar b { position:absolute; left:50%; transform:translateX(-50%); top:-15px; color:#111; font-size:10px; font-weight:500; white-space:nowrap; }
    .oil-cluster-label { color:#111; font-size:12px; text-align:center; white-space:nowrap; }
    .oil-legend { grid-column:2; grid-row:2; display:flex; align-items:end; justify-content:center; gap:18px; color:#333; font-size:12px; }
    .oil-legend span { display:flex; align-items:center; gap:5px; white-space:nowrap; }
    .oil-legend i { display:block; width:10px; height:10px; }
    .oil-bottom-wrap { margin-top:28px; max-height:none; overflow:visible; border:0; background:transparent; }
    #kpiTable.oil-bottom-grid { border-collapse:separate; border-spacing:0; background:transparent; }
    #kpiTable.oil-bottom-grid > tbody > tr:hover { background:transparent; }
    #kpiTable.oil-bottom-grid > tbody > tr > td { border:0; padding:0 8px; vertical-align:top; }
    .oil-report-cell { width:72%; }
    .oil-order-cell { width:28%; }
    .oil-report-header { background:#f3f3f3; text-align:center; padding:10px 8px 8px; }
    .oil-report-header h3 { margin:0 0 8px; color:#111; font-size:17px; }
    .oil-days { display:flex; justify-content:center; gap:72px; color:#707780; font-size:12px; font-weight:800; }
    .oil-days b { display:inline-block; min-width:58px; margin-left:8px; padding:5px 16px; background:white; color:#111; }
    .oil-report-table { width:100%; border-collapse:collapse; background:white; color:#555; }
    .oil-report-table th, .oil-report-table td { border:1px solid #111; padding:4px 5px; font-size:10px; text-align:center; vertical-align:middle; }
    .oil-report-table th { position:static; background:white; color:#555; font-weight:800; text-transform:none; line-height:1.05; }
    .oil-report-table td { background:#efefef; }
    .oil-report-table .oil-subtotal td { background:#ffd966; }
    .oil-report-table .oil-total td { background:#fff200; }
    .oil-report-table .oil-zero { color:#f05b5b; font-weight:800; }
    .oil-order-panel { background:white; border:1px solid #cbd5e1; min-height:330px; }
    .oil-order-title { background:var(--blue2); color:white; text-align:center; font-weight:800; padding:13px 8px; }
    .oil-order-kpis { display:grid; grid-template-columns:repeat(3,1fr); border-bottom:1px solid #dbe3ef; }
    .oil-order-kpi { background:#f8fafc; border-right:1px solid #e2e8f0; text-align:center; padding:12px 4px 9px; }
    .oil-order-kpi:last-child { border-right:0; }
    .oil-order-kpi strong { display:block; color:#d6335c; font-size:12px; }
    .oil-order-kpi span { color:#475569; font-size:12px; }
    .oil-order-table { width:100%; border-collapse:collapse; }
    .oil-order-table th, .oil-order-table td { border-bottom:1px solid #e2e8f0; padding:7px 6px; font-size:11px; text-align:center; }
    .oil-order-table th { position:static; background:#e2e8f0; color:#0f172a; text-transform:none; }
    .oil-order-table td:first-child { text-align:left; font-weight:700; color:#334155; }
    .oil-order-table .oil-order-hot { color:#d6335c; font-weight:800; }
    .kpi-report-table { margin-top:14px; max-height:none; overflow:visible; }
    .chart { display:flex; align-items:end; gap:12px; min-height:270px; padding:20px 16px 28px; border:1px solid var(--line); border-radius:8px; background:linear-gradient(180deg,#fff,#f8fbff); overflow:auto; box-shadow:inset 0 1px 0 rgba(255,255,255,.9); }
    .kpi-format-mode .chart { display:block; min-height:330px; padding:12px 14px 18px; }
    .kpi-chart-head { display:flex; align-items:center; gap:10px; margin-bottom:12px; color:#111827; font-size:11px; }
    .kpi-mini-tabs { display:grid; grid-template-columns:repeat(4, minmax(96px, 1fr)); gap:4px; flex:1; }
    .kpi-mini-tabs span, .kpi-mini-tabs button { border:1px solid #111; padding:7px 9px; background:white; color:#111; font:inherit; font-size:12px; text-align:left; cursor:pointer; }
    .kpi-mini-tabs span.active, .kpi-mini-tabs button.active { background:var(--teal); color:#031b1b; }
    .chart-plot { position:relative; min-height:258px; display:flex; align-items:end; gap:12px; overflow:auto; padding:26px 8px 8px; background:linear-gradient(180deg,rgba(248,250,252,.72),rgba(255,255,255,.65)), repeating-linear-gradient(to top, transparent 0, transparent 51px, rgba(100,116,139,.22) 52px); border-radius:8px; }
    .chart-target { position:absolute; left:8px; right:8px; border-top:2px dashed rgba(245,158,11,.76); color:#92400e; font-size:11px; font-weight:800; text-align:right; pointer-events:none; }
    .chart-target b { background:#fff7ed; border:1px solid #fed7aa; border-radius:999px; padding:2px 8px; }
    .chart-bar { min-width:54px; display:grid; align-content:end; gap:6px; text-align:center; color:#344054; font-size:11px; }
    .chart-bar i { display:block; height:var(--h); min-height:4px; border-radius:7px 7px 2px 2px; background:linear-gradient(180deg,var(--green),var(--teal)); box-shadow:0 10px 20px rgba(20,184,166,.20); animation:growBar .72s ease both; }
    .chart-bar.out i { background:linear-gradient(180deg,#e35d6a,#b51f32); }
    .schedule-strip { display:flex; gap:8px; min-height:92px; padding:10px; overflow:auto; border:1px solid var(--line); border-radius:8px; background:#f8fafc; }
    .schedule-cell { min-width:74px; border:1px solid #dbe3ef; border-radius:7px; background:white; padding:7px; display:grid; align-content:start; gap:6px; }
    .schedule-cell strong { color:var(--blue); font-size:12px; }
    .schedule-cell em { font-style:normal; font-size:11px; color:var(--muted); }
    .schedule-chip { display:block; overflow:hidden; white-space:nowrap; text-overflow:ellipsis; border-radius:5px; padding:3px 5px; color:white; background:var(--blue); font-size:11px; }
    .schedule-chip.late { background:var(--red); }
    .schedule-chip.near { background:#b45309; }
    .condition-cell { font-weight:800; text-align:center; }
    .cond-ok { background:#35f235; color:#063b16; }
    .cond-out { background:#ff1616; color:#210000; }
    .cond-warn { background:#fff37a; color:#3f3300; }
    .highlight { background:#fff9b1; }
    @keyframes flowPulse {
      0%,100% { transform:translateX(-2%); filter:saturate(1); }
      50% { transform:translateX(8%); filter:saturate(1.25); }
    }
    @keyframes barBreathe {
      0%,100% { transform:scaleY(.92); opacity:.82; }
      50% { transform:scaleY(1.06); opacity:1; }
    }
    @keyframes scanLine {
      from { transform:translateX(-62%); opacity:.35; }
      50% { opacity:1; }
      to { transform:translateX(62%); opacity:.35; }
    }
    @keyframes machineMove {
      0%,100% { transform:translateX(0); }
      50% { transform:translateX(118px); }
    }
    @keyframes statusPulse {
      0% { box-shadow:0 0 0 0 rgba(34,197,94,.55); }
      70% { box-shadow:0 0 0 9px rgba(34,197,94,0); }
      100% { box-shadow:0 0 0 0 rgba(34,197,94,0); }
    }
    @keyframes growBar {
      from { transform:scaleY(.18); opacity:.55; }
      to { transform:scaleY(1); opacity:1; }
    }
    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { animation-duration:.001ms !important; animation-iteration-count:1 !important; scroll-behavior:auto !important; }
    }
    .subtle-title { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:10px; }
    .subtle-title h3 { margin:0; color:var(--blue); }
    .print-only { display:none; }
    #epp { display:none !important; }
    @media print {
      header, .tabs, #stats, .dashboard-controls, .no-print { display:none !important; }
      main { width:100%; padding:0; }
      body { background:white; }
      .view { display:none !important; }
      .view.active { display:block !important; }
      .panel { box-shadow:none; border:0; padding:0; }
      .table-wrap { max-height:none; overflow:visible; border:0; }
      th { position:static; }
      .print-only { display:block; }
    }
    @media (max-width: 1180px) { .capture-layout { grid-template-columns:1fr; } .latest-captures-panel { position:static; } .latest-captures-wrap { max-height:520px; } }
    @media (max-width: 1180px) { .exec-alert-grid, .profile-grid, .kpi-main-strip { grid-template-columns:repeat(3,minmax(120px,1fr)); } .profile-hero, .profile-section-grid, .kpi-command-grid { grid-template-columns:1fr; } }
    @media (max-width: 900px) { .hero, .grid2 { display:block; } .brand { align-items:flex-start; } .corner-logo { width:96px; height:66px; margin-bottom:10px; } .toolbar, .movement-grid, .req-header-grid, .req-item-grid, .stats { grid-template-columns:1fr; } .exec-alert-grid, .profile-grid, .kpi-main-strip { grid-template-columns:repeat(2,minmax(120px,1fr)); } .capture-form-grid, .tire-track-form { grid-template-columns:repeat(2,minmax(0,1fr)); } .tire-kpi-short { grid-template-columns:repeat(2,minmax(0,1fr)); } header input { min-width:0; margin-top:10px; } .key-card { margin-top:14px; min-width:0; } .tabs { overflow:auto; flex-wrap:nowrap; } .tabs button { flex:0 0 auto; } }
    @media (max-width: 540px) { main { padding:9px; } .panel { padding:12px; } .exec-alert-grid, .profile-grid, .kpi-main-strip, .capture-form-grid, .tire-track-form, .tire-kpi-short { grid-template-columns:1fr; } .capture-section-title, .capture-form-grid .wide, .tire-track-form .wide { grid-column:1; } .capture-actions { display:grid; grid-template-columns:1fr; } .capture-actions .btn { width:100%; } .exec-alert { align-items:flex-start; flex-direction:column; } }
    @media (max-width: 1050px) { .dashboard-grid, .kpi-format-board, .kpi-special-mode #kpiCards, .kpi-diesel-mode #kpiCards, .diesel-card-grid, .diesel-visual-grid { grid-template-columns:1fr; } .diesel-bar-row { grid-template-columns:1fr; } .diesel-bar-row strong, .diesel-bar-row em { text-align:left; } }
  </style>
</head>
<body>
  <header class="hero">
    <div class="brand">
      <img class="corner-logo" src="/static/mga-corner-logo.jfif" alt="MGA">
      <div><h1>Portal MGA mantenimiento</h1><p>KPI, preventivos, bitacora, disponibilidad, diesel y filtros</p></div>
    </div>
    <div class="hero-visual" aria-hidden="true">
      <div class="ops-card"><span>Operacion</span><b>En vivo</b><i></i></div>
      <div class="ops-graph">
        <span>Tendencia semanal</span>
        <div class="ops-lines"><i style="height:22px"></i><i style="height:38px"></i><i style="height:28px"></i><i style="height:46px"></i><i style="height:34px"></i><i style="height:52px"></i><i style="height:42px"></i><i style="height:58px"></i></div>
        <div class="ops-legend"><span><b></b>KPI</span><span><b></b>Alertas</span></div>
      </div>
      <div class="maintenance-video" aria-label="Animacion de operacion de mantenimiento">
        <div class="maintenance-pulse"><i></i>Mtto activo</div>
        <div class="machine-track"></div>
        <div class="machine"></div>
      </div>
    </div>
    <div class="key-card"><label>Clave para editar<input id="apiKey" type="password" placeholder="Pegar clave aqui"></label></div>
  </header>
  <main>
    <nav class="tabs">
      <button class="active" data-tab="dashboard">Dashboard KPI</button>
      <button data-tab="fichaEquipo">Ficha equipo</button>
      <button data-tab="mensual">Reporte mensual/semanal</button>
      <button data-tab="preventivos">PR Preventivos</button>
      <button data-tab="backlog">Backlog</button>
      <button data-tab="ordenesTrabajo">Ordenes trabajo</button>
      <button data-tab="servicios">Servicios realizados</button>
      <button data-tab="ejecucionPreventivos">Ejecucion preventivos</button>
      <button data-tab="bitacora">Bitacora</button>
      <button data-tab="captura">Captura diaria</button>
      <button data-tab="disponibilidad">Disponibilidad</button>
      <button data-tab="requisiciones">Requisiciones</button>
      <button data-tab="seguimientoReq">Seguimiento req.</button>
      <button data-tab="mangueras">Mangueras</button>
      <button data-tab="diesel">Diesel</button>
      <button data-tab="llantasTrack">Seguimiento llantas</button>
      <button data-tab="refacciones">Refacciones equipo</button>
      <button data-tab="equipos">Filtros por equipo</button>
      <button data-tab="inventario">Concentrado / movimientos</button>
      <button data-tab="auditoria">Auditoria</button>
      <button data-tab="importar">Importar / exportar</button>
    </nav>
    <section class="stats" id="stats"></section>
    <section id="dashboard" class="view active">
      <div class="panel toolbar dashboard-controls">
        <label>Grupo<select id="kpiGroup"></select></label>
        <label>Desde<input id="kpiStart" type="date"></label>
        <label>Hasta<input id="kpiEnd" type="date"></label>
        <button class="btn" id="renderKpiBtn">Actualizar KPI</button>
        <button class="btn secondary" id="printKpiBtn">Imprimir PDF</button>
        <button class="btn secondary" id="kpiImageBtn">Descargar imagen</button>
        <button class="btn secondary" id="kpiExcelBtn">Excel editable</button>
        <button class="btn secondary" id="meetingModeBtn">Modo reunion</button>
      </div>
      <div class="panel toolbar kpi-sim-toolbar">
        <label class="inline-check"><input id="kpiSimEnabled" type="checkbox"> Modo simulacion</label>
        <label>Escenario<input id="kpiSimName" value="Escenario 1"></label>
        <label>Meta disp %<input id="kpiSimMetaAvailability" type="number" step="0.1" value="85"></label>
        <label>Meta util %<input id="kpiSimMetaUtilization" type="number" step="0.1" value="75"></label>
        <label>Meta conf %<input id="kpiSimMetaReliability" type="number" step="0.1" value="80"></label>
        <label>Meta TMEF h<input id="kpiSimMetaTmef" type="number" step="0.1" value="8"></label>
        <label>Meta TMPR h<input id="kpiSimMetaTmpr" type="number" step="0.1" value="4"></label>
        <label>Hrs periodo %<input id="kpiSimPeriod" type="number" step="1" value="100"></label>
        <label>Hrs trab %<input id="kpiSimWorked" type="number" step="1" value="100"></label>
        <label>Hrs MP %<input id="kpiSimMp" type="number" step="1" value="100"></label>
        <label>Hrs MC %<input id="kpiSimMc" type="number" step="1" value="100"></label>
        <label>Paradas %<input id="kpiSimStops" type="number" step="1" value="100"></label>
        <label>Hrs mision<input id="kpiSimMission" type="number" step="0.1" value="24"></label>
        <span class="muted kpi-sim-note">Solo cambia la vista y las descargas simuladas. No guarda datos reales.</span>
      </div>
      <div class="panel no-print">
        <div class="subtle-title"><h3>Indicadores principales</h3><span class="muted" id="kpiMainSummaryNote"></span></div>
        <div class="kpi-main-strip" id="kpiMainStrip"></div>
      </div>
      <div class="panel executive-board no-print">
        <div class="subtle-title"><h3>Prioridad operativa</h3><span class="muted" id="execUpdated">Alertas automaticas</span></div>
        <div class="exec-alert-grid" id="execCards"></div>
        <div class="exec-alert-list" id="execAlerts"></div>
      </div>
      <div class="panel dashboard-command no-print">
        <div class="subtle-title"><h3>Semaforo y Top 10 prioridades</h3><span class="muted" id="kpiCommandUpdated"></span></div>
        <div class="kpi-command-grid">
          <div>
            <div class="subtle-title"><h3>Top 10 equipos a atender</h3><span class="muted">Riesgo calculado</span></div>
            <div class="priority-list" id="kpiPriorityList"></div>
          </div>
          <div class="table-wrap">
            <table id="kpiSemaphoreTable"></table>
          </div>
        </div>
      </div>
      <div class="panel" id="kpiPrintArea">
        <div class="subtle-title"><h3 id="kpiTitle">Dashboard KPI</h3><span class="muted" id="portalUpdated"></span></div>
        <div class="kpi-format-board">
          <div class="kpi-side" id="kpiCards"></div>
          <div class="chart" id="kpiChart"></div>
          <div class="kpi-side" id="kpiSideCards"></div>
        </div>
        <div class="table-wrap kpi-report-table"><table id="kpiTable"></table></div>
      </div>
    </section>
    <section id="fichaEquipo" class="view">
      <div class="panel toolbar">
        <label>Equipo<select id="fichaEquipment"></select></label>
        <label>Desde<input id="fichaStart" type="date"></label>
        <label>Hasta<input id="fichaEnd" type="date"></label>
        <button class="btn" id="renderFichaBtn">Actualizar ficha</button>
        <button class="btn secondary" id="fichaGoCaptureBtn">Capturar diario</button>
        <button class="btn secondary" id="fichaGoServiceBtn">Nuevo preventivo</button>
      </div>
      <div class="profile-hero">
        <div class="profile-card" id="fichaHeader"></div>
        <div class="profile-grid" id="fichaMetrics"></div>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Alertas del equipo</h3><span class="muted" id="fichaAlertCount"></span></div>
        <div class="exec-alert-list" id="fichaAlerts"></div>
      </div>
      <div class="profile-section-grid">
        <div class="panel">
          <div class="subtle-title"><h3>Ultimos servicios</h3><span class="muted" id="fichaServiceCount"></span></div>
          <div class="table-wrap"><table id="fichaServicesTable"></table></div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Preventivos programados</h3><span class="muted" id="fichaPreventiveCount"></span></div>
          <div class="table-wrap"><table id="fichaPreventivesTable"></table></div>
        </div>
      </div>
      <div class="profile-section-grid">
        <div class="panel">
          <div class="subtle-title"><h3>Capturas recientes</h3><span class="muted" id="fichaCaptureCount"></span></div>
          <div class="table-wrap"><table id="fichaCapturesTable"></table></div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Refacciones y llantas</h3><span class="muted" id="fichaPartsCount"></span></div>
          <div class="table-wrap"><table id="fichaPartsTable"></table></div>
        </div>
      </div>
    </section>
    <section id="mensual" class="view">
      <div class="panel toolbar">
        <label>Mes<select id="monthlyMonth">
          <option value="1">Enero</option><option value="2">Febrero</option><option value="3">Marzo</option><option value="4">Abril</option>
          <option value="5">Mayo</option><option value="6">Junio</option><option value="7">Julio</option><option value="8">Agosto</option>
          <option value="9">Septiembre</option><option value="10">Octubre</option><option value="11">Noviembre</option><option value="12">Diciembre</option>
        </select></label>
        <label>Ano<input id="monthlyYear" type="number" min="2000" max="2100"></label>
        <button class="btn" id="monthlyPptBtn">Descargar PowerPoint</button>
      </div>
      <div class="panel toolbar">
        <label>Fecha base semana<input id="weeklyBase" type="date"></label>
        <label>Desde<input id="weeklyStart" type="date"></label>
        <label>Hasta<input id="weeklyEnd" type="date"></label>
        <button class="btn secondary" id="weeklyApplyBtn">Aplicar semana</button>
        <button class="btn" id="weeklyPptBtn">Descargar semanal</button>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Reporte mensual y semanal PowerPoint</h3><span class="muted" id="monthlyStatus"></span></div>
        <div class="stats">
          <div class="stat"><strong>Barrenacion</strong>KPI mensual</div>
          <div class="stat"><strong>Rezagado</strong>KPI mensual</div>
          <div class="stat"><strong>Acarreo</strong>KPI mensual</div>
          <div class="stat"><strong>Utilitario</strong>KPI mensual</div>
          <div class="stat"><strong>Llantas</strong>Vida util</div>
          <div class="stat"><strong>Diesel</strong>Consumo</div>
          <div class="stat"><strong>Aceites</strong>KPI mensual</div>
        </div>
        <p class="muted" id="weeklyStatus"></p>
      </div>
    </section>
    <section id="preventivos" class="view">
      <div class="panel toolbar">
        <label>Periodo<select id="prPeriod"><option>Mes</option><option>Semana</option><option>Año</option></select></label>
        <label>Fecha base<input id="prBase" type="date"></label>
        <label>Equipo<select id="prEquipment"></select></label>
        <label>Buscar<input id="prSearch" placeholder="Equipo, componente, estado"></label>
        <button class="btn" id="renderPrBtn">Consultar</button>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3 id="prTitle">Preventivos programados</h3><span class="muted" id="prCount"></span></div>
        <div class="schedule-strip" id="prCalendar"></div>
      </div>
      <div class="table-wrap"><table id="prTable"></table></div>
    </section>
    <section id="backlog" class="view">
      <div class="panel toolbar">
        <label>Desde<input id="backlogStart" type="date"></label>
        <label>Hasta<input id="backlogEnd" type="date"></label>
        <label>Nivel<select id="backlogLevel"><option value="">Todos</option><option>ALTA</option><option>MEDIA</option><option>BAJA</option></select></label>
        <label>Origen<select id="backlogSource"><option value="">Todos</option><option>Preventivo</option><option>Captura</option><option>OT</option><option>Requisicion</option></select></label>
        <label>Estado<select id="backlogStatus"><option value="">Todos</option><option>Pendiente</option><option>En proceso</option><option>Atendido</option><option>Cancelado</option></select></label>
        <label>Buscar<input id="backlogSearch" placeholder="Equipo, sistema, detalle"></label>
        <button class="btn" id="renderBacklogBtn">Actualizar</button>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3 id="backlogTitle">Backlog priorizado</h3><span class="muted" id="backlogCount"></span></div>
        <div class="stats" id="backlogStats"></div>
      </div>
      <div class="grid2">
        <div class="table-wrap"><table id="backlogTable"></table></div>
        <div class="table-wrap"><table id="backlogSystemTable"></table></div>
      </div>
    </section>
    <section id="ordenesTrabajo" class="view">
      <div class="grid2">
        <div class="panel">
          <div class="subtle-title"><h3>Orden de trabajo</h3><span class="muted" id="woStatus"></span></div>
          <div class="capture-form-grid">
            <input id="woId" type="hidden">
            <label>Folio<input id="woFolio" placeholder="Automatico" readonly></label>
            <label>Fecha<input id="woDate" type="date"></label>
            <label>Equipo<select id="woEquipment"></select></label>
            <label>Origen<select id="woOrigin"><option>MANUAL</option><option>CAPTURA</option><option>PREVENTIVO</option><option>LLANTA</option><option>STOCK</option><option>FICHA EQUIPO</option></select></label>
            <label>Prioridad<select id="woPriority"><option>MEDIA</option><option>ALTA</option><option>URGENTE</option><option>BAJA</option></select></label>
            <label>Estatus<select id="woState"><option>ABIERTA</option><option>EN PROCESO</option><option>CERRADA</option><option>CANCELADA</option></select></label>
            <label>Responsable<input id="woResponsible" placeholder="Responsable"></label>
            <label>Mecanico<input id="woMechanic" placeholder="Mecanico"></label>
            <label>Supervisor<input id="woSupervisor" placeholder="Supervisor"></label>
            <label class="wide">Descripcion del trabajo<textarea id="woDescription" rows="3" placeholder="Falla, condicion o trabajo requerido"></textarea></label>
            <label class="wide">Accion / cierre<textarea id="woAction" rows="2" placeholder="Trabajo realizado o accion pendiente"></textarea></label>
            <label class="wide">Refacciones usadas<textarea id="woParts" rows="2" placeholder="Refacciones usadas"></textarea></label>
            <label class="wide">Lubricantes<textarea id="woLubricants" rows="2" placeholder="Lubricantes usados"></textarea></label>
            <label class="wide">Evidencia / firma<textarea id="woEvidence" rows="2" placeholder="Foto, firma, folio o evidencia"></textarea></label>
          </div>
          <div class="req-actions capture-actions">
            <button class="btn secondary" id="woNewBtn">Nueva OT</button>
            <button class="btn" id="woSaveBtn">Guardar OT</button>
            <button class="btn secondary" id="woCloseBtn">Cerrar OT</button>
            <button class="btn danger" id="woDeleteBtn">Eliminar OT</button>
          </div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Resumen OT</h3><span class="muted" id="woSummaryText"></span></div>
          <div class="exec-alert-grid" id="woSummaryCards"></div>
        </div>
      </div>
      <div class="panel toolbar">
        <label>Equipo<select id="woFilterEquipment"></select></label>
        <label>Estatus<select id="woFilterStatus"><option value="">Todos</option><option>ABIERTA</option><option>EN PROCESO</option><option>CERRADA</option><option>CANCELADA</option></select></label>
        <label>Prioridad<select id="woFilterPriority"><option value="">Todas</option><option>URGENTE</option><option>ALTA</option><option>MEDIA</option><option>BAJA</option></select></label>
        <label>Buscar<input id="woSearch" placeholder="Folio, equipo, descripcion"></label>
        <button class="btn" id="woRefreshBtn">Actualizar</button>
      </div>
      <div class="table-wrap"><table id="woTable"></table></div>
    </section>
    <section id="servicios" class="view">
      <div class="panel toolbar">
        <label>Equipo<select id="srvEquipment"></select></label>
        <label>Servicio<select id="srvInterval"><option value="">Todos</option><option>PM1</option><option>PM2</option><option>PM3</option><option>PM4</option><option>250H</option><option>500H</option><option>750H</option><option>1000H</option></select></label>
        <label>Tipo<select id="srvType"><option value="">Todos</option><option>Programado</option><option>No programado</option></select></label>
        <label>Desde<input id="srvStart" type="date"></label>
        <label>Hasta<input id="srvEnd" type="date"></label>
        <label>Buscar<input id="srvSearch" placeholder="Componente, OT, notas"></label>
        <button class="btn" id="renderSrvBtn">Consultar</button>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3 id="srvTitle">Servicios realizados</h3><span class="muted" id="srvCount"></span></div>
        <div class="stats" id="srvStats"></div>
      </div>
      <div class="table-wrap"><table id="srvTable"></table></div>
    </section>
    <section id="ejecucionPreventivos" class="view">
      <div class="grid2">
        <div class="panel">
          <div class="subtle-title"><h3>Actualizar servicios preventivos</h3><span class="muted" id="prevExecStatus"></span></div>
          <p class="muted">Los servicios abiertos quedan pendientes. Solo cuando el estatus sea Cerrado se reflejan en Servicios realizados y se usan para el siguiente preventivo.</p>
          <div class="capture-form-grid">
            <div class="capture-section-title">Datos del servicio</div>
            <input id="prevExecId" type="hidden">
            <label>Fecha servicio<input id="prevExecDate" type="date"></label>
            <label>Equipo<select id="prevExecEquipment"></select></label>
            <label>Supervisor<input id="prevExecSupervisor" placeholder="Supervisor"></label>
            <label>Mecanico<input id="prevExecMechanic" placeholder="Mecanico"></label>
            <label>Tipo servicio<select id="prevExecServiceType"><option value="PM1">PM1 - 250H</option><option value="PM2">PM2 - 500H</option><option value="PM3">PM3 - 750H</option><option value="PM4">PM4 - 1000H</option></select></label>
            <label>Tipo de atributo<select id="prevExecAttribute"><option>GENERAL</option><option>MOTOR</option><option>ELECT</option><option>DIESEL</option><option>HIDRAULICO</option><option>TRANSMISION</option><option>LLANTAS</option><option>FRENOS</option><option>OTRO</option></select></label>
            <label>Horometro cierre<input id="prevExecMeter" type="number" step="0.1" min="0" value="0"></label>
            <label>Estatus<select id="prevExecState"><option>ABIERTO</option><option>EN PROCESO</option><option>CERRADO</option><option>CANCELADO</option></select></label>
            <label class="wide">Refacciones usadas<textarea id="prevExecParts" rows="3" placeholder="Ej. filtro aceite 1 pza; banda alternador 1 pza"></textarea></label>
            <div class="capture-section-title">Aceites y fluidos (litros)</div>
            <label class="capture-fluid">Aceite total<input id="prevExecOilTotal" type="number" step="0.1" min="0" value="0" readonly></label>
            <label class="capture-fluid">15W40<input id="prevExecOil15w40" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">HCO ISO 68<input id="prevExecOilHco68" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">SAE 30<input id="prevExecOilSae30" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">85W140<input id="prevExecOil85w140" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">ALMO<input id="prevExecAlmo" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">Refrigerante<input id="prevExecCoolant" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">Hidraulico VG100<input id="prevExecOilVg100" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">ATF<input id="prevExecAtf" type="number" step="0.1" min="0" value="0"></label>
            <div class="capture-section-title">Checklist de cierre</div>
            <label class="inline-check"><input id="prevChkInspection" type="checkbox"> Inspeccion realizada</label>
            <label class="inline-check"><input id="prevChkFilters" type="checkbox"> Filtros/refacciones aplicadas</label>
            <label class="inline-check"><input id="prevChkLubrication" type="checkbox"> Lubricacion registrada</label>
            <label class="inline-check"><input id="prevChkElectrical" type="checkbox"> Revision electrica</label>
            <label class="inline-check"><input id="prevChkTest" type="checkbox"> Prueba final</label>
            <label class="inline-check"><input id="prevChkSupervisor" type="checkbox"> Validado por supervisor</label>
            <label class="wide">Evidencia / firma<textarea id="prevExecEvidence" rows="2" placeholder="Folio, foto, firma o evidencia del cierre"></textarea></label>
            <label class="wide">Observaciones<textarea id="prevExecNotes" rows="3" placeholder="Trabajo realizado, pendientes o condicion encontrada"></textarea></label>
          </div>
          <div class="req-actions capture-actions">
            <button class="btn secondary" id="prevExecNewBtn">Nuevo</button>
            <button class="btn" id="prevExecSaveBtn">Guardar</button>
            <button class="btn secondary" id="prevExecCloseBtn">Cerrar servicio</button>
            <button class="btn danger" id="prevExecDeleteBtn">Eliminar</button>
            <button class="btn secondary" id="prevExecViewHistoryBtn">Ver servicios realizados</button>
          </div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Servicios abiertos</h3><span class="muted" id="prevExecOpenCount"></span></div>
          <div class="table-wrap"><table id="prevExecOpenTable"></table></div>
        </div>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Historial preventivo web</h3><span class="muted" id="prevExecClosedCount"></span></div>
      </div>
      <div class="table-wrap"><table id="prevExecClosedTable"></table></div>
    </section>
    <section id="bitacora" class="view">
      <div class="panel toolbar">
        <label>Equipo<select id="bitEquipment"></select></label>
        <label>Desde<input id="bitStart" type="date"></label>
        <label>Hasta<input id="bitEnd" type="date"></label>
        <label>Buscar<input id="bitSearch" placeholder="Componente, falla, observacion"></label>
        <button class="btn" id="renderBitBtn">Actualizar</button>
      </div>
      <div class="table-wrap"><table id="bitTable"></table></div>
    </section>
    <section id="captura" class="view">
      <div class="grid2 capture-layout">
        <div class="panel">
          <div class="subtle-title"><h3>Captura diaria</h3><span class="muted" id="capStatus"></span></div>
          <div class="capture-form-grid">
            <div class="capture-section-title">Datos de operacion</div>
            <label>Fecha<input id="capDate" type="date"></label>
            <label>Turno<select id="capShift"><option>Turno 1</option><option>Turno 2</option><option>General</option></select></label>
            <label>Equipo<select id="capEquipment"></select></label>
            <label>Componente<select id="capComponent"></select></label>
            <label>Horometro inicial<input id="capHi" type="number" step="0.1" min="0" value="0"><span class="muted" id="capHiHint"></span></label>
            <label>Horometro final<input id="capHf" type="number" step="0.1" min="0" value="0"></label>
            <label>Hrs trabajadas<input id="capWorked" type="number" step="0.1" min="0" value="0"></label>
            <label>Hrs MP<input id="capMp" type="number" step="0.1" min="0" value="0"></label>
            <label>Hrs MC<input id="capMc" type="number" step="0.1" min="0" value="0"></label>
            <label>Stand By<input id="capStandby" type="number" step="0.1" min="0" value="0"></label>
            <label># Paradas<input id="capStops" type="number" step="1" min="0" value="0"></label>
            <label>Estatus<select id="capCaptureStatus"><option>Disponible</option><option>No Disponible</option><option>Stand By</option><option>Operativa</option></select></label>
            <div class="capture-section-title">Aceites y fluidos (litros)</div>
            <label class="capture-fluid">Aceite total<input id="capOil" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">15W40<input id="capOil15w40" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">HCO ISO 68<input id="capOilHco68" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">SAE 30<input id="capOilSae30" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">85W140<input id="capOil85w140" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">ALMO<input id="capAlmo" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">Refrigerante<input id="capCoolant" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">Hidraulico VG100<input id="capOilVg100" type="number" step="0.1" min="0" value="0"></label>
            <label class="capture-fluid">ATF<input id="capAtf" type="number" step="0.1" min="0" value="0"></label>
            <div class="capture-section-title">Falla y observaciones</div>
            <label class="wide">Falla<input id="capFault" placeholder="Falla detectada"></label>
            <label class="wide">Desgaste<input id="capWear" placeholder="Desgaste observado"></label>
            <label class="wide">Observaciones<textarea id="capObservations" rows="3" placeholder="Detalle de la captura"></textarea></label>
          </div>
          <div class="req-actions capture-actions">
            <button class="btn secondary" id="capNewBtn">Nueva captura</button>
            <button class="btn" id="capSaveBtn">Guardar captura</button>
            <button class="btn secondary" id="capSaveNewBtn">Guardar y nuevo</button>
            <button class="btn secondary" id="capRefreshBtn">Actualizar bitacora</button>
          </div>
          <div class="quick-alerts" id="capQuickAlerts"></div>
        </div>
        <div class="panel latest-captures-panel">
          <div class="subtle-title"><h3>Ultimas capturas</h3><span class="muted" id="capRecentCount"></span></div>
          <div class="table-wrap latest-captures-wrap"><table id="capRecentTable"></table></div>
        </div>
      </div>
    </section>
    <section id="disponibilidad" class="view">
      <div class="panel toolbar">
        <label>Categoria / equipo<input id="dispSearch" placeholder="Buscar"></label>
        <label>Condicion<select id="dispStatus"><option value="">Todas</option><option>DISPONIBLE</option><option>FUERA DE SERVICIO</option><option>OPERATIVA</option></select></label>
        <button class="btn" id="renderDispBtn">Actualizar</button>
      </div>
      <div class="table-wrap"><table id="dispTable"></table></div>
    </section>
    <section id="requisiciones" class="view">
      <div class="grid2">
        <div class="panel">
          <div class="subtle-title"><h3>Generar requisicion</h3><span class="muted" id="reqStatus"></span></div>
          <div class="req-header-grid">
            <label>Folio<input id="reqFolio"></label>
            <label>Fecha solicitada<input id="reqDate" type="date"></label>
            <label>Fecha autorizacion<input id="reqAuthDate" type="date"></label>
            <label>Equipo<select id="reqEquipment"><option>PARA STOCK</option></select></label>
            <label>Centro costos<input id="reqCostCenter"></label>
            <label>Area solicita<input id="reqArea" value="MTTO"></label>
            <label>Ubicacion<input id="reqLocation" value="PROVIDENCIA"></label>
            <label>Unidad solicita<input id="reqRequestingUnit" value="TALLER CENTRAL"></label>
            <label>Unidad operativa<input id="reqOperatingUnit" value="PROVIDENCIA"></label>
            <label>Prioridad<select id="reqPriority"><option>URGENTE</option><option>ORDINARIA</option></select></label>
            <label>Recomendacion<select id="reqRecommendation"><option>ORIGINAL</option><option>FABRICACION LOCAL</option></select></label>
            <label>Estatus<select id="reqReqStatus"><option>Abierta</option><option>Autorizada</option><option>Surtida</option><option>Cancelada</option></select></label>
            <label class="wide">Notas<textarea id="reqNotes" rows="2"></textarea></label>
          </div>
          <div class="req-actions">
            <button class="btn secondary" id="reqNewBtn">Nueva</button>
            <button class="btn" id="reqSaveBtn">Guardar requisicion</button>
            <button class="btn secondary" id="reqPdfBtn">Descargar PDF</button>
            <button class="btn secondary" id="reqPrintBtn">Imprimir</button>
            <button class="btn danger" id="reqDeleteBtn">Eliminar</button>
          </div>
          <h3>Partida</h3>
          <div class="req-item-grid">
            <label>Cantidad<input id="reqItemQty" type="number" step="0.01" value="1"></label>
            <label>Unidad<select id="reqItemUnit"></select></label>
            <label>No. parte<input id="reqItemPart"></label>
            <label>Descripcion<input id="reqItemDesc"></label>
          </div>
          <div class="req-actions">
            <button class="btn" id="reqAddItemBtn">Agregar / actualizar partida</button>
            <button class="btn secondary" id="reqClearItemBtn">Nueva partida</button>
            <button class="btn danger" id="reqDeleteItemBtn">Eliminar partida</button>
          </div>
          <div class="table-wrap" style="max-height:300px; margin-top:10px;"><table id="reqItemsTable"></table></div>
        </div>
        <div class="panel">
          <h3>Consulta Clave / Producto</h3>
          <label>Buscar por clave o producto<input id="reqProductSearch" placeholder="Ej. 384-8612, bomba, filtro"></label>
          <div class="table-wrap" style="max-height:300px; margin-top:10px;"><table id="reqProductsTable"></table></div>
          <h3>Requisiciones guardadas</h3>
          <div class="table-wrap" style="max-height:360px;"><table id="reqListTable"></table></div>
        </div>
      </div>
    </section>
    <section id="seguimientoReq" class="view">
      <div class="panel toolbar">
        <label>Buscar<input id="trackSearch" placeholder="Folio, OC, proveedor, estatus"></label>
        <label>Estatus<select id="trackStatus"><option value="">Todos</option><option value="sin_oc">Sin OC</option><option value="con_oc">Con OC</option><option value="recibido">Recibido</option></select></label>
        <input id="trackImportFile" type="file" accept=".xlsx,.xlsm">
        <label class="inline-check"><input id="trackCreateMissing" type="checkbox"> Crear folios faltantes</label>
        <button class="btn" id="trackImportBtn">Importar seguimiento</button>
        <button class="btn secondary" id="trackExportBtn">Descargar Excel</button>
        <button class="btn secondary" id="trackRefreshBtn">Actualizar</button>
      </div>
      <div class="grid2">
        <div class="panel">
          <div class="subtle-title"><h3>Seguimiento de compras</h3><span class="muted" id="trackSummary"></span></div>
          <div class="table-wrap" style="max-height:620px;"><table id="trackTable"></table></div>
          <pre id="trackImportResult"></pre>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Actualizar requisicion</h3><span class="muted" id="trackSelected"></span></div>
          <div class="req-header-grid">
            <label>Folio<input id="trackFolio" readonly></label>
            <label>Fecha requisicion<input id="trackReqDate" readonly></label>
            <label>Equipo<input id="trackEquipment" readonly></label>
            <label>Estatus compras<input id="trackPurchaseStatus" list="trackStatusOptions"></label>
            <datalist id="trackStatusOptions"><option>EN REVISION</option><option>COTIZANDO</option><option>POR AUTORIZAR</option><option>CON ORDEN DE COMPRA</option><option>PARCIAL</option><option>RECIBIDO</option><option>CANCELADO</option></datalist>
            <label>Orden de compra<input id="trackPurchaseOrder" placeholder="OC / PO"></label>
            <label>Fecha OC<input id="trackPurchaseOrderDate" type="date"></label>
            <label>Proveedor<input id="trackSupplier"></label>
            <label>Comprador<input id="trackBuyer"></label>
            <label>T.E / Promesa<input id="trackExpectedDate" type="text" placeholder="3 SEMANAS, INMEDIATA o fecha"></label>
            <label>Fecha recibido<input id="trackReceivedDate" type="date"></label>
            <label class="wide">Notas compras<textarea id="trackNotes" rows="4"></textarea></label>
          </div>
          <div class="req-actions">
            <button class="btn" id="trackSaveBtn">Guardar seguimiento</button>
            <button class="btn secondary" id="trackOpenReqBtn">Abrir en requisiciones</button>
          </div>
          <div class="table-wrap" style="max-height:260px; margin-top:10px;"><table id="trackItemsTable"></table></div>
        </div>
      </div>
    </section>
    <section id="mangueras" class="view">
      <div class="panel toolbar">
        <label>Periodo<select id="hosePeriod"><option>Mes</option><option>Semana</option><option>Rango</option></select></label>
        <label>Fecha base<input id="hoseBase" type="date"></label>
        <label>Desde<input id="hoseStart" type="date"></label>
        <label>Hasta<input id="hoseEnd" type="date"></label>
        <label>Equipo<select id="hoseFilterEquipment"></select></label>
        <button class="btn secondary" id="hoseApplyPeriodBtn">Aplicar periodo</button>
        <button class="btn" id="hoseRefreshBtn">Actualizar</button>
      </div>
      <div class="stats" id="hoseStats"></div>
      <div class="grid2">
        <div class="panel">
          <div class="subtle-title"><h3>Captura de mangueras y conexiones</h3><span class="muted" id="hoseEditStatus"></span></div>
          <div class="movement-grid">
            <label>Fecha<input id="hoseDate" type="date"></label>
            <label>Equipo<select id="hoseEquipment"></select></label>
            <label>Sistema<select id="hoseSystem"><option>HIDRAULICO</option><option>AIRE</option><option>AGUA</option><option>DIESEL</option><option>LUBRICACION</option><option>FRENOS</option><option>OTRO</option></select></label>
            <label>Tipo<select id="hoseType"><option>MANGUERA</option><option>CONEXION</option><option>ADAPTADOR</option><option>ABRAZADERA</option><option>OTRO</option></select></label>
            <label>Diametro<input id="hoseDiameter" placeholder="1/2, 3/4, #8"></label>
            <label>Longitud m<input id="hoseLength" type="number" step="0.01" value="0"></label>
            <label>Cantidad<input id="hoseQty" type="number" step="0.01" value="1"></label>
            <label>Vida est. dias<input id="hoseLifeDays" type="number" step="1" value="30"></label>
            <label>Consumo aprox/sem<input id="hoseEstimatedWeekly" type="number" step="0.01" value="0"></label>
            <label>Causa<input id="hoseReason"></label>
            <label>Tecnico<input id="hoseTech"></label>
            <label class="wide">Notas<textarea id="hoseNotes" rows="2"></textarea></label>
          </div>
          <div class="req-actions">
            <button class="btn secondary" id="hoseNewBtn">Nuevo</button>
            <button class="btn" id="hoseSaveBtn">Guardar cambio</button>
            <button class="btn danger" id="hoseDeleteBtn">Eliminar cambio</button>
          </div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Resumen de consumo</h3><span class="muted" id="hosePeriodLabel"></span></div>
          <div class="table-wrap" style="max-height:520px;"><table id="hoseSummaryTable"></table></div>
        </div>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Historial de cambios</h3><span class="muted">Selecciona una fila para editar</span></div>
        <div class="table-wrap" style="max-height:460px;"><table id="hoseRecordsTable"></table></div>
      </div>
    </section>
    <section id="diesel" class="view">
      <div class="panel toolbar">
        <label>Periodo<select id="dieselPeriod"><option>Mes</option><option>Semana</option><option>Rango</option></select></label>
        <label>Fecha base<input id="dieselBase" type="date"></label>
        <label>Desde<input id="dieselStart" type="date"></label>
        <label>Hasta<input id="dieselEnd" type="date"></label>
        <label>Equipo<select id="dieselFilterEquipment"></select></label>
        <label>Meta L/H<input id="dieselMeta" type="number" step="0.1" value="25"></label>
        <button class="btn secondary" id="dieselApplyPeriodBtn">Aplicar periodo</button>
        <button class="btn" id="dieselRefreshBtn">Actualizar diesel</button>
        <button class="btn secondary" id="dieselPrintBtn">Imprimir reporte</button>
        <button class="btn secondary" id="dieselExcelBtn">Descargar Excel</button>
      </div>
      <div class="stats" id="dieselStats"></div>
      <div class="grid2">
        <div class="panel">
          <div class="subtle-title"><h3>Captura diesel por equipo / turno</h3><span class="muted" id="dieselEditStatus"></span></div>
          <div class="movement-grid">
            <label>Fecha<input id="dieselDate" type="date"></label>
            <label>Equipo<select id="dieselEquipment"></select></label>
            <label>Condicion<select id="dieselCondition"><option>DISPONIBLE</option><option>FUERA DE SERVICIO</option><option>OPERATIVA</option></select></label>
            <label>Turno<select id="dieselShift"><option>1</option><option>2</option><option>GENERAL</option></select></label>
            <label>Horometro inicial<input id="dieselHi" type="number" step="0.1" value="0"></label>
            <label>Horometro final<input id="dieselHf" type="number" step="0.1" value="0"></label>
            <label>Horas trabajadas<input id="dieselHours" type="number" step="0.1" value="0"></label>
            <label>Consumo diesel L<input id="dieselLiters" type="number" step="0.1" value="0"></label>
            <label>Operador<input id="dieselOperator"></label>
            <label>Despachador<input id="dieselDispatcher"></label>
            <label>Supervisor<input id="dieselSupervisor"></label>
            <label class="wide">Notas<textarea id="dieselNotes" rows="2"></textarea></label>
          </div>
          <div class="req-actions">
            <button class="btn secondary" id="dieselNewBtn">Nueva captura</button>
            <button class="btn" id="dieselSaveBtn">Guardar captura</button>
            <button class="btn danger" id="dieselDeleteBtn">Eliminar captura</button>
          </div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Control diario de tanque</h3><span class="muted">Llegada y existencias por dia</span></div>
          <div class="movement-grid">
            <label>Fecha<input id="dieselDayDate" type="date"></label>
            <label>Llegada diesel L<input id="dieselReceived" type="number" step="0.1" value="0"></label>
            <label>Existencia inicial L<input id="dieselInitial" type="number" step="0.1" value="0"></label>
            <label>Existencia final L<input id="dieselFinal" type="number" step="0.1" value="0"></label>
            <label>Exist. PROSERMIN<input id="dieselProserminStock" type="number" step="0.1" value="0"></label>
            <label class="wide">Proveedor / zona<input id="dieselSupplier"></label>
            <label class="wide">Notas<textarea id="dieselDayNotes" rows="2"></textarea></label>
            <button class="btn wide" id="dieselDaySaveBtn">Guardar dia</button>
          </div>
        </div>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3 id="dieselTitle">Rendimiento diesel</h3><span class="muted" id="dieselUpdated"></span></div>
        <div class="table-wrap"><table id="dieselReportTable"></table></div>
      </div>
      <div class="grid2">
        <div class="panel">
          <h3>Consumo diario</h3>
          <div class="table-wrap" style="max-height:360px;"><table id="dieselDailyTable"></table></div>
        </div>
        <div class="panel">
          <h3>Capturas diesel</h3>
          <div class="table-wrap" style="max-height:360px;"><table id="dieselRecordsTable"></table></div>
        </div>
      </div>
    </section>
    <section id="llantasTrack" class="view">
      <div class="grid2 tire-track-layout">
        <div class="panel">
          <div class="subtle-title"><h3>Seguimiento de llantas</h3><span class="muted" id="tireTrackStatus"></span></div>
          <div class="tire-track-form">
            <label>Fecha<input id="tireTrackDate" type="date"></label>
            <label>Tipo<select id="tireTrackType"><option>INSPECCION</option><option>MOVIMIENTO</option><option>MODIFICACION</option><option>MONTAJE</option><option>ROTACION</option><option>REPARACION</option><option>DESMONTAJE</option><option>BAJA</option></select></label>
            <label>Serie llanta<input id="tireTrackCode" list="tireTrackCodes" placeholder="Serie / codigo"><datalist id="tireTrackCodes"></datalist></label>
            <label>Equipo<select id="tireTrackEquipment"></select></label>
            <label>Posicion<input id="tireTrackPosition" placeholder="Ej. DEL IZQ"></label>
            <label>Estatus<select id="tireTrackMountStatus"><option>MONTADA</option><option>ALMACEN</option><option>REPARACION</option><option>BAJA</option></select></label>
            <label>Marca<input id="tireTrackBrand"></label>
            <label>Modelo<input id="tireTrackModel"></label>
            <label>Medida<input id="tireTrackSize"></label>
            <label>Hor. montaje<input id="tireTrackInstallMeter" type="number" step="0.1" min="0" value="0"></label>
            <label>Hor. actual<input id="tireTrackCurrentMeter" type="number" step="0.1" min="0" value="0"></label>
            <label>Vida objetivo h<input id="tireTrackTargetHours" type="number" step="0.1" min="0" value="0"></label>
            <label>Piso inicial mm<input id="tireTrackTreadInitial" type="number" step="0.1" min="0" value="0"></label>
            <label>Piso actual mm<input id="tireTrackTreadCurrent" type="number" step="0.1" min="0" value="0"></label>
            <label>Presion PSI<input id="tireTrackPressure" type="number" step="0.1" min="0" value="0"></label>
            <label>Tecnico<input id="tireTrackTechnician"></label>
            <label class="wide">Notas<textarea id="tireTrackNotes" rows="3" placeholder="Inspeccion, movimiento o modificacion realizada"></textarea></label>
          </div>
          <div class="req-actions">
            <button class="btn secondary" id="tireTrackNewBtn">Nueva</button>
            <button class="btn" id="tireTrackSaveBtn">Guardar seguimiento</button>
            <button class="btn secondary" id="tireTrackKpiBtn">Ver KPI Llantas</button>
            <button class="btn secondary" id="tireTrackRefreshBtn">Actualizar</button>
          </div>
        </div>
        <div class="panel">
          <div class="subtle-title"><h3>Resumen KPI llantas</h3><span class="muted" id="tireTrackCount"></span></div>
          <div class="tire-kpi-short" id="tireTrackSummary"></div>
          <div class="table-wrap" style="max-height:420px;"><table id="tireTrackTable"></table></div>
        </div>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Historial de inspecciones y movimientos</h3><span class="muted" id="tireEventCount"></span></div>
        <div class="table-wrap" style="max-height:360px;"><table id="tireEventTable"></table></div>
      </div>
    </section>
    <section id="equipos" class="view">
      <div class="panel toolbar">
        <label>Equipo<select id="equipmentSelect"></select></label>
        <label>Servicio<select id="serviceSelect"><option value="">Todos</option></select></label>
        <label>Buscar<input id="filterSearch" placeholder="No. parte, descripcion"></label>
        <label>Estado<select id="statusSelect"><option value="">Todos</option><option>Disponible</option><option>Faltante</option><option>Sin inventario</option></select></label>
        <button class="btn" id="refreshBtn">Actualizar</button>
      </div>
      <div class="table-wrap"><table id="filtersTable"></table></div>
    </section>
    <section id="inventario" class="view">
      <div class="grid2">
        <div class="panel">
          <h3>Concentrado de filtros</h3>
          <div class="toolbar" style="grid-template-columns:1fr 160px;">
            <label>Buscar<input id="inventorySearch" placeholder="No. parte o descripcion"></label>
            <button class="btn secondary" id="exportBtn">Exportar Excel</button>
          </div>
          <div class="table-wrap" style="margin-top:10px;"><table id="inventoryTable"></table></div>
        </div>
        <div class="panel">
          <h3>Entrada / salida / ajuste</h3>
          <div class="movement-grid">
            <label>No. parte<input id="movPart"></label>
            <label>Tipo<select id="movType"><option>ENTRADA</option><option>SALIDA</option><option>AJUSTE</option></select></label>
            <label>Cantidad<input id="movQty" type="number" step="0.01" value="1"></label>
            <label>Unidad<input id="movUnit" value="PZA"></label>
            <label class="wide">Descripcion<input id="movDesc"></label>
            <label>Equipo<input id="movEquipment"></label>
            <label>Servicio<input id="movService"></label>
            <label>Referencia<input id="movRef"></label>
            <label>Usuario<input id="movUser"></label>
            <label class="wide">Notas<textarea id="movNotes" rows="3"></textarea></label>
            <button class="btn wide" id="movementBtn">Guardar movimiento</button>
          </div>
          <h3>Ultimos movimientos</h3>
          <div class="table-wrap" style="max-height:300px;"><table id="movementTable"></table></div>
        </div>
      </div>
    </section>
    <section id="epp" class="view">
      <div class="stats" id="eppStats"></div>
      <datalist id="eppWorkerList"></datalist>
      <div class="panel toolbar">
        <label>Buscar<input id="eppSearch" placeholder="Codigo, descripcion, trabajador"></label>
        <label>Estado<select id="eppStatus"><option value="">Todos</option><option>OK</option><option>BAJO MINIMO</option><option>SIN STOCK</option><option>PROXIMO A REPOSICION</option><option>VENCIDO POR VIDA UTIL</option></select></label>
        <button class="btn" id="eppRefreshBtn">Actualizar EPP</button>
        <button class="btn secondary" id="eppExportBtn">Exportar Excel</button>
      </div>
      <div class="grid2">
        <div class="panel">
          <h3>Inventario EPP</h3>
          <div class="movement-grid">
            <label>Codigo<input id="eppCode"></label>
            <label>Categoria<input id="eppCategory" placeholder="CABEZA, MANOS, RESPIRATORIO"></label>
            <label>Talla<input id="eppSize"></label>
            <label>Unidad<input id="eppUnit" value="PZA"></label>
            <label class="wide">Descripcion<input id="eppDesc"></label>
            <label>Existencia<input id="eppQty" type="number" step="0.01" value="0"></label>
            <label>Minimo<input id="eppMin" type="number" step="0.01" value="0"></label>
            <label>Vida util dias<input id="eppLife" type="number" step="1" value="0"></label>
            <label>Area / riesgo<input id="eppRisk"></label>
            <label>Ubicacion<input id="eppLocation"></label>
            <label class="wide">Notas mantenimiento<input id="eppNotes"></label>
            <label><span>Capacitacion</span><select id="eppTraining"><option value="0">No requerida</option><option value="1">Requerida</option></select></label>
          </div>
          <div class="req-actions">
            <button class="btn secondary" id="eppNewBtn">Nuevo</button>
            <button class="btn" id="eppSaveBtn">Guardar / modificar</button>
            <button class="btn danger" id="eppDeleteBtn">Eliminar</button>
            <button class="btn danger" id="eppDeleteAllBtn">Eliminar todo EPP</button>
            <button class="btn secondary" id="eppQrBtn">QR articulo</button>
          </div>
          <div class="table-wrap" style="max-height:360px; margin-top:10px;"><table id="eppTable"></table></div>
        </div>
        <div class="panel">
          <h3>Entrada / salida / ajuste</h3>
          <div class="movement-grid">
            <label>Fecha<input id="eppMovDate" type="date"></label>
            <label>Tipo<select id="eppMovType"><option>ENTRADA</option><option>SALIDA</option><option>AJUSTE</option></select></label>
            <label>Cantidad<input id="eppMovQty" type="number" step="0.01" value="1"></label>
            <label>Referencia<input id="eppMovRef"></label>
            <label>Trabajador<input id="eppMovWorker" list="eppWorkerList"></label>
            <label>No. empleado<input id="eppMovEmployee"></label>
            <label>Area<input id="eppMovArea"></label>
            <label class="wide">Notas<textarea id="eppMovNotes" rows="2"></textarea></label>
            <button class="btn wide" id="eppMovBtn">Guardar movimiento</button>
          </div>
          <h3>Entrega a trabajador</h3>
          <div class="movement-grid">
            <label>Fecha<input id="eppDelDate" type="date"></label>
            <label>Trabajador<input id="eppDelWorker" list="eppWorkerList"></label>
            <label>No. empleado<input id="eppDelEmployee"></label>
            <label>Cantidad<input id="eppDelQty" type="number" step="0.01" value="1"></label>
            <label>Area / puesto<input id="eppDelArea"></label>
            <label>Recibio<input id="eppDelReceived"></label>
            <label>Firma texto<input id="eppDelSignature"></label>
            <label>Capacitado<select id="eppDelTraining"><option value="1">Si</option><option value="0">No</option></select></label>
            <label>Estado<select id="eppDelCondition"><option>ENTREGADO</option><option>REPOSICION</option><option>DAÑADO</option><option>BAJA</option></select></label>
            <label class="wide">Notas<textarea id="eppDelNotes" rows="2"></textarea></label>
            <button class="btn wide" id="eppDelBtn">Registrar entrega</button>
            <button class="btn secondary wide" id="eppLastPdfBtn">PDF ultima entrega</button>
          </div>
        </div>
      </div>
      <div class="panel">
        <h3>Catalogo de trabajadores</h3>
        <div class="movement-grid">
          <label>No. empleado<input id="eppWorkerEmployee"></label>
          <label>Trabajador<input id="eppWorkerName"></label>
          <label>Area<input id="eppWorkerArea"></label>
          <label>Puesto<input id="eppWorkerPosition"></label>
        </div>
        <div class="req-actions">
          <button class="btn" id="eppWorkerSaveBtn">Guardar trabajador</button>
          <button class="btn danger" id="eppWorkerDeleteBtn">Eliminar trabajador</button>
          <button class="btn secondary" id="eppWorkerUseBtn">Usar en entrega</button>
        </div>
        <div class="table-wrap" style="max-height:220px; margin-top:10px;"><table id="eppWorkersTable"></table></div>
      </div>
      <div class="grid2">
        <div class="panel">
          <h3>Entregas registradas</h3>
          <div class="table-wrap" style="max-height:340px;"><table id="eppDeliveriesTable"></table></div>
        </div>
        <div class="panel">
          <h3>Movimientos EPP</h3>
          <div class="table-wrap" style="max-height:340px;"><table id="eppMovementsTable"></table></div>
        </div>
      </div>
      <div class="panel">
        <h3>Importar inventario EPP desde Excel</h3>
        <p class="muted">Columnas aceptadas: codigo/clave, descripcion, cantidad, unidad, minimo, ubicacion, categoria, talla, vida util dias, area/riesgo.</p>
        <input id="eppImportFile" type="file" accept=".xlsx,.xls">
        <button class="btn danger" id="eppImportBtn">Importar y reemplazar EPP</button>
        <pre id="eppImportResult"></pre>
      </div>
    </section>
    <section id="refacciones" class="view">
      <div class="panel toolbar">
        <label>Equipo<select id="spareEquipment"></select></label>
        <label>Estado<select id="spareStatus"><option value="">Todos</option><option>DISPONIBLE</option><option>FALTANTE</option><option>SIN INVENTARIO</option></select></label>
        <label>Buscar<input id="spareSearch" placeholder="Parte, descripcion, sistema"></label>
        <button class="btn" id="renderSpareBtn">Actualizar</button>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Refacciones por equipo</h3><span class="muted" id="spareCount"></span></div>
        <div class="stats" id="spareStats"></div>
      </div>
      <div class="table-wrap"><table id="spareTable"></table></div>
    </section>
    <section id="auditoria" class="view">
      <div class="panel toolbar">
        <label>Desde<input id="auditStart" type="date"></label>
        <label>Hasta<input id="auditEnd" type="date"></label>
        <label>Modulo<select id="auditModule"><option value="">Todos</option></select></label>
        <label>Buscar<input id="auditSearch" placeholder="Usuario, accion, resumen"></label>
        <button class="btn" id="auditRefreshBtn">Actualizar</button>
      </div>
      <div class="panel">
        <div class="subtle-title"><h3>Auditoria de cambios</h3><span class="muted" id="auditCount"></span></div>
      </div>
      <div class="table-wrap"><table id="auditTable"></table></div>
    </section>
    <section id="importar" class="view">
      <div class="panel">
        <h3>Importar inventario desde Excel</h3>
        <p class="muted">Columnas aceptadas: No. parte, descripcion, cantidad/existencia, unidad, minimo, ubicacion.</p>
        <input id="importFile" type="file" accept=".xlsx,.xls">
        <button class="btn danger" id="importBtn">Importar y reemplazar inventario</button>
        <pre id="importResult"></pre>
      </div>
    </section>
  </main>
  <script>
    let data = { equipment: [], inventory: [], movements: [], summary: {} };
    let portal = { equipment: [], preventives: [], service_history: [], preventive_execution: {records: []}, work_orders: {records: []}, parts_manuals: {manuals: [], rows: [], summary: {}}, audit_log: [], backlog: {items: [], summary: {}, systems: []}, captures: [], availability: [], settings: {}, period: {}, products: [] };
    let products = [];
    let requisitions = [];
    let hoses = { records: [], summary: [], totals: {}, start: "", end: "", period_days: 0 };
    let diesel = { equipment: [], records: [], days: [], rows: [], totals: {}, start: "", end: "", meta_lh: 25 };
    let epp = { items: [], movements: [], deliveries: [], workers: [], summary: {} };
    let currentEppWorkerId = null;
    let currentReqId = null;
    let currentReqItemIndex = null;
    let currentReqItems = [];
    let currentTrackId = null;
    let currentTrackItems = [];
    let currentHoseId = null;
    let currentHoseRecord = null;
    let currentDieselId = null;
    let currentDieselRecord = null;
    let currentCaptureRecord = null;
    let currentCaptureRows = [];
    let currentPreventiveExecutionRecord = null;
    let currentWorkOrderRecord = null;
    let selectedKpiMetric = "availability";
    let monthlyPeriodInitialized = false;
    let kpiSimulationInitialized = false;
    const AUTO_REFRESH_MS = 15000;
    const $ = (id) => document.getElementById(id);
    const apiKey = $("apiKey");
    apiKey.value = localStorage.getItem("mgaFilterApiKey") || "";
    if(apiKey.value.trim() === "X-MGA-API-Key"){
      apiKey.value = "";
      localStorage.removeItem("mgaFilterApiKey");
    }
    apiKey.addEventListener("input", () => localStorage.setItem("mgaFilterApiKey", apiKey.value.trim()));
    apiKey.addEventListener("keydown", (ev) => { if(ev.key === "Enter") load(true).catch(showError); });
    function hasApiKey(show=true){
      const value = apiKey.value.trim();
      if(value && value !== "X-MGA-API-Key") return true;
      apiKey.value = "";
      localStorage.removeItem("mgaFilterApiKey");
      if(show){
        alert("Para modificar inventario pega la clave real. Esta en MGA Mantenimiento > Red > API key cloud.");
        apiKey.focus();
      }
      return false;
    }
    function headers(json=false){ const h = {}; if(apiKey.value.trim()) h["X-MGA-API-Key"] = apiKey.value.trim(); if(json) h["Content-Type"]="application/json"; return h; }
    async function apiError(response){
      const text = await response.text();
      try {
        const payload = JSON.parse(text);
        return payload.detail || text;
      } catch {
        return text;
      }
    }
    function showError(error){ alert(error.message || String(error)); }
    function esc(v){ return String(v ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[c])); }
    function num(v){ const n = Number(v || 0); return Number.isInteger(n) ? String(n) : n.toFixed(2); }
    function one(v){ return `${Number(v || 0).toFixed(1)}`; }
    function pct(v){ return `${one(v)}%`; }
    function statusClass(s){ return s === "Disponible" ? "ok" : (s === "Faltante" ? "bad" : "warn"); }
    function shortText(v, limit=130){
      const text = String(v || "").replace(/\s+/g, " ").trim();
      return text.length > limit ? `${text.slice(0, limit - 1)}...` : text;
    }
    function serviceFiltersText(row){
      if(row.filters_text) return String(row.filters_text);
      if(Array.isArray(row.filters_used)){
        return row.filters_used.map(item => `${item.part_number || ""} ${item.quantity || ""} ${item.unit || ""}`.trim()).filter(Boolean).join("; ");
      }
      if(row.filters_used && String(row.filters_used).trim().startsWith("[")){
        try {
          return JSON.parse(row.filters_used).map(item => `${item.part_number || ""} ${item.quantity || ""} ${item.unit || ""}`.trim()).filter(Boolean).join("; ");
        } catch(_err) {}
      }
      const notes = String(row.notes || "");
      const match = notes.match(/Filtros descontados:\s*([^\n]+)/i);
      return match ? match[1].trim() : "";
    }
    function serviceOilsText(row){
      const raw = row.oils_used || "";
      if(!raw) return "";
      if(typeof raw === "object") return Object.entries(raw).map(([k,v]) => `${k}: ${v}`).join("; ");
      try {
        return Object.entries(JSON.parse(raw)).map(([k,v]) => `${k}: ${v}`).join("; ");
      } catch(_err) {
        return String(raw);
      }
    }
    function setDashboardMode(mode){
      const area = $("kpiPrintArea");
      area.classList.toggle("kpi-format-mode", mode === "format");
      area.classList.toggle("kpi-special-mode", mode === "special");
      area.classList.toggle("kpi-oil-mode", mode === "oil");
      area.classList.toggle("kpi-diesel-mode", mode === "diesel");
      area.classList.remove("simulation");
      $("kpiSideCards").innerHTML = "";
      $("kpiTable").className = "";
      const tableWrap = $("kpiTable").closest(".table-wrap");
      if(tableWrap) tableWrap.classList.remove("oil-bottom-wrap");
    }
    function metricCardHtml(label, value, note, width, bad=false){
      const safeWidth = Math.max(Math.min(Number(width || 0),100),0);
      return `<div class="metric-card ${bad ? "bad" : ""}" style="--ring:${safeWidth * 3.6}deg"><div class="metric-body"><span>${esc(label)}</span><strong>${esc(value)}</strong><small class="muted">${esc(note)}</small><div class="bar-track"><i class="bar-fill" style="width:${safeWidth}%"></i></div></div><div class="metric-ring"><b>${Math.round(safeWidth)}%</b></div></div>`;
    }
    const monthNames = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
    function periodTitle(start){
      const date = parseIsoDate(start || $("kpiStart").value);
      return `${monthNames[date.getMonth()].toUpperCase()} ${date.getFullYear()}`;
    }
    function kpiSheetKind(){
      const group = String($("kpiGroup").value || "").toUpperCase();
      if(group.includes("ACEITE")) return "oil";
      if(group.includes("LLANTA")) return "tire";
      if(group.includes("DIESEL")) return "diesel";
      if(group.includes("REZAGADO")) return "machine";
      return "machine";
    }
    function kpiBaseHeight(kind){
      return kind === "tire" ? 1295 : 927;
    }
    function kpiExactCss(kind="machine", forPrint=false){
      const page = kind === "tire" ? "letter portrait" : "letter landscape";
      return `
        ${forPrint ? `@page { size:${page}; margin:0; } html,body{margin:0;background:white;}` : ""}
        .kpi-sheet{width:1200px;min-height:${kpiBaseHeight(kind)}px;background:#eeeeee;color:#041b40;font-family:Segoe UI,Arial,sans-serif;box-sizing:border-box;padding:0;overflow:visible;}
        .kpi-sheet *{box-sizing:border-box;letter-spacing:0;}
        .kpi-head{height:64px;background:#123d82;color:white;display:flex;align-items:center;justify-content:center;position:relative;border-bottom:4px solid #0aa6a6;}
        .kpi-head h1{margin:0;font-size:30px;line-height:1;font-weight:800;text-transform:uppercase;}
        .kpi-logo{position:absolute;left:34px;top:14px;font-weight:800;font-size:19px;color:white;}
        .kpi-logo::after{content:"";display:block;width:54px;height:4px;background:#e11d48;margin-top:4px;}
        .kpi-title-row{height:54px;display:flex;align-items:center;justify-content:center;color:#06306e;font-size:24px;font-weight:800;background:#f2f2f2;}
        .kpi-board{display:grid;grid-template-columns:330px 1fr 280px;gap:12px;padding:0 26px 12px;}
        .kpi-card-grid{display:grid;grid-template-columns:1fr 1fr;border:1px solid #d8d8d8;background:white;}
        .kpi-card{height:128px;border-right:1px solid #ddd;border-bottom:1px solid #ddd;padding:16px 12px;text-align:center;background:white;}
        .kpi-card h3{margin:0 0 10px;color:#666;font-size:16px;}
        .kpi-card strong{display:block;color:#666;font-size:30px;}
        .kpi-card .bar{height:10px;background:#e5e7eb;margin:14px 4px 8px;border-radius:0;overflow:hidden;}
        .kpi-card .bar i{display:block;height:100%;background:#0aa6a6;}
        .kpi-card.bad .bar i{background:#e11d48;}
        .kpi-card small{color:#5f6774;font-size:12px;}
        .kpi-chart-panel{background:white;padding:16px 20px 10px;min-height:360px;}
        .kpi-tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin:0 0 18px;}
        .kpi-tabs span{border:1px solid #111;padding:8px 10px;font-size:13px;background:white;color:#111;}
        .kpi-tabs span.active{background:#0aa6a6;color:#041b40;}
        .kpi-bars{height:250px;display:flex;gap:18px;align-items:flex-end;border-bottom:1px solid #ccc;border-left:1px solid #eee;padding:18px 14px 0;position:relative;}
        .kpi-bars::before,.kpi-bars::after{content:"";position:absolute;left:14px;right:14px;border-top:1px dashed #999;}
        .kpi-bars::before{top:70px}.kpi-bars::after{top:125px}
        .kpi-bar{width:54px;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;}
        .kpi-bar em{font-style:normal;font-size:12px;color:#303846;margin-bottom:4px;}
        .kpi-bar i{display:block;width:46px;height:var(--h);min-height:4px;background:#0aa6a6;border:1px solid #006f70;}
        .kpi-bar.out i{background:#e11d48;border-color:#be123c;}
        .kpi-bar b{font-size:10px;color:#27364a;margin-top:8px;font-weight:600;text-align:center;max-width:76px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .kpi-report-name{text-align:center;font-size:17px;font-weight:800;color:#000;margin:10px 0 6px;}
        .kpi-days{display:flex;gap:70px;justify-content:center;align-items:center;margin-bottom:8px;font-size:13px;color:#777;font-weight:700;}
        .kpi-days b{display:inline-block;background:white;color:#000;min-width:80px;padding:6px 20px;margin-left:10px;}
        .kpi-table-wrap{padding:0 90px 22px;}
        .kpi-exact-table{width:100%;border-collapse:collapse;background:white;font-size:10px;color:#334155;}
        .kpi-exact-table th{background:white;color:#555;border:1px solid #111;font-weight:800;text-align:center;padding:6px 4px;}
        .kpi-exact-table td{border:1px solid #111;text-align:center;padding:5px 4px;background:white;}
        .kpi-exact-table .badtext{color:#e11d48}.kpi-exact-table .oktext{color:#0aa6a6}
        .diesel-export{background:#f4f7fb;padding:18px;color:#172033;}
        .diesel-export-head{height:82px;display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:12px;padding:16px 18px;border-radius:8px;background:#071f49;color:white;}
        .diesel-export-head h1{margin:0;font-size:30px;line-height:1;font-weight:800;text-transform:uppercase;}
        .diesel-export-head span{display:block;color:#93c5fd;font-size:12px;font-weight:800;text-transform:uppercase;}
        .diesel-export-head em{display:block;margin-top:5px;color:#dbeafe;font-size:13px;font-style:normal;}
        .diesel-export-score{min-width:260px;text-align:right;}
        .diesel-export-score strong{display:block;font-size:31px;line-height:1;}
        .diesel-export-score small{display:block;margin-top:5px;color:#dbeafe;font-size:12px;}
        .diesel-card-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin-bottom:12px;}
        .diesel-card{position:relative;min-height:106px;overflow:hidden;border:1px solid #dbe3ef;border-radius:8px;padding:12px 12px 10px;background:linear-gradient(180deg,#fff,#f8fbff);box-shadow:0 8px 18px rgba(7,31,73,.06);}
        .diesel-card::before{content:"";position:absolute;inset:0 auto 0 0;width:4px;background:#009c9a;}
        .diesel-card.is-bad::before{background:#c81e1e;}
        .diesel-card span{display:block;color:#667085;font-size:10px;font-weight:800;text-transform:uppercase;}
        .diesel-card strong{display:block;margin-top:7px;color:#071f49;font-size:23px;line-height:1;}
        .diesel-card small{display:block;min-height:24px;margin-top:6px;color:#475569;font-size:11px;line-height:1.15;}
        .diesel-meter{height:6px;margin-top:8px;overflow:hidden;border-radius:999px;background:#e8eef6;}
        .diesel-meter i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,#009c9a,#18b7a6);}
        .diesel-card.is-bad .diesel-meter i{background:linear-gradient(90deg,#e11d48,#c81e1e);}
        .diesel-visual-grid{display:grid;grid-template-columns:1.45fr .95fr;gap:12px;align-items:stretch;margin-bottom:12px;}
        .diesel-panel{min-width:0;border:1px solid #dbe3ef;border-radius:8px;padding:13px;background:white;box-shadow:0 8px 18px rgba(7,31,73,.05);}
        .diesel-panel-head{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:10px;}
        .diesel-panel-head span{color:#071f49;font-size:14px;font-weight:800;}
        .diesel-panel-head b{color:#667085;font-size:10px;font-weight:800;text-transform:uppercase;}
        .diesel-bars-list{display:grid;gap:7px;}
        .diesel-bar-row{display:grid;grid-template-columns:128px minmax(110px,1fr) 72px 66px;gap:9px;align-items:center;min-height:28px;}
        .diesel-bar-label{min-width:0;}
        .diesel-bar-label b{display:block;overflow:hidden;color:#071f49;font-size:12px;text-overflow:ellipsis;white-space:nowrap;}
        .diesel-bar-label span{display:block;overflow:hidden;color:#667085;font-size:9px;font-weight:800;text-overflow:ellipsis;text-transform:uppercase;white-space:nowrap;}
        .diesel-bar-track{height:11px;overflow:hidden;border-radius:999px;background:#edf2f7;}
        .diesel-bar-track i{display:block;height:100%;min-width:3px;border-radius:999px;background:linear-gradient(90deg,#009c9a,#18b7a6);}
        .diesel-bar-row.is-bad .diesel-bar-track i{background:linear-gradient(90deg,#e11d48,#c81e1e);}
        .diesel-bar-row strong{color:#172033;font-size:11px;text-align:right;white-space:nowrap;}
        .diesel-bar-row em{color:#475569;font-size:10px;font-style:normal;text-align:right;white-space:nowrap;}
        .diesel-performance-grid{display:grid;gap:10px;}
        .diesel-target{border:1px solid #e2e8f0;border-radius:8px;padding:12px;background:#f8fafc;}
        .diesel-target-top{display:flex;align-items:end;justify-content:space-between;gap:10px;margin-bottom:10px;}
        .diesel-target-top span{color:#667085;font-size:10px;font-weight:800;text-transform:uppercase;}
        .diesel-target-top strong{color:#071f49;font-size:30px;line-height:1;white-space:nowrap;}
        .diesel-target-top small{color:#475569;font-size:11px;text-align:right;}
        .diesel-target-meter{position:relative;height:15px;overflow:hidden;border-radius:999px;background:#e8eef6;}
        .diesel-target-meter i{display:block;height:100%;min-width:3px;border-radius:999px;background:linear-gradient(90deg,#009c9a,#f59e0b,#e11d48);}
        .diesel-target-meter .diesel-target-marker{position:absolute;top:-5px;bottom:-5px;left:var(--target);width:2px;background:#071f49;box-shadow:0 0 0 2px rgba(255,255,255,.85);}
        .diesel-split{border:1px solid #e2e8f0;border-radius:8px;padding:12px;background:white;}
        .diesel-split h4{margin:0 0 9px;color:#071f49;font-size:13px;}
        .diesel-split-track{display:flex;height:18px;overflow:hidden;border-radius:999px;background:#e8eef6;}
        .diesel-split-track i{display:block;height:100%;min-width:0;}
        .diesel-split-track .mga{background:#009c9a;}
        .diesel-split-track .pro{background:#f59e0b;}
        .diesel-split-legend{display:grid;gap:6px;margin-top:9px;}
        .diesel-split-legend > span{display:flex;align-items:center;justify-content:space-between;gap:10px;color:#475569;font-size:11px;}
        .diesel-split-legend em{font-style:normal;}
        .diesel-split-legend i{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px;vertical-align:-1px;}
        .diesel-split-legend .mga{background:#009c9a;}
        .diesel-split-legend .pro{background:#f59e0b;}
        .diesel-watch-list{display:grid;gap:6px;}
        .diesel-watch-item{display:flex;justify-content:space-between;gap:10px;padding:7px 9px;border-radius:7px;background:#f8fafc;color:#475569;font-size:11px;}
        .diesel-watch-item b{color:#071f49;}
        .diesel-empty{display:grid;min-height:190px;place-items:center;color:#667085;font-weight:700;}
        .diesel-table-wrap{border:0;background:transparent;overflow:visible;}
        .diesel-table{width:100%;border-collapse:separate;border-spacing:0;overflow:hidden;border:1px solid #dbe3ef;border-radius:8px;background:white;color:#172033;}
        .diesel-table th{position:static;padding:8px 7px;background:#e8eef7;color:#243042;font-size:10px;text-align:center;text-transform:uppercase;}
        .diesel-table td{padding:7px;border-bottom:1px solid #e5ebf3;color:#334155;font-size:11px;text-align:center;vertical-align:middle;}
        .diesel-table tbody tr:last-child td{border-bottom:0;}
        .diesel-table .diesel-eq{color:#071f49;font-weight:800;text-align:left;}
        .diesel-table .diesel-number{font-variant-numeric:tabular-nums;text-align:right;}
        .diesel-table .diesel-total td{background:#f0fdfa;color:#071f49;font-weight:800;}
        .pill{display:inline-block;padding:2px 7px;border-radius:999px;font-weight:700;font-size:10px;}
        .ok{color:#047857;background:#d1fae5;}
        .bad{color:#b91c1c;background:#fee2e2;}
        .warn{color:#92400e;background:#fef3c7;}
        .oil-sheet{background:#f6f8fb;}
        .oil-stats{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;padding:14px 18px;}
        .oil-stat{background:white;border:1px solid #d6dee9;padding:12px;text-align:center;}
        .oil-stat strong{display:block;color:#0b2f6f;font-size:26px}.oil-stat span{color:#667085;font-weight:700;font-size:12px;}
        .oil-grid{display:grid;grid-template-columns:1fr 1.25fr;gap:14px;padding:0 18px 16px;}
        .oil-panel{background:white;border:1px solid #d6dee9;padding:14px;}
        .oil-panel h2{margin:0 0 12px;color:#0b2f6f;font-size:20px;}
        .oil-bars{height:260px;display:flex;align-items:flex-end;gap:16px;border-bottom:1px solid #ccd5e1;padding:10px 10px 0;}
        .oil-bar{width:50px;text-align:center;font-size:10px;color:#334155}.oil-bar i{display:block;height:var(--h);background:#0aa6a6;margin:4px auto 7px;width:38px;}
        .oil-export{padding:12px 14px 18px;background:#eeeeee;color:#333;}
        .oil-export .oil-title{display:grid;grid-template-columns:1fr 2fr 1fr;align-items:center;height:34px;margin:0 0 12px;background:white;color:#333;text-align:center;font-size:20px;font-weight:800;}
        .oil-export .oil-month{font-size:14px;}
        .oil-export .kpi-format-board{display:grid;grid-template-columns:minmax(260px,.7fr) minmax(430px,1.2fr) minmax(260px,.7fr);gap:22px;align-items:start;}
        .oil-export .kpi-side{display:grid;grid-template-columns:1fr;gap:40px;border:0;background:transparent;align-self:start;}
        .oil-metric-section{background:white;border:1px solid #e5e7eb;}
        .oil-metric-title{height:34px;display:flex;align-items:center;justify-content:center;color:#707780;font-weight:800;font-size:18px;}
        .oil-metric-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;background:#eeeeee;}
        .oil-metric-cell{min-height:92px;background:white;display:grid;align-content:center;justify-items:center;gap:8px;padding:8px 6px;}
        .oil-metric-cell strong{color:#777d86;font-size:31px;line-height:1;}
        .oil-metric-line{height:10px;width:100%;background:#eef1f4;}
        .oil-metric-line i{display:block;height:100%;width:100%;background:#009c9a;}
        .oil-metric-line.oil-red i{background:#d76f75;}
        .oil-metric-line.oil-darkred i{background:#a40000;}
        .oil-metric-cell span{color:#4b5563;font-size:12px;}
        .oil-export .chart{min-height:330px;padding:14px 18px 10px;border:1px solid #d8dee8;border-radius:0;background:white;display:block;overflow:hidden;}
        .oil-chart-grid{display:grid;grid-template-columns:42px 1fr;grid-template-rows:250px 38px;column-gap:8px;}
        .oil-axis{grid-row:1;display:flex;flex-direction:column;justify-content:space-between;align-items:end;padding:0 2px 0 0;color:#111;font-size:12px;}
        .oil-plot{position:relative;grid-column:2;grid-row:1;display:flex;align-items:stretch;gap:14px;padding:0 8px;border-bottom:1px solid #d9d9d9;background:repeating-linear-gradient(to top, transparent 0, transparent 49px, #d9d9d9 50px);}
        .oil-cluster{flex:1 1 62px;min-width:54px;display:grid;grid-template-rows:1fr auto;justify-items:center;gap:8px;}
        .oil-bars-stack{height:100%;display:flex;align-items:flex-end;gap:2px;}
        .oil-series-bar{width:9px;min-height:1px;position:relative;}
        .oil-series-bar b{position:absolute;left:50%;transform:translateX(-50%);top:-15px;color:#111;font-size:10px;font-weight:500;white-space:nowrap;}
        .oil-cluster-label{color:#111;font-size:12px;text-align:center;white-space:nowrap;}
        .oil-legend{grid-column:2;grid-row:2;display:flex;align-items:end;justify-content:center;gap:18px;color:#333;font-size:12px;}
        .oil-legend span{display:flex;align-items:center;gap:5px;white-space:nowrap;}
        .oil-legend i{display:block;width:10px;height:10px;}
        .oil-export .oil-bottom-wrap{margin-top:28px;border:0;background:transparent;overflow:visible;}
        .oil-bottom-grid{width:100%;border-collapse:separate;border-spacing:0;background:transparent;}
        .oil-bottom-grid>tbody>tr>td{border:0;padding:0 8px;vertical-align:top;}
        .oil-report-cell{width:72%;}
        .oil-order-cell{width:28%;}
        .oil-report-header{background:#f3f3f3;text-align:center;padding:10px 8px 8px;}
        .oil-report-header h3{margin:0 0 8px;color:#111;font-size:17px;}
        .oil-days{display:flex;justify-content:center;gap:72px;color:#707780;font-size:12px;font-weight:800;}
        .oil-days b{display:inline-block;min-width:58px;margin-left:8px;padding:5px 16px;background:white;color:#111;}
        .oil-report-table{width:100%;border-collapse:collapse;background:white;color:#555;}
        .oil-report-table th,.oil-report-table td{border:1px solid #111;padding:4px 5px;font-size:10px;text-align:center;vertical-align:middle;}
        .oil-report-table th{position:static;background:white;color:#555;font-weight:800;text-transform:none;line-height:1.05;}
        .oil-report-table td{background:#efefef;}
        .oil-report-table .oil-subtotal td{background:#ffd966;}
        .oil-report-table .oil-total td{background:#fff200;}
        .oil-report-table .oil-zero{color:#f05b5b;font-weight:800;}
        .oil-order-panel{background:white;border:1px solid #cbd5e1;min-height:330px;}
        .oil-order-title{background:#0d3272;color:white;text-align:center;font-weight:800;padding:13px 8px;}
        .oil-order-kpis{display:grid;grid-template-columns:repeat(3,1fr);border-bottom:1px solid #dbe3ef;}
        .oil-order-kpi{background:#f8fafc;border-right:1px solid #e2e8f0;text-align:center;padding:12px 4px 9px;}
        .oil-order-kpi:last-child{border-right:0;}
        .oil-order-kpi strong{display:block;color:#d6335c;font-size:12px;}
        .oil-order-kpi span{color:#475569;font-size:12px;}
        .oil-order-table{width:100%;border-collapse:collapse;}
        .oil-order-table th,.oil-order-table td{border-bottom:1px solid #e2e8f0;padding:7px 6px;font-size:11px;text-align:center;}
        .oil-order-table th{position:static;background:#e2e8f0;color:#0f172a;text-transform:none;}
        .oil-order-table td:first-child{text-align:left;font-weight:700;color:#334155;}
        .oil-order-table .oil-order-hot{color:#d6335c;font-weight:800;}
        .tire-sheet{background:#f5f8fc;}
        .tire-stats{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;padding:14px 12px;}
        .tire-stat{background:white;border:1px solid #d8dee8;padding:12px;text-align:center;clip-path:polygon(8px 0,100% 0,calc(100% - 8px) 100%,0 100%);}
        .tire-stat strong{display:block;font-size:26px;color:#0b2f6f}.tire-stat span{font-size:12px;color:#667085;font-weight:700;}
        .tire-table{width:100%;border-collapse:collapse;background:white;font-size:15px;color:#5c6673;}
        .tire-table th{background:#e8eef7;border:1px solid #cdd5df;padding:9px;color:#516173;text-align:center;}
        .tire-table td{border:1px solid #d8dee8;padding:8px;text-align:center;}
        .life-cell{display:flex;align-items:center;gap:8px;justify-content:flex-end}.life-bar{width:78px;height:18px;background:#e5e7eb}.life-bar i{display:block;height:100%;background:#0aa6a6;width:var(--w);}
        ${forPrint ? `.kpi-sheet{width:100%;min-height:0;overflow:visible;transform-origin:top left;} body{display:block;} .kpi-board{grid-template-columns:28% 45% 25%;gap:8px;padding:0 14px 10px;} .kpi-table-wrap{padding:0 30px 18px;} .kpi-card{height:116px;padding:12px 8px;} .kpi-card strong{font-size:25px;} .kpi-chart-panel{min-height:330px;padding:12px 14px 8px;} .kpi-bars{height:230px;gap:10px;} .kpi-bar{width:46px;} .kpi-bar i{width:38px;} .oil-export .kpi-format-board{grid-template-columns:25% 47% 25%;gap:12px;} .oil-report-table th,.oil-report-table td{font-size:8.5px;padding:3px 4px;} .oil-order-table th,.oil-order-table td{font-size:9px;padding:5px 4px;} .tire-table{font-size:11px;} .tire-table th,.tire-table td{padding:5px;} .print-note{display:none;}` : ""}
      `;
    }
    function metricProgress(value, target, inverse=false){
      const v = Number(value || 0), t = Math.max(Number(target || 1), 1);
      return inverse ? Math.min((t / Math.max(v, 0.1)) * 100, 100) : Math.min((v / t) * 100, 100);
    }
    function exactMachineHtml(){
      const report = simulatedKpiReport(calculateKpiRows());
      const settings = currentKpiSettings();
      const targets = {
        availability:Number(settings.meta_availability || 85),
        utilization:Number(settings.meta_utilization || 75),
        reliability:Number(settings.meta_reliability || 80),
        tmef:Number(settings.meta_tmef || 8),
        tmpr:Number(settings.meta_tmpr || 4),
      };
      const metric = kpiMetricTabs.some(item => item.key === selectedKpiMetric) ? selectedKpiMetric : "availability";
      const metricLabel = kpiMetricTabs.find(item => item.key === metric)?.label || "% Disponibilidad";
      const axisMax = kpiMetricAxisMax(metric, report.rows, targets[metric]);
      const chartRows = report.rows.slice(0, 8);
      const tabs = kpiMetricTabs.map(item => `<span class="${item.key === metric ? "active" : ""}">${esc(item.label)}</span>`).join("");
      const bars = chartRows.map(row => {
        const value = kpiMetricValue(row, metric);
        const height = Math.max(Math.min(value / Math.max(axisMax, 1), 1) * 210, 4);
        return `<div class="kpi-bar ${row.out ? "out" : ""}"><em>${esc(kpiMetricText(row, metric))}</em><i style="--h:${height}px"></i><b>${esc(row.code || "-")}</b></div>`;
      }).join("") || `<p>Sin datos KPI.</p>`;
      const tableRows = report.rows.map(row => `<tr><td>${esc(row.code)}</td><td>${esc(row.description)}</td><td>${one(row.period)}</td><td>${one(row.mp)}</td><td>${one(row.mc)}</td><td>${one(row.worked)}</td><td>${num(row.stops)}</td><td class="${row.availability < targets.availability ? "badtext" : "oktext"}">${esc(row.availabilityText)}</td><td class="${row.utilization < targets.utilization ? "badtext" : "oktext"}">${esc(row.utilizationText)}</td><td class="${row.reliability < targets.reliability ? "badtext" : "oktext"}">${pct(row.reliability)}</td><td>${one(row.tmef)}</td><td>${one(row.tmpr)}</td><td>${esc(row.out ? "FUERA" : row.status)}</td></tr>`).join("");
      return `<section class="kpi-sheet">
        <div class="kpi-title-row"><div class="kpi-logo">MGA</div>${esc(report.group)}</div>
        <div class="kpi-board">
          <div class="kpi-card-grid">
            <div class="kpi-card ${report.totals.availability < targets.availability ? "bad" : ""}"><h3>% Disponibilidad</h3><strong>${pct(report.totals.availability)}</strong><div class="bar"><i style="width:${metricProgress(report.totals.availability, targets.availability)}%"></i></div><small>Meta ${pct(targets.availability)}</small></div>
            <div class="kpi-card"><h3>Meta</h3><strong>${pct(targets.availability)}</strong><div class="bar"><i style="width:${targets.availability}%"></i></div><small>${one(report.totals.availability - targets.availability)}%</small></div>
            <div class="kpi-card ${report.totals.utilization < targets.utilization ? "bad" : ""}"><h3>% Utilizacion</h3><strong>${pct(report.totals.utilization)}</strong><div class="bar"><i style="width:${metricProgress(report.totals.utilization, targets.utilization)}%"></i></div><small>Meta ${pct(targets.utilization)}</small></div>
            <div class="kpi-card"><h3>Meta</h3><strong>${pct(targets.utilization)}</strong><div class="bar"><i style="width:${targets.utilization}%"></i></div><small>${one(report.totals.utilization - targets.utilization)}%</small></div>
          </div>
          <div class="kpi-chart-panel"><div class="kpi-tabs">${tabs}</div><div class="kpi-bars">${bars}</div><div style="text-align:center;margin-top:8px;font-weight:700;color:#52627a">${esc(metricLabel)}</div></div>
          <div class="kpi-card-grid">
            <div class="kpi-card ${report.totals.reliability < targets.reliability ? "bad" : ""}"><h3>Confiabilidad</h3><strong>${pct(report.totals.reliability)}</strong><div class="bar"><i style="width:${metricProgress(report.totals.reliability, targets.reliability)}%"></i></div><small>Meta ${pct(targets.reliability)}</small></div>
            <div class="kpi-card ${report.totals.tmef < targets.tmef ? "bad" : ""}"><h3>TMEF</h3><strong>${one(report.totals.tmef)} hrs</strong><div class="bar"><i style="width:${metricProgress(report.totals.tmef, targets.tmef)}%"></i></div><small>Meta ${one(targets.tmef)} h</small></div>
            <div class="kpi-card ${report.totals.tmpr > targets.tmpr ? "bad" : ""}"><h3>TMPR</h3><strong>${one(report.totals.tmpr)} hrs</strong><div class="bar"><i style="width:${metricProgress(report.totals.tmpr, targets.tmpr, true)}%"></i></div><small>Meta ${one(targets.tmpr)} h</small></div>
            <div class="kpi-card"><h3>Meta Conf.</h3><strong>${pct(targets.reliability)}</strong><div class="bar"><i style="width:${targets.reliability}%"></i></div><small>${one(report.totals.reliability - targets.reliability)}%</small></div>
          </div>
        </div>
        <div class="kpi-report-name">REPORTE SEMANAL DE INDICADORES</div>
        <div class="kpi-days"><span>Dia Inicial:<b>${Number(String(report.start).slice(-2))}</b></span><span>Dia Final:<b>${Number(String(report.end).slice(-2))}</b></span></div>
        <div class="kpi-table-wrap"><table class="kpi-exact-table"><thead><tr><th># Eco</th><th>Equipo</th><th>Hrs Periodo</th><th>Hrs MP</th><th>Hrs MC</th><th>Hrs Trab</th><th># Paradas</th><th>% Disp</th><th>% Util</th><th>Confiabilidad</th><th>TMEF</th><th>TMPR</th><th>Estatus</th></tr></thead><tbody>${tableRows}<tr><td></td><td><b>Total ${esc(report.group)}</b></td><td><b>${one(report.totals.period)}</b></td><td><b>${one(report.totals.mp)}</b></td><td><b>${one(report.totals.mc)}</b></td><td><b>${one(report.totals.worked)}</b></td><td><b>${num(report.totals.stops)}</b></td><td><b>${pct(report.totals.availability)}</b></td><td><b>${pct(report.totals.utilization)}</b></td><td><b>${pct(report.totals.reliability)}</b></td><td><b>${one(report.totals.tmef)}</b></td><td><b>${one(report.totals.tmpr)}</b></td><td></td></tr></tbody></table></div>
      </section>`;
    }
    function exactOilHtml(){
      const report = oilRowsForPeriod();
      const accStart = `${String(report.end || report.start).slice(0,4)}-01-01`;
      const accumulated = oilRowsForRange(accStart, report.end, report.cols);
      const month = oilMonthLabel(report.end || report.start);
      return `<section class="kpi-sheet oil-export">
        <div class="oil-title"><span class="oil-month">${esc(month)}</span><span>Consumo de aceite de equipos "Providencia"</span><span class="oil-month">${esc(month)}</span></div>
        <div class="kpi-format-board">
          <div class="kpi-side">
            ${oilMetricSection("Consumo de Aceite HCO", report.totals.oil_hco_iso68, accumulated.totals.oil_hco_iso68, "teal")}
            ${oilMetricSection("Consumo de Aceite SAE 30", report.totals.oil_trans_sae30, accumulated.totals.oil_trans_sae30, "red")}
          </div>
          <div class="chart">${oilChartHtml(report)}</div>
          <div class="kpi-side">
            ${oilMetricSection("Consumo de Aceite de Motor", report.totals.oil_motor_15w40, accumulated.totals.oil_motor_15w40, "red")}
            ${oilMetricSection("Consumo de Aceite SAE 50", report.totals.oil_sae50, accumulated.totals.oil_sae50, "red")}
          </div>
        </div>
        <div class="table-wrap oil-bottom-wrap"><table class="oil-bottom-grid"><tbody><tr><td class="oil-report-cell">${oilReportTableHtml(report)}</td><td class="oil-order-cell">${oilOrderPanelHtml(report)}</td></tr></tbody></table></div>
      </section>`;
    }
    function exactTireHtml(){
      const tire = portal.tire_kpi || {};
      const rows = Array.isArray(tire.rows) ? tire.rows : [];
      const summary = tire.summary || {};
      const tableRows = rows.map(row => {
        const value = Math.max(Math.min(Number(row.life_percent || row.tread_remaining_percent || 0), 100), 0);
        const status = String(row.control_status || "");
        const color = status === "OK" ? "#0aa6a6" : (status === "PROXIMA" || status === "REVISION" ? "#b45309" : "#e11d48");
        return `<tr><td>${esc(row.equipment_code)}</td><td>${esc(row.tire_code)}</td><td>${esc(row.position)}</td><td>${one(row.hours_used)}</td><td>${one(row.life_remaining_hours)}</td><td><div class="life-cell"><span>${pct(value)}</span><div class="life-bar"><i style="--w:${value}%"></i></div></div></td><td>${pct(value)}</td><td style="color:${color};font-weight:800">${esc(status || "S/D")}</td></tr>`;
      }).join("");
      return `<section class="kpi-sheet tire-sheet">
        <div class="kpi-head"><div class="kpi-logo">MGA</div><h1>VIDA UTIL DE LLANTAS - ${esc(periodTitle($("kpiStart").value))}</h1></div>
        <div class="tire-stats">
          <div class="tire-stat"><strong>${summary.total || rows.length || 0}</strong><span>Llantas</span></div>
          <div class="tire-stat"><strong>${pct(summary.avg_life || 0)}</strong><span>Vida prom.</span></div>
          <div class="tire-stat"><strong>${summary.critical || 0}</strong><span>Criticas</span></div>
          <div class="tire-stat"><strong>${summary.soon || 0}</strong><span>Proximas</span></div>
          <div class="tire-stat"><strong>${one(summary.avg_remaining_hours || 0)}</strong><span>Hrs rest. prom.</span></div>
        </div>
        <div style="padding:0 12px 18px;"><table class="tire-table"><thead><tr><th>Equipo</th><th>Llanta</th><th>Pos.</th><th>Hrs uso</th><th>Hrs rest.</th><th>% vida</th><th>% piso</th><th>KPI</th></tr></thead><tbody>${tableRows || `<tr><td colspan="8">Sin llantas registradas</td></tr>`}</tbody></table></div>
      </section>`;
    }
    function exactDieselHtml(){
      const report = dieselKpiRowsForPeriod();
      const avg = report.totals.rendimiento_lh;
      return `<section class="kpi-sheet diesel-export">
        <div class="diesel-export-head">
          <div><span>MGA Mantenimiento</span><h1>KPI Diesel</h1><em>${esc(report.start)} a ${esc(report.end)}</em></div>
          <div class="diesel-export-score"><span>Rendimiento promedio</span><strong>${avg == null ? "S/H" : `${one(avg)} L/H`}</strong><small>Meta ${one(report.meta)} L/H</small></div>
        </div>
        <div class="diesel-card-grid">${dieselCardsHtml(report)}</div>
        ${dieselVisualHtml(report)}
        <div class="diesel-table-wrap"><table class="diesel-table">${dieselTableHtml(report)}</table></div>
      </section>`;
    }
    function buildExactKpiHtml(){
      const kind = kpiSheetKind();
      if(kind === "oil") return {kind, width:1200, height:927, html:exactOilHtml()};
      if(kind === "tire") return {kind, width:1200, height:1295, html:exactTireHtml()};
      if(kind === "diesel") return {kind, width:1200, height:927, html:exactDieselHtml()};
      return {kind, width:1200, height:927, html:exactMachineHtml()};
    }
    function printExactKpi(){
      const doc = buildExactKpiHtml();
      const win = window.open("", "_blank");
      if(!win) return alert("Permite ventanas emergentes para imprimir el KPI.");
      win.document.open();
      win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>Formato KPI</title><style>${kpiExactCss(doc.kind, true)}</style></head><body>${doc.html}<scr` + `ipt>window.onload=function(){setTimeout(function(){window.focus();window.print();},500);};</scr` + `ipt></body></html>`);
      win.document.close();
    }
    function measureExactKpiHeight(doc, css){
      const host = document.createElement("div");
      host.style.position = "absolute";
      host.style.left = "-20000px";
      host.style.top = "0";
      host.style.width = `${doc.width}px`;
      host.style.background = "#fff";
      host.style.pointerEvents = "none";
      host.style.zIndex = "-1";
      host.innerHTML = `<style>${css}</style>${doc.html}`;
      document.body.appendChild(host);
      const sheet = host.querySelector(".kpi-sheet");
      const rect = sheet ? sheet.getBoundingClientRect() : {height: doc.height};
      const height = Math.ceil(Math.max(doc.height, sheet ? sheet.scrollHeight : 0, sheet ? sheet.offsetHeight : 0, rect.height || 0)) + 4;
      document.body.removeChild(host);
      return height;
    }
    async function downloadKpiImage(){
      const doc = buildExactKpiHtml();
      const css = kpiExactCss(doc.kind, false);
      const height = measureExactKpiHeight(doc, css);
      const safeHtml = doc.html.replaceAll("<br>", "<br/>");
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${doc.width}" height="${height}" viewBox="0 0 ${doc.width} ${height}"><foreignObject width="${doc.width}" height="${height}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:${doc.width}px;min-height:${height}px"><style>${css}</style>${safeHtml}</div></foreignObject></svg>`;
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = doc.width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, doc.width, height);
        ctx.drawImage(image, 0, 0);
        const a = document.createElement("a");
        a.href = canvas.toDataURL("image/png");
        a.download = `Formato_KPI_${($("kpiGroup").value || "KPI").replaceAll(" ","_")}.png`;
        a.click();
      };
      image.onerror = () => alert("No se pudo generar la imagen KPI.");
      image.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
    }
    async function downloadKpiExcel(){
      const group = $("kpiGroup").value || "Todos los equipos";
      const normalized = group.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
      if(normalized.includes("ACEITE") || normalized.includes("LLANTA") || normalized.includes("DIESEL")){
        alert("Excel editable disponible para Barrenacion, Rezagado, Acarreo y Utilitario.");
        return;
      }
      const params = new URLSearchParams({
        group,
        start: $("kpiStart").value || "",
        end: $("kpiEnd").value || ""
      });
      if(kpiSimulationActive()){
        params.set("simulation", "true");
        params.set("sim_name", $("kpiSimName").value || "Escenario KPI");
        params.set("sim_period_percent", String(simNumber("kpiSimPeriod", 100)));
        params.set("sim_worked_percent", String(simNumber("kpiSimWorked", 100)));
        params.set("sim_mp_percent", String(simNumber("kpiSimMp", 100)));
        params.set("sim_mc_percent", String(simNumber("kpiSimMc", 100)));
        params.set("sim_stops_percent", String(simNumber("kpiSimStops", 100)));
        params.set("sim_mission_hours", String(simNumber("kpiSimMission", (portal.settings || {}).reliability_mission_hours || 24)));
        params.set("sim_meta_availability", String(simNumber("kpiSimMetaAvailability", (portal.settings || {}).meta_availability || 85)));
        params.set("sim_meta_utilization", String(simNumber("kpiSimMetaUtilization", (portal.settings || {}).meta_utilization || 75)));
        params.set("sim_meta_reliability", String(simNumber("kpiSimMetaReliability", (portal.settings || {}).meta_reliability || 80)));
        params.set("sim_meta_tmef", String(simNumber("kpiSimMetaTmef", (portal.settings || {}).meta_tmef || 8)));
        params.set("sim_meta_tmpr", String(simNumber("kpiSimMetaTmpr", (portal.settings || {}).meta_tmpr || 4)));
      }
      const response = await fetch(`/api/kpi-format/excel?${params.toString()}`, {headers: headers(), cache: "no-store"});
      if(!response.ok) throw new Error(await apiError(response));
      const blob = await response.blob();
      const link = document.createElement("a");
      const safeGroup = group.replace(/^Equipos de\s+/i, "").replace(/[^\w.-]+/g, "_").replace(/^_+|_+$/g, "") || "KPI";
      const start = $("kpiStart").value || "inicio";
      const end = $("kpiEnd").value || "fin";
      link.href = URL.createObjectURL(blob);
      link.download = `${kpiSimulationActive() ? "KPI_SIMULACION" : "KPI"}_${safeGroup}_${start}_${end}.xlsx`;
      document.body.appendChild(link);
      link.click();
      URL.revokeObjectURL(link.href);
      link.remove();
    }
    async function load(){
      const [r, p, prod, req, hosePayload, dieselPayload] = await Promise.all([
        fetch("/api/filter-inventory", {headers: headers(), cache:"no-store"}),
        fetch("/api/portal", {headers: headers(), cache:"no-store"}),
        fetch("/api/products?limit=25000", {headers: headers(), cache:"no-store"}),
        fetch("/api/requisitions", {headers: headers(), cache:"no-store"}),
        fetch("/api/hose-changes", {headers: headers(), cache:"no-store"}),
        fetch("/api/diesel", {headers: headers(), cache:"no-store"})
      ]);
      if(!r.ok) throw new Error(await apiError(r));
      if(!p.ok) throw new Error(await apiError(p));
      if(!prod.ok) throw new Error(await apiError(prod));
      if(!req.ok) throw new Error(await apiError(req));
      if(!hosePayload.ok) throw new Error(await apiError(hosePayload));
      if(!dieselPayload.ok) throw new Error(await apiError(dieselPayload));
      data = await r.json();
      portal = await p.json();
      products = (await prod.json()).products || [];
      const reqPayload = await req.json();
      hoses = await hosePayload.json();
      diesel = await dieselPayload.json();
      epp = { items: [], movements: [], deliveries: [], workers: [], summary: {} };
      requisitions = reqPayload.requisitions || [];
      if(!currentReqId && !$("reqFolio").value) newRequisition(reqPayload.next_folio);
      renderAll();
    }
    function renderStats(){
      const s = data.summary || {};
      $("stats").innerHTML = [
        ["Equipos", s.equipment || 0],
        ["Filtros FS", s.filters || 0],
        ["Partes almacen", s.inventory_parts || 0],
        ["Faltantes", s.shortage_filters || 0],
        ["Sin inventario", s.unknown_filters || 0],
      ].map(([k,v], idx) => {
        const seed = Math.max(Number(v || 0), 1);
        const bars = [0,1,2,3,4].map(step => `<i style="height:${10 + ((seed + idx * 7 + step * 9) % 18)}px"></i>`).join("");
        return `<div class="stat"><span class="stat-icon"></span><div><strong>${v}</strong><span>${esc(k)}</span></div><div class="stat-spark">${bars}</div></div>`;
      }).join("");
    }
    function activateTab(tabId){
      const button = document.querySelector(`.tabs button[data-tab="${tabId}"]`);
      if(!button) return;
      document.querySelectorAll(".tabs button").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
      button.classList.add("active");
      $(tabId).classList.add("active");
      $("stats").style.display = tabId === "dashboard" ? "" : "none";
    }
    function executiveAlerts(){
      const preventives = preventiveRowsWithWebClosures(portal.preventives || []);
      const overdue = preventives.filter(row => ["VENCIDO","URGENTE"].includes(String(row.status || "").toUpperCase()));
      const servicesOpen = preventiveExecutionRows().filter(row => !isPreventiveClosed(row) && String(row.status || "").toUpperCase() !== "CANCELADO");
      const openOrders = workOrderRows().filter(row => !workOrderClosed(row));
      const spareRows = (portal.parts_manuals && Array.isArray(portal.parts_manuals.rows)) ? portal.parts_manuals.rows : [];
      const spareShort = spareRows.filter(row => ["FALTANTE","SIN INVENTARIO"].includes(String(row.inventory_status || "").toUpperCase()));
      const tireRows = (portal.tire_kpi && Array.isArray(portal.tire_kpi.rows)) ? portal.tire_kpi.rows : [];
      const tireCritical = tireRows.filter(row => ["CRITICA","PROXIMA"].includes(String(row.control_status || "").toUpperCase()));
      const captures = Array.isArray(portal.captures) ? portal.captures : [];
      const noDisponible = captures.filter(row => String(row.status || "").toUpperCase().includes("NO DISP")).slice(0, 50);
      const oilReport = oilRowsForRange($("kpiStart").value || (portal.period || {}).start || toIsoDate(new Date()), $("kpiEnd").value || (portal.period || {}).end || toIsoDate(new Date()));
      const oilLiters = Number(oilReport.totals.total_liters || 0);
      return {
        cards: [
          {label:"Preventivos vencidos", value:overdue.length, note:"PM vencido/urgente", tab:"preventivos", tone:overdue.length ? "bad" : "ok"},
          {label:"Servicios abiertos", value:servicesOpen.length, note:"Pendientes de cierre", tab:"ejecucionPreventivos", tone:servicesOpen.length ? "warn" : "ok"},
          {label:"OT abiertas", value:openOrders.length, note:"Ordenes en seguimiento", tab:"ordenesTrabajo", tone:openOrders.length ? "warn" : "ok"},
          {label:"Refacciones faltantes", value:spareShort.length, note:"Faltante o sin inventario", tab:"refacciones", tone:spareShort.length ? "bad" : "ok"},
          {label:"Llantas criticas", value:tireCritical.length, note:"Critica/proxima", tab:"llantasTrack", tone:tireCritical.length ? "warn" : "ok"},
          {label:"No disponibles", value:noDisponible.length, note:"Capturas recientes", tab:"captura", tone:noDisponible.length ? "bad" : "ok"},
        ],
        alerts: [
          ...overdue.slice(0,3).map(row => ({tone:"bad", label:"PM", text:`${row.equipment_code || ""} ${row.service_interval || row.component || ""}: ${row.status || ""}`})),
          ...servicesOpen.slice(0,2).map(row => ({tone:"warn", label:"Servicio", text:`${row.equipment_code || ""} ${row.service_type || ""}: ${row.status || "ABIERTO"}`})),
          ...openOrders.slice(0,3).map(row => ({tone:row.priority === "URGENTE" || row.priority === "ALTA" ? "bad" : "warn", label:"OT", text:`${row.folio || ""} ${row.equipment_code || ""}: ${row.description || ""}`})),
          ...spareShort.slice(0,2).map(row => ({tone:"bad", label:"Stock", text:`${row.equipment_code || ""} ${row.description || row.part_number || ""}: ${row.inventory_status || ""}`})),
          ...tireCritical.slice(0,2).map(row => ({tone:"warn", label:"Llanta", text:`${row.equipment_code || ""} ${row.tire_code || ""}: ${row.control_status || ""}`})),
        ],
      };
    }
    function renderExecutiveBoard(){
      const report = executiveAlerts();
      $("execCards").innerHTML = report.cards.map(card => `<div class="exec-card ${card.tone === "bad" ? "bad" : (card.tone === "warn" ? "warn" : "")}" data-exec-tab="${esc(card.tab)}" data-exec-group="${esc(card.kpiGroup || "")}"><strong>${esc(card.value)}</strong><span>${esc(card.label)}</span><small>${esc(card.note)}</small></div>`).join("");
      $("execAlerts").innerHTML = report.alerts.length
        ? report.alerts.map(alert => `<div class="exec-alert ${alert.tone === "bad" ? "bad" : "warn"}"><b>${esc(alert.label)}</b><span>${esc(alert.text)}</span></div>`).join("")
        : `<div class="exec-alert"><b>OK</b><span>Sin alertas criticas principales en el periodo actual.</span></div>`;
      $("execUpdated").textContent = portal.updated_at || portal.generated_at ? `Actualizado ${portal.updated_at || portal.generated_at}` : "Alertas automaticas";
      document.querySelectorAll("[data-exec-tab]").forEach(card => card.addEventListener("click", () => {
        if(card.dataset.execGroup && $("kpiGroup")){
          $("kpiGroup").value = card.dataset.execGroup;
          renderDashboard();
        }
        activateTab(card.dataset.execTab);
      }));
    }
    function kpiRiskRows(report){
      const rows = Array.isArray(report?.rows) ? report.rows : [];
      const preventives = preventiveRowsWithWebClosures(portal.preventives || []);
      const spareRows = (portal.parts_manuals && Array.isArray(portal.parts_manuals.rows)) ? portal.parts_manuals.rows : [];
      const tireRows = (portal.tire_kpi && Array.isArray(portal.tire_kpi.rows)) ? portal.tire_kpi.rows : [];
      const orderRows = workOrderRows().filter(row => !workOrderClosed(row));
      const byCode = (list, code) => list.filter(row => normalizedText(row.equipment_code || row.equipment || row.code) === normalizedText(code));
      const metaAvailability = Number((portal.settings || {}).meta_availability || 85);
      const metaUtilization = Number((portal.settings || {}).meta_utilization || 75);
      return rows.map(row => {
        const code = row.code || "";
        const ots = byCode(orderRows, code);
        const pm = byCode(preventives, code).filter(item => ["VENCIDO","URGENTE"].includes(String(item.status || "").toUpperCase()));
        const stock = byCode(spareRows, code).filter(item => ["FALTANTE","SIN INVENTARIO"].includes(String(item.inventory_status || "").toUpperCase()));
        const tires = byCode(tireRows, code).filter(item => ["CRITICA","PROXIMA"].includes(String(item.control_status || "").toUpperCase()));
        const noCapture = !Number(row.worked || 0) && !Number(row.mp || 0) && !Number(row.mc || 0) && !Number(row.stops || 0);
        let score = 0;
        const reasons = [];
        if(row.out || noCapture){ score += 28; reasons.push(row.out ? "Fuera/no disponible" : "Sin captura"); }
        if(Number(row.availability || 0) && Number(row.availability || 0) < metaAvailability){ score += Math.min((metaAvailability - Number(row.availability || 0)) * 1.2, 30); reasons.push("Disponibilidad baja"); }
        if(Number(row.utilization || 0) && Number(row.utilization || 0) < metaUtilization){ score += Math.min((metaUtilization - Number(row.utilization || 0)) * .7, 18); reasons.push("Utilizacion baja"); }
        if(Number(row.mc || 0) > 0){ score += Math.min(Number(row.mc || 0) * 2.5, 22); reasons.push(`${one(row.mc)} h MC`); }
        if(Number(row.stops || 0) > 0){ score += Math.min(Number(row.stops || 0) * 8, 24); reasons.push(`${num(row.stops)} parada(s)`); }
        if(ots.length){ score += Math.min(ots.length * 12 + ots.filter(item => ["URGENTE","ALTA"].includes(String(item.priority || "").toUpperCase())).length * 10, 34); reasons.push(`${ots.length} OT abierta(s)`); }
        if(pm.length){ score += Math.min(pm.length * 18, 36); reasons.push(`${pm.length} PM critico(s)`); }
        if(stock.length){ score += Math.min(stock.length * 6, 20); reasons.push(`${stock.length} stock alerta`); }
        if(tires.length){ score += Math.min(tires.length * 8, 24); reasons.push(`${tires.length} llanta(s)`); }
        score = Math.round(Math.max(score, 0));
        const color = noCapture ? "gray" : score >= 70 ? "red" : score >= 35 ? "yellow" : "green";
        return {
          ...row,
          score,
          semaphore: color,
          reasons: reasons.length ? reasons.slice(0,4).join(" | ") : "Sin alerta principal",
          open_orders: ots.length,
          pm_due: pm.length,
          stock_alert: stock.length,
          tire_alert: tires.length,
        };
      }).sort((a,b) => Number(b.score || 0) - Number(a.score || 0) || String(a.code || "").localeCompare(String(b.code || "")));
    }
    function renderKpiCommandCenter(report){
      const rows = kpiRiskRows(report);
      $("kpiCommandUpdated").textContent = `${rows.length} equipo(s) evaluado(s)`;
      const top = rows.slice(0,10);
      $("kpiPriorityList").innerHTML = top.length ? top.map((row, idx) => {
        const tone = row.semaphore === "red" ? "bad" : (row.semaphore === "yellow" ? "warn" : "");
        return `<div class="priority-row ${tone}" data-kpi-risk-eq="${esc(row.code)}"><div class="priority-rank">${idx + 1}</div><div><strong>${esc(row.code)} - ${esc(row.description || "")}</strong><span>${esc(row.reasons)}</span></div><div class="priority-score">${row.score}</div></div>`;
      }).join("") : `<div class="exec-alert"><b>OK</b><span>Sin equipos para priorizar.</span></div>`;
      $("kpiSemaphoreTable").innerHTML = `<thead><tr><th></th><th>Equipo</th><th>Disp.</th><th>Util.</th><th>MC</th><th>Paradas</th><th>OT</th><th>PM</th><th>Stock</th><th>Llantas</th><th>Riesgo</th></tr></thead><tbody>` +
        rows.map(row => `<tr data-kpi-risk-eq="${esc(row.code)}" style="cursor:pointer"><td><span class="semaphore-dot sem-${esc(row.semaphore)}"></span></td><td>${esc(row.code)} ${esc(row.description || "")}</td><td>${esc(row.availabilityText || pct(row.availability || 0))}</td><td>${esc(row.utilizationText || pct(row.utilization || 0))}</td><td>${one(row.mc || 0)}</td><td>${num(row.stops || 0)}</td><td>${row.open_orders}</td><td>${row.pm_due}</td><td>${row.stock_alert}</td><td>${row.tire_alert}</td><td><b>${row.score}</b></td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-kpi-risk-eq]").forEach(el => el.addEventListener("click", () => {
        const code = el.dataset.kpiRiskEq || "";
        if($("fichaEquipment")) $("fichaEquipment").value = code;
        renderEquipmentProfile();
        activateTab("fichaEquipo");
      }));
    }
    function renderKpiMainSummary(report, settings){
      const totals = report?.totals || {};
      const metaAvailability = Number(settings?.meta_availability || 85);
      const metaUtilization = Number(settings?.meta_utilization || 75);
      const metaReliability = Number(settings?.meta_reliability || 80);
      const metaTmef = Number(settings?.meta_tmef || 8);
      const metaTmpr = Number(settings?.meta_tmpr || 4);
      const riskRows = kpiRiskRows(report);
      const critical = riskRows.filter(row => row.semaphore === "red").length;
      const warn = riskRows.filter(row => row.semaphore === "yellow").length;
      const cards = [
        {label:"Disponibilidad", value:pct(totals.availability || 0), note:`Meta ${pct(metaAvailability)}`, width:totals.availability || 0, bad:Number(totals.availability || 0) < metaAvailability},
        {label:"Utilizacion", value:pct(totals.utilization || 0), note:`Meta ${pct(metaUtilization)}`, width:totals.utilization || 0, bad:Number(totals.utilization || 0) < metaUtilization},
        {label:"Confiabilidad", value:pct(totals.reliability || 0), note:`Meta ${pct(metaReliability)}`, width:totals.reliability || 0, bad:Number(totals.reliability || 0) < metaReliability},
        {label:"TMEF", value:`${one(totals.tmef || 0)} h`, note:`Meta ${one(metaTmef)} h`, width:Math.min((Number(totals.tmef || 0) / Math.max(metaTmef, 1)) * 100, 100), bad:Number(totals.tmef || 0) < metaTmef},
        {label:"TMPR", value:`${one(totals.tmpr || 0)} h`, note:`Meta ${one(metaTmpr)} h`, width:Math.min((metaTmpr / Math.max(Number(totals.tmpr || 0), .1)) * 100, 100), bad:Number(totals.tmpr || 0) > metaTmpr},
        {label:"Equipos criticos", value:String(critical), note:`${warn} en atencion`, width:Math.max(100 - critical * 18 - warn * 8, 0), bad:critical > 0, warn:critical === 0 && warn > 0},
      ];
      $("kpiMainSummaryNote").textContent = `${report.group || "Dashboard"} | ${report.start || ""} a ${report.end || ""}`;
      $("kpiMainStrip").innerHTML = cards.map(card => {
        const tone = card.bad ? "bad" : (card.warn ? "warn" : "");
        return `<div class="kpi-main-card ${tone}"><span>${esc(card.label)}</span><strong>${esc(card.value)}</strong><small>${esc(card.note)}</small><div class="kpi-main-meter"><i style="width:${Math.max(Math.min(Number(card.width || 0), 100), 0)}%"></i></div></div>`;
      }).join("");
    }
    function workOrderRows(){
      const payload = portal.work_orders || {};
      return Array.isArray(payload.records) ? payload.records : [];
    }
    function workOrderClosed(row){
      return ["CERRADA","CERRADO","CANCELADA","CANCELADO"].includes(String(row.status || "").toUpperCase());
    }
    function workOrderClass(row){
      const status = String(row.status || "").toUpperCase();
      if(status.includes("CERR")) return "ok";
      if(status.includes("CANCEL")) return "bad";
      return row.priority === "URGENTE" || row.priority === "ALTA" ? "bad" : "warn";
    }
    function selectedFichaCode(){
      return $("fichaEquipment")?.value || portalEquipment()[0]?.code || portalEquipment()[0]?.equipment_code || "";
    }
    function rowEquipmentCode(row){
      return String(row?.equipment_code || row?.equipment || row?.code || "").trim();
    }
    function rowsForEquipment(rows, code){
      const wanted = normalizedText(code);
      return (Array.isArray(rows) ? rows : []).filter(row => normalizedText(rowEquipmentCode(row)) === wanted);
    }
    function renderEquipmentProfile(){
      const code = selectedFichaCode();
      const start = $("fichaStart").value || $("kpiStart").value || (portal.period || {}).start || toIsoDate(new Date());
      const end = $("fichaEnd").value || $("kpiEnd").value || (portal.period || {}).end || start;
      if(!code){
        $("fichaHeader").innerHTML = `<h3>Ficha equipo</h3><p class="muted">No hay equipos cargados.</p>`;
        ["fichaMetrics","fichaAlerts","fichaServicesTable","fichaPreventivesTable","fichaCapturesTable","fichaPartsTable"].forEach(id => $(id).innerHTML = "");
        return;
      }
      const eq = portalEquipment().find(item => normalizedText(item.code || item.equipment_code) === normalizedText(code)) || {};
      const captures = rowsForEquipment(portal.captures || [], code).filter(row => !row.work_date || inRange(row.work_date, start, end)).sort((a,b) => String(b.work_date || "").localeCompare(String(a.work_date || "")));
      const services = rowsForEquipment(portal.service_history || [], code).filter(row => !row.completed_date || inRange(row.completed_date, start, end)).sort((a,b) => String(b.completed_date || "").localeCompare(String(a.completed_date || "")));
      const preventives = preventiveRowsWithWebClosures(portal.preventives || []).filter(row => normalizedText(row.equipment_code) === normalizedText(code)).sort((a,b) => Number(a.hours_remaining || 999999) - Number(b.hours_remaining || 999999));
      const spareRows = rowsForEquipment((portal.parts_manuals || {}).rows || [], code);
      const tireRows = rowsForEquipment((portal.tire_kpi || {}).rows || [], code);
      const orders = rowsForEquipment(workOrderRows(), code).sort((a,b) => String(b.date || "").localeCompare(String(a.date || "")));
      const openOrders = orders.filter(row => !workOrderClosed(row));
      const oilReport = oilRowsForRange(start, end);
      const oilRow = (oilReport.rows || []).find(row => normalizedText(row.code) === normalizedText(code)) || {};
      const mp = captures.reduce((sum,row) => sum + Number(row.mp_hours || 0), 0);
      const mc = captures.reduce((sum,row) => sum + Number(row.mc_hours || 0), 0);
      const worked = captures.reduce((sum,row) => sum + Number(row.worked_hours || 0), 0);
      const stops = captures.reduce((sum,row) => sum + Number(row.stops || 0), 0);
      const noDisp = captures.filter(row => unavailable(row.status)).length;
      const overdue = preventives.filter(row => ["VENCIDO","URGENTE"].includes(String(row.status || "").toUpperCase()));
      const spareShort = spareRows.filter(row => ["FALTANTE","SIN INVENTARIO"].includes(String(row.inventory_status || "").toUpperCase()));
      const tireCritical = tireRows.filter(row => ["CRITICA","PROXIMA"].includes(String(row.control_status || "").toUpperCase()));
      const lastCapture = captures[0] || {};
      $("fichaHeader").innerHTML = `
        <h3>${esc(code)}</h3>
        <p><b>${esc(eq.description || eq.family || "Equipo sin descripcion")}</b></p>
        <p>Periodo: ${esc(start)} a ${esc(end)}</p>
        <p>Ultima captura: ${esc(lastCapture.work_date || "S/D")} ${esc(lastCapture.status || "")}</p>
        <div class="profile-badges">
          <span class="profile-badge ${noDisp ? "bad" : ""}">${noDisp ? "Revision disponibilidad" : "Operacion OK"}</span>
          <span class="profile-badge ${overdue.length ? "bad" : ""}">${overdue.length} PM critico(s)</span>
          <span class="profile-badge ${openOrders.length ? "warn" : ""}">${openOrders.length} OT abierta(s)</span>
          <span class="profile-badge ${spareShort.length ? "warn" : ""}">${spareShort.length} refaccion(es) alerta</span>
        </div>
        <div class="profile-actions">
          <button class="btn secondary" type="button" data-profile-tab="captura">Captura diaria</button>
          <button class="btn" type="button" data-profile-ot="${esc(code)}">Generar OT</button>
          <button class="btn secondary" type="button" data-profile-tab="ejecucionPreventivos">Servicio preventivo</button>
          <button class="btn secondary" type="button" data-profile-tab="refacciones">Refacciones</button>
          <button class="btn secondary" type="button" data-profile-tab="llantasTrack">Llantas</button>
        </div>`;
      $("fichaMetrics").innerHTML = [
        ["Hrs trabajadas", one(worked)],
        ["Hrs MP", one(mp)],
        ["Hrs MC", one(mc)],
        ["Paradas", stops],
        ["Aceites L", one(oilRow.total_liters || 0)],
        ["OT abiertas", openOrders.length],
      ].map(([label,value]) => `<div class="profile-metric"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>`).join("");
      const alerts = [
        ...openOrders.slice(0,4).map(row => ({tone:row.priority === "URGENTE" || row.priority === "ALTA" ? "bad" : "warn", label:"OT", text:`${row.folio || ""} ${row.priority || ""}: ${row.description || ""}`})),
        ...overdue.slice(0,4).map(row => ({tone:"bad", label:"PM", text:`${row.service_interval || ""} ${row.component || ""}: ${row.status || ""}, faltan ${one(row.hours_remaining || 0)} h`})),
        ...spareShort.slice(0,4).map(row => ({tone:"warn", label:"Stock", text:`${row.description || row.part_number || ""}: ${row.inventory_status || ""}`})),
        ...tireCritical.slice(0,4).map(row => ({tone:"warn", label:"Llanta", text:`${row.tire_code || ""} pos. ${row.position || ""}: ${row.control_status || ""}`})),
        ...captures.filter(row => unavailable(row.status)).slice(0,3).map(row => ({tone:"bad", label:"Disp.", text:`${row.work_date || ""}: ${row.status || ""} ${row.fault || ""}`})),
      ];
      $("fichaAlertCount").textContent = `${alerts.length} alerta(s)`;
      $("fichaAlerts").innerHTML = alerts.length ? alerts.map(alert => `<div class="exec-alert ${alert.tone === "bad" ? "bad" : "warn"}"><b>${esc(alert.label)}</b><span>${esc(alert.text)}</span></div>`).join("") : `<div class="exec-alert"><b>OK</b><span>Sin alertas principales para este equipo.</span></div>`;
      $("fichaServiceCount").textContent = `${services.length} servicio(s)`;
      $("fichaServicesTable").innerHTML = `<thead><tr><th>Fecha</th><th>Servicio</th><th>Componente</th><th>Horometro</th><th>Estado</th><th>Detalle</th></tr></thead><tbody>` +
        (services.slice(0,30).map(row => `<tr><td>${esc(row.completed_date || "")}</td><td>${esc([row.service_name,row.service_interval].filter(Boolean).join(" / "))}</td><td>${esc(row.component || "")}</td><td>${one(row.completed_meter || 0)}</td><td>${esc(row.status || "")}</td><td>${esc(shortText(row.notes || serviceOilsText(row) || serviceFiltersText(row), 120))}</td></tr>`).join("") || `<tr><td colspan="6">Sin servicios en el periodo.</td></tr>`) + `</tbody>`;
      $("fichaPreventiveCount").textContent = `${preventives.length} preventivo(s)`;
      $("fichaPreventivesTable").innerHTML = `<thead><tr><th>Servicio</th><th>Componente</th><th>Ultimo</th><th>Proximo</th><th>Hrs rest.</th><th>Estado</th></tr></thead><tbody>` +
        (preventives.slice(0,30).map(row => `<tr><td>${esc(row.service_interval || row.service_name || "")}</td><td>${esc(row.component || "")}</td><td>${one(row.last_service_meter || 0)}</td><td>${one(row.next_service_meter || 0)}</td><td>${one(row.hours_remaining || 0)}</td><td><span class="pill ${row.status === "PROGRAMADO" ? "ok" : (row.status === "PROXIMO" ? "warn" : "bad")}">${esc(row.status || "")}</span></td></tr>`).join("") || `<tr><td colspan="6">Sin preventivos programados.</td></tr>`) + `</tbody>`;
      $("fichaCaptureCount").textContent = `${captures.length} captura(s)`;
      $("fichaCapturesTable").innerHTML = `<thead><tr><th>Fecha</th><th>Turno</th><th>Comp.</th><th>HI</th><th>HF</th><th>Trab.</th><th>MP</th><th>MC</th><th>Estatus</th><th>Falla</th></tr></thead><tbody>` +
        (captures.slice(0,40).map(row => `<tr><td>${esc(row.work_date || "")}</td><td>${esc(row.shift || "")}</td><td>${esc(row.component || row.component_name || "")}</td><td>${one(row.hi || 0)}</td><td>${one(row.hf || 0)}</td><td>${one(row.worked_hours || 0)}</td><td>${one(row.mp_hours || 0)}</td><td>${one(row.mc_hours || 0)}</td><td>${esc(row.status || "")}</td><td>${esc(shortText(row.fault || row.observations || "", 80))}</td></tr>`).join("") || `<tr><td colspan="10">Sin capturas en el periodo.</td></tr>`) + `</tbody>`;
      const partRows = [
        ...spareRows.slice(0,25).map(row => ({kind:"Refaccion", code:row.part_number || row.equivalent_part || "", desc:row.description || "", status:row.inventory_status || "", extra:`Req. ${one(row.quantity || 0)} / Disp. ${row.available == null ? "S/D" : one(row.available)}`})),
        ...tireRows.slice(0,20).map(row => ({kind:"Llanta", code:row.tire_code || "", desc:`Pos. ${row.position || ""} ${row.brand || ""}`, status:row.control_status || row.status || "", extra:`Vida ${one(row.remaining_hours || 0)} h`})),
      ];
      $("fichaPartsCount").textContent = `${partRows.length} registro(s)`;
      $("fichaPartsTable").innerHTML = `<thead><tr><th>Tipo</th><th>Codigo</th><th>Descripcion</th><th>Estado</th><th>Detalle</th></tr></thead><tbody>` +
        (partRows.map(row => `<tr><td>${esc(row.kind)}</td><td>${esc(row.code)}</td><td>${esc(shortText(row.desc, 100))}</td><td>${esc(row.status)}</td><td>${esc(row.extra)}</td></tr>`).join("") || `<tr><td colspan="5">Sin refacciones o llantas relacionadas.</td></tr>`) + `</tbody>`;
      document.querySelectorAll("[data-profile-tab]").forEach(button => button.addEventListener("click", () => activateTab(button.dataset.profileTab)));
      document.querySelectorAll("[data-profile-ot]").forEach(button => button.addEventListener("click", () => {
        newWorkOrder({
          equipment_code: button.dataset.profileOt || code,
          equipment_description: eq.description || eq.family || "",
          origin: "FICHA EQUIPO",
          priority: overdue.length || noDisp ? "ALTA" : "MEDIA",
          description: alerts[0]?.text || `Seguimiento de mantenimiento para ${code}`,
        });
      }));
    }
    function newWorkOrder(prefill={}){
      currentWorkOrderRecord = null;
      $("woId").value = prefill.id || "";
      $("woFolio").value = prefill.folio || "";
      $("woDate").value = prefill.date || toIsoDate(new Date());
      $("woEquipment").value = prefill.equipment_code || selectedFichaCode() || "";
      $("woOrigin").value = prefill.origin || "MANUAL";
      $("woPriority").value = prefill.priority || "MEDIA";
      $("woState").value = prefill.status || "ABIERTA";
      $("woResponsible").value = prefill.responsible || "";
      $("woMechanic").value = prefill.mechanic || "";
      $("woSupervisor").value = prefill.supervisor || "";
      $("woDescription").value = prefill.description || "";
      $("woAction").value = prefill.action || "";
      $("woParts").value = prefill.parts_used || "";
      $("woLubricants").value = prefill.lubricants_used || "";
      $("woEvidence").value = prefill.evidence_note || "";
      $("woStatus").textContent = prefill.folio ? `Editando ${prefill.folio}` : "Nueva OT";
      activateTab("ordenesTrabajo");
    }
    function workOrderPayload(close=false){
      const code = $("woEquipment").value || "";
      const eq = portalEquipment().find(item => normalizedText(item.code || item.equipment_code) === normalizedText(code)) || {};
      return {
        id: $("woId").value || "",
        folio: $("woFolio").value || "",
        date: $("woDate").value || toIsoDate(new Date()),
        equipment_code: code,
        equipment_description: eq.description || eq.family || "",
        origin: $("woOrigin").value,
        priority: $("woPriority").value,
        status: close ? "CERRADA" : $("woState").value,
        responsible: $("woResponsible").value.trim(),
        mechanic: $("woMechanic").value.trim(),
        supervisor: $("woSupervisor").value.trim(),
        description: $("woDescription").value.trim(),
        action: $("woAction").value.trim(),
        parts_used: $("woParts").value.trim(),
        lubricants_used: $("woLubricants").value.trim(),
        evidence_note: $("woEvidence").value.trim(),
      };
    }
    function fillWorkOrder(row){
      currentWorkOrderRecord = row;
      newWorkOrder(row);
    }
    function filteredWorkOrders(){
      const selected = $("woFilterEquipment").value || "";
      const status = $("woFilterStatus").value || "";
      const priority = $("woFilterPriority").value || "";
      const search = normalizedText($("woSearch").value || "");
      return workOrderRows().filter(row => {
        const eqOk = !selected || normalizedText(row.equipment_code) === normalizedText(selected);
        const statusOk = !status || String(row.status || "").toUpperCase() === status;
        const priorityOk = !priority || String(row.priority || "").toUpperCase() === priority;
        const text = normalizedText([row.folio,row.equipment_code,row.origin,row.priority,row.status,row.responsible,row.description,row.action].join(" "));
        return eqOk && statusOk && priorityOk && (!search || text.includes(search));
      }).sort((a,b) => {
        const rank = {URGENTE:0, ALTA:1, MEDIA:2, BAJA:3};
        return (workOrderClosed(a) ? 1 : 0) - (workOrderClosed(b) ? 1 : 0)
          || (rank[a.priority] ?? 9) - (rank[b.priority] ?? 9)
          || String(b.date || "").localeCompare(String(a.date || ""));
      });
    }
    function renderWorkOrders(){
      const rows = filteredWorkOrders();
      const all = workOrderRows();
      const open = all.filter(row => !workOrderClosed(row)).length;
      const process = all.filter(row => String(row.status || "").toUpperCase().includes("PROCESO")).length;
      const urgent = all.filter(row => !workOrderClosed(row) && ["URGENTE","ALTA"].includes(String(row.priority || "").toUpperCase())).length;
      const closed = all.filter(workOrderClosed).length;
      $("woSummaryText").textContent = `${rows.length} mostrada(s)`;
      $("woSummaryCards").innerHTML = [
        ["Abiertas", open, open ? "warn" : ""],
        ["En proceso", process, process ? "warn" : ""],
        ["Urgentes/altas", urgent, urgent ? "bad" : ""],
        ["Cerradas/cancel.", closed, ""],
      ].map(([label,value,tone]) => `<div class="exec-card ${tone}"><strong>${value}</strong><span>${esc(label)}</span><small>Ordenes de trabajo</small></div>`).join("");
      $("woTable").innerHTML = `<thead><tr><th>Folio</th><th>Fecha</th><th>Equipo</th><th>Origen</th><th>Prioridad</th><th>Estatus</th><th>Responsable</th><th>Descripcion</th><th>Accion</th></tr></thead><tbody>` +
        (rows.map(row => `<tr data-wo-id="${esc(row.id || row.folio || "")}" style="cursor:pointer"><td>${esc(row.folio || "")}</td><td>${esc(row.date || "")}</td><td>${esc(row.equipment_code || "")}</td><td>${esc(row.origin || "")}</td><td><span class="pill ${workOrderClass(row)}">${esc(row.priority || "")}</span></td><td><span class="pill ${workOrderClass(row)}">${esc(row.status || "")}</span></td><td>${esc(row.responsible || row.mechanic || "")}</td><td>${esc(shortText(row.description || "", 120))}</td><td>${esc(shortText(row.action || "", 100))}</td></tr>`).join("") || `<tr><td colspan="9">Sin ordenes de trabajo.</td></tr>`) + `</tbody>`;
      document.querySelectorAll("[data-wo-id]").forEach(row => row.addEventListener("click", () => {
        const record = workOrderRows().find(item => String(item.id || item.folio || "") === String(row.dataset.woId || ""));
        if(record) fillWorkOrder(record);
      }));
    }
    async function saveWorkOrder(close=false){
      if(!hasApiKey(true)) return;
      const payload = workOrderPayload(close);
      if(!payload.equipment_code) return alert("Selecciona un equipo.");
      if(!payload.description) return alert("Describe el trabajo de la OT.");
      if(close && !payload.action && !confirm("No capturaste accion/cierre. ¿Cerrar OT de todos modos?")) return;
      const response = await fetch("/api/work-orders/records", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      if(result.portal) portal = result.portal;
      renderPortalSelectors();
      renderWorkOrders();
      renderExecutiveBoard();
      renderEquipmentProfile();
      renderBacklog();
      if(result.record) fillWorkOrder(result.record);
      $("woStatus").textContent = close ? "OT cerrada." : "OT guardada.";
    }
    async function deleteWorkOrder(){
      if(!hasApiKey(true)) return;
      const id = $("woId").value || $("woFolio").value || currentWorkOrderRecord?.id || currentWorkOrderRecord?.folio || "";
      if(!id) return alert("Selecciona una OT.");
      if(!confirm("¿Eliminar esta orden de trabajo?")) return;
      const response = await fetch("/api/work-orders/records/delete", {method:"POST", headers:headers(true), body:JSON.stringify({id})});
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      if(result.portal) portal = result.portal;
      newWorkOrder();
      renderPortalSelectors();
      renderWorkOrders();
      renderExecutiveBoard();
      renderEquipmentProfile();
      renderBacklog();
      $("woStatus").textContent = "OT eliminada.";
    }
    function renderSelectors(){
      const current = $("equipmentSelect").value;
      $("equipmentSelect").innerHTML = (data.equipment || []).map(e => `<option value="${esc(e.code)}">${esc(e.code)} - ${esc(e.description || e.family || "")}</option>`).join("");
      if(current) $("equipmentSelect").value = current;
      renderServiceOptions();
    }
    function selectedEquipment(){ return (data.equipment || []).find(e => e.code === $("equipmentSelect").value); }
    function renderServiceOptions(){
      const eq = selectedEquipment();
      const previous = $("serviceSelect").value;
      const services = [...new Set(((eq && eq.filters) || []).map(f => f.service_interval || "").filter(Boolean))];
      $("serviceSelect").innerHTML = `<option value="">Todos</option>` + services.map(s => `<option>${esc(s)}</option>`).join("");
      $("serviceSelect").value = services.includes(previous) ? previous : "";
    }
    function renderFilters(){
      const eq = selectedEquipment();
      const service = $("serviceSelect").value;
      const status = $("statusSelect").value;
      const search = ($("filterSearch").value || "").toUpperCase();
      let rows = ((eq && eq.filters) || []).filter(f => !service || f.service_interval === service);
      rows = rows.filter(f => !status || f.inventory_status === status);
      rows = rows.filter(f => !search || [f.part_number,f.donaldson_part,f.description,f.service_interval,f.item_type].join(" ").toUpperCase().includes(search));
      $("filtersTable").innerHTML = `<thead><tr><th>Equipo</th><th>Servicio</th><th>Tipo</th><th>No. parte</th><th>Donaldson</th><th>Descripcion</th><th>Req.</th><th>Exist.</th><th>Faltante</th><th>Estado</th></tr></thead><tbody>` +
        rows.map(f => `<tr><td>${esc(eq?.code || "")}</td><td>${esc(f.service_interval)}</td><td>${esc(f.item_type)}</td><td>${esc(f.part_number)}</td><td>${esc(f.donaldson_part)}</td><td>${esc(f.description)}</td><td>${num(f.quantity)} ${esc(f.unit||"PZA")}</td><td>${f.available == null ? "" : num(f.available)}</td><td>${Number(f.shortage||0)>0 ? num(f.shortage) : ""}</td><td><span class="pill ${statusClass(f.inventory_status)}">${esc(f.inventory_status)}</span></td></tr>`).join("") +
        `</tbody>`;
    }
    function renderInventory(){
      const search = ($("inventorySearch").value || "").toUpperCase();
      const rows = (data.inventory || []).filter(i => !search || [i.equipment,i.item_type,i.part_number,i.description,i.location].join(" ").toUpperCase().includes(search));
      $("inventoryTable").innerHTML = `<thead><tr><th>Equipo</th><th>Tipo</th><th>No. parte</th><th>Descripcion</th><th>Exist.</th><th>Unidad</th><th>Min.</th><th>Ubicacion</th><th>Actualizado</th></tr></thead><tbody>` +
        rows.map(i => `<tr><td>${esc(i.equipment || i.equipment_codes)}</td><td>${esc(i.item_type)}</td><td>${esc(i.part_number)}</td><td>${esc(i.description)}</td><td>${num(i.quantity)}</td><td>${esc(i.unit||"PZA")}</td><td>${num(i.min_stock)}</td><td>${esc(i.location)}</td><td>${esc(i.updated_at)}</td></tr>`).join("") +
        `</tbody>`;
    }
    function renderMovements(){
      const rows = data.movements || [];
      $("movementTable").innerHTML = `<thead><tr><th>Fecha</th><th>Parte</th><th>Tipo</th><th>Cant.</th><th>Saldo</th><th>Ref.</th></tr></thead><tbody>` +
        rows.map(m => `<tr><td>${esc(m.movement_date)}</td><td>${esc(m.part_number)}</td><td>${esc(m.movement_type)}</td><td>${num(m.quantity)}</td><td>${num(m.balance_after)}</td><td>${esc(m.reference)}</td></tr>`).join("") +
        `</tbody>`;
    }
    function eppSelectedCode(){ return ($("eppCode").value || "").trim().toUpperCase(); }
    function eppStatusClass(status){
      if(status === "OK") return "ok";
      if(status === "SIN STOCK" || status === "VENCIDO POR VIDA UTIL") return "bad";
      return "warn";
    }
    function eppWorkerLabel(row){
      return row && row.employee_id ? `${row.employee_id} - ${row.worker_name}` : (row?.worker_name || "");
    }
    function findEppWorker(value){
      const text = String(value || "").toUpperCase();
      return (epp.workers || []).find(row => [eppWorkerLabel(row), row.worker_name, row.employee_id].some(v => String(v || "").toUpperCase() === text));
    }
    function fillWorkerForm(row){
      currentEppWorkerId = row?.id || null;
      $("eppWorkerEmployee").value = row?.employee_id || "";
      $("eppWorkerName").value = row?.worker_name || "";
      $("eppWorkerArea").value = row?.area || "";
      $("eppWorkerPosition").value = row?.position || "";
    }
    function applyWorkerToEppForms(row){
      if(!row) return;
      $("eppDelWorker").value = row.worker_name || "";
      $("eppDelEmployee").value = row.employee_id || "";
      $("eppDelArea").value = row.area || row.position || "";
      if(!$("eppDelReceived").value) $("eppDelReceived").value = row.worker_name || "";
      $("eppMovWorker").value = row.worker_name || "";
      $("eppMovEmployee").value = row.employee_id || "";
      $("eppMovArea").value = row.area || row.position || "";
    }
    function renderEppWorkers(){
      $("eppWorkerList").innerHTML = (epp.workers || []).filter(row => row.active !== 0).map(row => `<option value="${esc(eppWorkerLabel(row))}"></option>`).join("");
      $("eppWorkersTable").innerHTML = `<thead><tr><th>No.</th><th>Trabajador</th><th>Area</th><th>Puesto</th></tr></thead><tbody>` +
        (epp.workers || []).map(row => `<tr data-epp-worker="${row.id}"><td>${esc(row.employee_id)}</td><td>${esc(row.worker_name)}</td><td>${esc(row.area)}</td><td>${esc(row.position)}</td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-epp-worker]").forEach(tr => tr.addEventListener("click", () => {
        const row = (epp.workers || []).find(item => String(item.id) === String(tr.dataset.eppWorker));
        if(row) fillWorkerForm(row);
      }));
    }
    function clearEppForm(){
      ["eppCode","eppCategory","eppSize","eppRisk","eppLocation","eppNotes","eppDesc"].forEach(id => $(id).value = "");
      $("eppUnit").value = "PZA"; $("eppQty").value = "0"; $("eppMin").value = "0"; $("eppLife").value = "0"; $("eppTraining").value = "0";
      $("eppMovDate").value = toIsoDate(new Date()); $("eppDelDate").value = toIsoDate(new Date());
      $("eppMovQty").value = "1"; $("eppDelQty").value = "1";
    }
    function fillEppForm(row){
      $("eppCode").value = row.code || "";
      $("eppDesc").value = row.description || "";
      $("eppCategory").value = row.category || "";
      $("eppSize").value = row.size || "";
      $("eppUnit").value = row.unit || "PZA";
      $("eppQty").value = row.quantity || 0;
      $("eppMin").value = row.min_stock || 0;
      $("eppLife").value = row.useful_life_days || 0;
      $("eppRisk").value = row.risk_area || "";
      $("eppLocation").value = row.location || "";
      $("eppNotes").value = row.maintenance_notes || "";
      $("eppTraining").value = row.training_required ? "1" : "0";
    }
    function renderEpp(){
      const s = epp.summary || {};
      $("eppStats").innerHTML = [
        ["Articulos", s.items || 0],
        ["Existencia total", num(s.total_quantity || 0)],
        ["Bajo minimo", s.low_stock || 0],
        ["Sin stock", s.out_stock || 0],
        ["Prox. reposicion", s.due_soon || 0],
        ["Vencidos", s.overdue || 0],
        ["Trabajadores", s.workers || 0],
        ["Entregas", s.deliveries || 0],
      ].map(([k,v]) => `<div class="stat"><strong>${v}</strong>${k}</div>`).join("");
      renderEppWorkers();
      const search = ($("eppSearch").value || "").toUpperCase();
      const status = $("eppStatus").value;
      const rows = (epp.items || []).filter(row => (!status || row.status === status) && (!search || [row.code,row.description,row.category,row.size,row.risk_area,row.location,row.next_worker].join(" ").toUpperCase().includes(search)));
      $("eppTable").innerHTML = `<thead><tr><th>Estado</th><th>Codigo</th><th>Descripcion</th><th>Categoria</th><th>Talla</th><th>Exist.</th><th>Unidad</th><th>Min.</th><th>Vida dias</th><th>Prox. repos.</th><th>QR</th><th>Area/riesgo</th><th>Ubicacion</th></tr></thead><tbody>` +
        rows.map(row => {
          const cls = eppStatusClass(row.status);
          return `<tr data-epp-code="${esc(row.code)}"><td><span class="pill ${cls}">${esc(row.status)}</span></td><td>${esc(row.code)}</td><td>${esc(row.description)}</td><td>${esc(row.category)}</td><td>${esc(row.size)}</td><td>${num(row.quantity)}</td><td>${esc(row.unit || "PZA")}</td><td>${num(row.min_stock)}</td><td>${num(row.useful_life_days)}</td><td>${esc(row.next_due_date || "")}</td><td><a href="/api/epp/items/${encodeURIComponent(row.code)}/qr.pdf" target="_blank">QR</a></td><td>${esc(row.risk_area)}</td><td>${esc(row.location)}</td></tr>`;
        }).join("") + `</tbody>`;
      document.querySelectorAll("[data-epp-code]").forEach(tr => tr.addEventListener("click", () => {
        const row = (epp.items || []).find(item => item.code === tr.dataset.eppCode);
        if(row) fillEppForm(row);
        renderEppHistory();
      }));
      renderEppHistory();
    }
    function renderEppHistory(){
      const code = eppSelectedCode();
      const match = row => !code || String(row.item_code || row.code || "").toUpperCase() === code;
      const deliveries = (epp.deliveries || []).filter(match).slice(0, 120);
      $("eppDeliveriesTable").innerHTML = `<thead><tr><th>Fecha</th><th>Codigo</th><th>Trabajador</th><th>No.</th><th>Area</th><th>Cant.</th><th>Vence</th><th>Firma</th><th>Estado</th><th>PDF</th></tr></thead><tbody>` +
        deliveries.map(row => `<tr><td>${esc(row.delivery_date)}</td><td>${esc(row.item_code)}</td><td>${esc(row.worker_name)}</td><td>${esc(row.employee_id)}</td><td>${esc(row.area)}</td><td>${num(row.quantity)}</td><td>${esc(row.due_date)}</td><td>${esc(row.signature)}</td><td>${esc(row.condition_status)}</td><td><a href="/api/epp/deliveries/${row.id}/pdf" target="_blank">PDF</a></td></tr>`).join("") + `</tbody>`;
      const movements = (epp.movements || []).filter(match).slice(0, 120);
      $("eppMovementsTable").innerHTML = `<thead><tr><th>Fecha</th><th>Tipo</th><th>Codigo</th><th>Cant.</th><th>Saldo</th><th>Trabajador</th><th>Area</th><th>Ref.</th></tr></thead><tbody>` +
        movements.map(row => `<tr><td>${esc(row.movement_date)}</td><td>${esc(row.movement_type)}</td><td>${esc(row.item_code)}</td><td>${num(row.quantity)}</td><td>${num(row.balance_after)}</td><td>${esc(row.worker_name)}</td><td>${esc(row.area)}</td><td>${esc(row.reference)}</td></tr>`).join("") + `</tbody>`;
    }
    function eppItemPayload(){
      return {
        code:$("eppCode").value, description:$("eppDesc").value, category:$("eppCategory").value, size:$("eppSize").value,
        unit:$("eppUnit").value, quantity:$("eppQty").value, min_stock:$("eppMin").value, useful_life_days:$("eppLife").value,
        risk_area:$("eppRisk").value, location:$("eppLocation").value, training_required:$("eppTraining").value === "1",
        maintenance_notes:$("eppNotes").value,
      };
    }
    async function saveEppItem(){
      if(!hasApiKey(true)) return;
      const r = await fetch("/api/epp/items", {method:"POST", headers:headers(true), body:JSON.stringify(eppItemPayload())});
      if(!r.ok) return alert(await apiError(r));
      await refreshEpp();
    }
    async function deleteEppItem(){
      if(!hasApiKey(true)) return;
      const code = eppSelectedCode();
      if(!code) return alert("Selecciona o captura un codigo EPP.");
      if(!confirm(`Eliminar EPP ${code}?`)) return;
      const r = await fetch("/api/epp/items/delete", {method:"POST", headers:headers(true), body:JSON.stringify({code})});
      if(!r.ok) return alert(await apiError(r));
      clearEppForm(); await refreshEpp();
    }
    async function deleteAllEpp(){
      if(!hasApiKey(true)) return;
      if(prompt("Escribe ELIMINAR para borrar todo el EPP:") !== "ELIMINAR") return;
      const r = await fetch("/api/epp/delete-all", {method:"POST", headers:headers(true), body:JSON.stringify({confirm:"ELIMINAR"})});
      if(!r.ok) return alert(await apiError(r));
      clearEppForm(); await refreshEpp();
    }
    async function saveEppWorker(){
      if(!hasApiKey(true)) return;
      const payload = {id:currentEppWorkerId, employee_id:$("eppWorkerEmployee").value, worker_name:$("eppWorkerName").value, area:$("eppWorkerArea").value, position:$("eppWorkerPosition").value};
      const r = await fetch("/api/epp/workers", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!r.ok) return alert(await apiError(r));
      await refreshEpp();
    }
    async function deleteEppWorker(){
      if(!hasApiKey(true)) return;
      if(!currentEppWorkerId) return alert("Selecciona un trabajador.");
      if(!confirm("Eliminar trabajador del catalogo EPP?")) return;
      const r = await fetch("/api/epp/workers/delete", {method:"POST", headers:headers(true), body:JSON.stringify({id:currentEppWorkerId})});
      if(!r.ok) return alert(await apiError(r));
      fillWorkerForm(null); await refreshEpp();
    }
    function openEppQr(){
      const code = eppSelectedCode();
      if(!code) return alert("Selecciona o captura un codigo EPP.");
      window.open(`/api/epp/items/${encodeURIComponent(code)}/qr.pdf`, "_blank");
    }
    function openLastEppDeliveryPdf(){
      const code = eppSelectedCode();
      const rows = (epp.deliveries || []).filter(row => !code || String(row.item_code || "").toUpperCase() === code);
      if(!rows.length) return alert("No hay entrega para generar PDF.");
      window.open(`/api/epp/deliveries/${rows[0].id}/pdf`, "_blank");
    }
    function applySelectedWorker(){
      const row = currentEppWorkerId ? (epp.workers || []).find(item => String(item.id) === String(currentEppWorkerId)) : findEppWorker($("eppWorkerName").value);
      if(!row) return alert("Selecciona un trabajador.");
      applyWorkerToEppForms(row);
    }
    async function saveEppMovement(){
      if(!hasApiKey(true)) return;
      const worker = findEppWorker($("eppMovWorker").value);
      if(worker) { $("eppMovWorker").value = worker.worker_name || ""; $("eppMovEmployee").value = worker.employee_id || ""; $("eppMovArea").value = worker.area || worker.position || ""; }
      const payload = { code:eppSelectedCode(), movement_date:$("eppMovDate").value, movement_type:$("eppMovType").value, quantity:$("eppMovQty").value, worker_name:$("eppMovWorker").value, employee_id:$("eppMovEmployee").value, area:$("eppMovArea").value, reference:$("eppMovRef").value, notes:$("eppMovNotes").value };
      const r = await fetch("/api/epp/movements", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!r.ok) return alert(await apiError(r));
      $("eppMovQty").value = "1"; $("eppMovRef").value = ""; $("eppMovNotes").value = ""; await refreshEpp();
    }
    async function saveEppDelivery(){
      if(!hasApiKey(true)) return;
      const worker = findEppWorker($("eppDelWorker").value);
      if(worker) applyWorkerToEppForms(worker);
      const payload = { code:eppSelectedCode(), delivery_date:$("eppDelDate").value, worker_name:$("eppDelWorker").value, employee_id:$("eppDelEmployee").value, area:$("eppDelArea").value, quantity:$("eppDelQty").value, received_by:$("eppDelReceived").value, signature:$("eppDelSignature").value, training_done:$("eppDelTraining").value === "1", condition_status:$("eppDelCondition").value, notes:$("eppDelNotes").value };
      const r = await fetch("/api/epp/deliveries", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!r.ok) return alert(await apiError(r));
      const saved = await r.json().catch(() => ({}));
      $("eppDelQty").value = "1"; $("eppDelNotes").value = ""; await refreshEpp();
      if(saved.delivery_id) window.open(`/api/epp/deliveries/${saved.delivery_id}/pdf`, "_blank");
    }
    async function refreshEpp(){
      const r = await fetch("/api/epp", {headers: headers()});
      if(!r.ok) throw new Error(await apiError(r));
      epp = await r.json();
      renderEpp();
    }
    async function exportEpp(){
      const r = await fetch("/api/epp/export", {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      const blob = await r.blob(); const a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = "Inventario_EPP_MGA.xlsx"; a.click();
    }
    async function importEpp(){
      const file = $("eppImportFile").files[0]; if(!file) return alert("Selecciona un Excel.");
      if(!hasApiKey(true)) return;
      const dataUrl = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload=()=>res(fr.result); fr.onerror=rej; fr.readAsDataURL(file); });
      const r = await fetch("/api/epp/import", {method:"POST", headers:headers(true), body:JSON.stringify({file_name:file.name, data:String(dataUrl), replace:true})});
      const payload = await r.json().catch(() => ({}));
      $("eppImportResult").textContent = JSON.stringify(payload, null, 2);
      if(r.ok) await refreshEpp();
    }
    function parseIsoDate(value){
      const parts = String(value || "").split("-").map(Number);
      if(parts.length !== 3 || parts.some(Number.isNaN)) return new Date();
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    function toIsoDate(date){
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, "0");
      const d = String(date.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    function addDays(date, days){ const copy = new Date(date); copy.setDate(copy.getDate() + days); return copy; }
    function periodRange(period, baseValue){
      const base = parseIsoDate(baseValue || (portal.period || {}).start || toIsoDate(new Date()));
      const key = String(period || "Mes").toLowerCase();
      if(key.startsWith("sem")){
        const start = addDays(base, -((base.getDay() + 6) % 7));
        return [toIsoDate(start), toIsoDate(addDays(start, 6))];
      }
      if(key.startsWith("a")){
        return [`${base.getFullYear()}-01-01`, `${base.getFullYear()}-12-31`];
      }
      const start = new Date(base.getFullYear(), base.getMonth(), 1);
      const end = new Date(base.getFullYear(), base.getMonth() + 1, 0);
      return [toIsoDate(start), toIsoDate(end)];
    }
    function inRange(value, start, end){ return value && value >= start && value <= end; }
    function dateList(start, end, limit=45){
      const rows = [];
      let cursor = parseIsoDate(start);
      const stop = parseIsoDate(end);
      while(cursor <= stop && rows.length < limit){ rows.push(toIsoDate(cursor)); cursor = addDays(cursor, 1); }
      return rows;
    }
    function unavailable(status){
      const text = String(status || "").toUpperCase();
      return text.includes("NO DISPONIBLE") || text.includes("FUERA") || text.includes("NO DISP") || text.includes("REPARACION") || text.includes("REPARACIÓN") || text.includes("MANTENIMIENTO");
    }
    function normalizedText(value){
      return String(value || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    }
    function equipmentKeys(value){
      const text = normalizedText(value);
      if(!text) return [];
      const candidates = new Set([text, text.split(/\s+-\s+|\s+\(|\s+/)[0]]);
      const keys = new Set();
      candidates.forEach(candidate => {
        const raw = normalizedText(candidate).replace(/[^A-Z0-9]/g, "");
        if(!raw) return;
        keys.add(raw);
        keys.add(raw.replace(/([A-Z]+)0+(\d)/g, "$1$2"));
      });
      return [...keys].filter(Boolean);
    }
    const stoppageRules = [
      ["Hidraulico", ["HIDRAUL", "HCO", "MANGUERA", "BOMBA", "CILINDRO", "FUGA"]],
      ["Electrico", ["ELECT", "BATERIA", "ALTERNADOR", "CABLE", "SENSOR", "CORTO", "FUSIBLE"]],
      ["Motor/Diesel", ["MOTOR", "DIESEL", "COMBUST", "INYECTOR", "TURBO", "ACEITE MOTOR", "15W40"]],
      ["Transmision", ["TRANSM", "CONVERTIDOR", "DIFERENCIAL", "MANDO", "CAJA", "EJE"]],
      ["Frenos", ["FRENO", "BALATA", "PASTILLA"]],
      ["Llantas", ["LLANTA", "NEUMATIC", "RIN", "PONCH", "PRESION"]],
      ["Estructural", ["ESTRUCT", "CHASIS", "BRAZO", "CUCHARON", "FISURA", "SOLDAD"]],
      ["Perforacion", ["PERFOR", "BARRA", "BROCA", "PERCUS", "COMPRESOR", "PIERNA"]],
      ["Lubricacion", ["LUBRIC", "GRASA", "ENGRASE", "REFRIGERANTE", "SAE", "ISO 68"]],
      ["Operacion", ["OPERADOR", "OPERACION", "GALERIA", "ACCESO", "TRAFICO", "VENTILACION"]],
    ];
    function classifyStoppage(...values){
      const text = normalizedText(values.filter(Boolean).join(" "));
      if(!text) return "Sin clasificar";
      for(const [label, tokens] of stoppageRules){
        if(tokens.some(token => text.includes(token))) return label;
      }
      return "Otros";
    }
    function backlogLevel(score){ return score >= 80 ? "ALTA" : (score >= 50 ? "MEDIA" : "BAJA"); }
    function levelClass(level){ return level === "ALTA" ? "bad" : (level === "MEDIA" ? "warn" : "ok"); }
    function flowClass(status){
      const text = String(status || "").toUpperCase();
      if(text.includes("ATEND")) return "ok";
      if(text.includes("CANCEL")) return "bad";
      if(text.includes("PROCESO")) return "warn";
      return "warn";
    }
    function kpiStatusFromCondition(value){
      const text = normalizedText(value);
      if(!text) return "";
      if(text.includes("STAND")) return "Stand By";
      if(unavailable(text)) return "No Disponible";
      if(text.includes("DISPONIBLE")) return "Disponible";
      if(text.includes("OPERATIVA")) return "Operativa";
      return "No Disponible";
    }
    function availabilityStatusMap(){
      const map = new Map();
      (portal.availability || []).forEach(row => {
        const status = kpiStatusFromCondition(row.condition);
        if(!status) return;
        [...equipmentKeys(row.eco), ...equipmentKeys(row.equipment)].forEach(key => {
          if(key && !map.has(key)) map.set(key, status);
        });
      });
      return map;
    }
    function availabilityStatusForEquipment(eq, map){
      for(const key of [...equipmentKeys(eq.code || eq.equipment_code), ...equipmentKeys(eq.description)]){
        if(map.has(key)) return map.get(key);
      }
      return "";
    }
    function groupMatches(eq, group){
      const key = String(group || "Todos").toUpperCase();
      const code = String(eq.code || eq.equipment_code || "").toUpperCase();
      const text = `${code} ${eq.description || ""} ${eq.family || ""}`.toUpperCase();
      if(key.includes("TODOS")) return true;
      if(key.includes("BARRENACION")) return code.startsWith("JL") || code.startsWith("JA") || text.includes("JUMBO") || text.includes("BARREN") || text.includes("ANCLADOR");
      if(key.includes("REZAGADO")) return code.startsWith("ST") || text.includes("SCOOP") || text.includes("CATERPILLAR") || text.includes("EPROC") || text.includes("R1300") || text.includes("R1600") || text.includes("REZAG");
      if(key.includes("ACARREO")) return ["CBP001","MG044"].includes(code.replace(/[^A-Z0-9]/g, ""));
      if(key.includes("UTILITARIO")) return ["RET009","RET010"].includes(code.replace(/[^A-Z0-9]/g, ""));
      return true;
    }
    function kpiRequiredComponent(code){
      const normalized = String(code || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
      if(normalized.startsWith("JL") || normalized.startsWith("JA")) return "ELECT";
      if(normalized.startsWith("ST") || normalized.startsWith("RET") || ["CBP001","MG044"].includes(normalized)) return "DIESEL";
      return "";
    }
    function kpiCaptureComponentMatches(equipmentCode, componentName){
      const required = kpiRequiredComponent(equipmentCode);
      return !required || normalizedText(componentName).includes(required);
    }
    function metric(period, worked, mp, mc, stops, missionHours=12){
      const available = Math.max(Number(period || 0) - Number(mp || 0) - Number(mc || 0), 0);
      const availability = period > 0 ? Math.max(Math.min((available / period) * 100, 100), 0) : 0;
      const utilization = available > 0 ? Math.max(Math.min((Number(worked || 0) / available) * 100, 100), 0) : 0;
      const stopCount = Math.max(Number(stops || 0), 0);
      let tmef = 0;
      let tmpr = 0;
      let reliability = 0;
      if(stopCount){
        tmef = Number(worked || 0) ? (Number(worked || 0) / stopCount) : 0;
        tmpr = Number(mc || 0) / stopCount;
        reliability = tmef && missionHours ? Math.max(Math.min(Math.exp(-(Number(missionHours || 0) / tmef)) * 100, 100), 0) : 0;
      } else {
        tmef = Number(worked || 0);
        tmpr = 0;
        reliability = Number(worked || 0) > 0 ? 100 : 0;
      }
      return {
        available,
        availability,
        utilization,
        tmef,
        tmpr,
        reliability,
      };
    }
    function simNumber(id, fallback){
      const el = $(id);
      const value = Number(el ? el.value : fallback);
      return Number.isFinite(value) ? value : Number(fallback || 0);
    }
    function kpiSimulationActive(){
      const group = String($("kpiGroup").value || "").toUpperCase();
      return Boolean($("kpiSimEnabled")?.checked) && !group.includes("ACEITE") && !group.includes("LLANTA") && !group.includes("DIESEL");
    }
    function currentKpiSettings(){
      const base = portal.settings || {};
      if(!kpiSimulationActive()) return base;
      return {
        ...base,
        meta_availability: simNumber("kpiSimMetaAvailability", base.meta_availability || 85),
        meta_utilization: simNumber("kpiSimMetaUtilization", base.meta_utilization || 75),
        meta_reliability: simNumber("kpiSimMetaReliability", base.meta_reliability || 80),
        meta_tmef: simNumber("kpiSimMetaTmef", base.meta_tmef || 8),
        meta_tmpr: simNumber("kpiSimMetaTmpr", base.meta_tmpr || 4),
        reliability_mission_hours: simNumber("kpiSimMission", base.reliability_mission_hours || base.mission_hours || 12),
      };
    }
    function initializeKpiSimulationSettings(force=false){
      if(kpiSimulationInitialized && !force) return;
      const settings = portal.settings || {};
      $("kpiSimMetaAvailability").value = Number(settings.meta_availability || 85);
      $("kpiSimMetaUtilization").value = Number(settings.meta_utilization || 75);
      $("kpiSimMetaReliability").value = Number(settings.meta_reliability || 80);
      $("kpiSimMetaTmef").value = Number(settings.meta_tmef || 8);
      $("kpiSimMetaTmpr").value = Number(settings.meta_tmpr || 4);
      $("kpiSimMission").value = Number(settings.reliability_mission_hours || settings.mission_hours || 12);
      kpiSimulationInitialized = true;
    }
    function simulatedKpiReport(report){
      if(!kpiSimulationActive()) return report;
      const missionHours = simNumber("kpiSimMission", (portal.settings || {}).reliability_mission_hours || (portal.settings || {}).mission_hours || 12);
      const factors = {
        period: Math.max(simNumber("kpiSimPeriod", 100), 0) / 100,
        worked: Math.max(simNumber("kpiSimWorked", 100), 0) / 100,
        mp: Math.max(simNumber("kpiSimMp", 100), 0) / 100,
        mc: Math.max(simNumber("kpiSimMc", 100), 0) / 100,
        stops: Math.max(simNumber("kpiSimStops", 100), 0) / 100,
      };
      const rows = report.rows.map(source => {
        const row = {...source};
        row.period = Math.max(Number(source.period || 0) * factors.period, 0);
        row.worked = Math.max(Number(source.worked || 0) * factors.worked, 0);
        row.mp = Math.max(Number(source.mp || 0) * factors.mp, 0);
        row.mc = Math.max(Number(source.mc || 0) * factors.mc, 0);
        row.stops = Math.max(Math.round(Number(source.stops || 0) * factors.stops), 0);
        row.out = row.worked <= 0 && unavailable(row.status);
        const values = row.out ? {available:0, availability:0, utilization:0, tmef:0, tmpr:0, reliability:0} : metric(row.period, row.worked, row.mp, row.mc, row.stops, missionHours);
        Object.assign(row, values);
        row.availabilityText = row.out ? "FUERA" : pct(row.availability);
        row.utilizationText = row.out ? "FUERA" : pct(row.utilization);
        return row;
      });
      const totals = rows.reduce((acc, row) => {
        acc.period += Number(row.period || 0); acc.worked += Number(row.worked || 0); acc.mp += Number(row.mp || 0); acc.mc += Number(row.mc || 0); acc.stops += Number(row.stops || 0); acc.available += Number(row.available || 0);
        return acc;
      }, {period:0, worked:0, mp:0, mc:0, stops:0, available:0});
      totals.availability = totals.period ? (totals.available / totals.period) * 100 : 0;
      totals.utilization = totals.available ? (totals.worked / totals.available) * 100 : 0;
      totals.tmef = totals.stops ? (totals.worked ? totals.worked / totals.stops : 0) : totals.worked;
      totals.tmpr = totals.stops ? totals.mc / totals.stops : 0;
      totals.reliability = totals.tmef && missionHours ? Math.max(Math.min(Math.exp(-(missionHours / totals.tmef)) * 100, 100), 0) : (totals.worked > 0 && !totals.stops ? 100 : 0);
      return {...report, rows, totals, simulation:{enabled:true, name:$("kpiSimName").value || "Escenario KPI"}};
    }
    function setOptions(selectId, options, allLabel="Todos"){
      const select = $(selectId);
      const current = select.value;
      select.innerHTML = `<option value="">${esc(allLabel)}</option>` + options.map(item => `<option value="${esc(item.value)}">${esc(item.label)}</option>`).join("");
      if([...select.options].some(opt => opt.value === current)) select.value = current;
    }
    function portalEquipment(){
      const rows = Array.isArray(portal.equipment) ? portal.equipment : [];
      return rows.filter(e => e && (e.code || e.equipment_code));
    }
    function desktopKpiReport(group, start, end){
      const reports = portal.kpi_reports && typeof portal.kpi_reports === "object" ? portal.kpi_reports : {};
      const source = reports[group];
      if(!source || source.start !== start || source.end !== end) return null;
      const rows = Array.isArray(source.rows) ? source.rows.map(sourceRow => {
        const period = Number(sourceRow.period ?? sourceRow.period_hours ?? 0);
        const mp = Number(sourceRow.mp ?? sourceRow.mp_hours ?? 0);
        const mc = Number(sourceRow.mc ?? sourceRow.mc_hours ?? 0);
        const worked = Number(sourceRow.worked ?? sourceRow.worked_hours ?? 0);
        const availabilityText = String(sourceRow.availabilityText || sourceRow.availability_text || "").trim();
        const utilizationText = String(sourceRow.utilizationText || sourceRow.utilization_text || "").trim();
        const out = availabilityText.toUpperCase() === "FUERA" || utilizationText.toUpperCase() === "FUERA";
        const availability = Number(sourceRow.availability || 0);
        const utilization = Number(sourceRow.utilization || 0);
        return {
          code: sourceRow.code || "",
          description: sourceRow.description || "",
          family: sourceRow.family || "",
          status: sourceRow.status || (out ? "FUERA" : ""),
          period,
          worked,
          mp,
          mc,
          stops: Number(sourceRow.stops || 0),
          available: out ? 0 : Math.max(period - mp - mc, 0),
          availability,
          utilization,
          tmef: Number(sourceRow.tmef || 0),
          tmpr: Number(sourceRow.tmpr || 0),
          reliability: Number(sourceRow.reliability || 0),
          out,
          availabilityText: availabilityText || (out ? "FUERA" : pct(availability)),
          utilizationText: utilizationText || (out ? "FUERA" : pct(utilization)),
        };
      }) : [];
      return {
        group: source.group || group,
        start,
        end,
        rows,
        totals: source.totals || {period:0, worked:0, mp:0, mc:0, stops:0, availability:0, utilization:0, tmef:0, tmpr:0, reliability:0},
        source: "desktop-kpi-report",
      };
    }
    function renderReqEquipmentOptions(){
      const select = $("reqEquipment");
      const current = select.value || "PARA STOCK";
      const seen = new Set();
      const options = [];
      function addOption(value, label){
        const cleanValue = String(value || "").trim();
        if(!cleanValue || seen.has(cleanValue)) return;
        seen.add(cleanValue);
        options.push({value: cleanValue, label: label || cleanValue});
      }
      addOption("PARA STOCK", "PARA STOCK");
      addOption("TALLER", "TALLER");
      [...portalEquipment(), ...(Array.isArray(data.equipment) ? data.equipment : [])].forEach(e => {
        const code = e.code || e.equipment_code || "";
        const description = e.description || e.family || "";
        addOption(code, description ? `${code} - ${description}` : code);
      });
      if(current && !seen.has(current)) addOption(current, current);
      select.innerHTML = options.map(item => `<option value="${esc(item.value)}">${esc(item.label)}</option>`).join("");
      select.value = seen.has(current) ? current : "PARA STOCK";
    }
    function renderPortalSelectors(){
      const period = portal.period || {};
      const today = toIsoDate(new Date());
      if(!monthlyPeriodInitialized){
        $("monthlyMonth").value = String(period.month || (new Date()).getMonth() + 1);
        $("monthlyYear").value = String(period.year || (new Date()).getFullYear());
        monthlyPeriodInitialized = true;
      }
      if(!$("weeklyBase").value) $("weeklyBase").value = period.end || period.start || today;
      if(!$("weeklyStart").value || !$("weeklyEnd").value) applyWeeklyPeriod(false);
      if(!$("kpiStart").value) $("kpiStart").value = period.start || today;
      if(!$("kpiEnd").value) $("kpiEnd").value = period.end || today;
      if(!$("fichaStart").value) $("fichaStart").value = period.start || today;
      if(!$("fichaEnd").value) $("fichaEnd").value = period.end || today;
      initializeKpiSimulationSettings();
      if(!$("prBase").value) $("prBase").value = period.start || today;
      if(!$("backlogStart").value) $("backlogStart").value = period.start || today;
      if(!$("backlogEnd").value) $("backlogEnd").value = period.end || today;
      if(!$("srvStart").value) $("srvStart").value = period.capture_start || period.start || today;
      if(!$("srvEnd").value) $("srvEnd").value = period.capture_end || period.end || today;
      if(!$("prevExecDate").value) $("prevExecDate").value = today;
      if(!$("woDate").value) $("woDate").value = today;
      if(!$("bitStart").value) $("bitStart").value = period.start || today;
      if(!$("bitEnd").value) $("bitEnd").value = period.end || today;
      if(!$("auditStart").value) $("auditStart").value = period.start || today;
      if(!$("auditEnd").value) $("auditEnd").value = period.end || today;
      if(!$("dieselStart").value) $("dieselStart").value = diesel.start || period.start || today;
      if(!$("dieselEnd").value) $("dieselEnd").value = diesel.end || period.end || today;
      if(!$("dieselBase").value) $("dieselBase").value = diesel.start || period.start || today;
      if(!$("dieselDate").value) $("dieselDate").value = today;
      if(!$("dieselDayDate").value) $("dieselDayDate").value = today;
      $("dieselMeta").value = diesel.meta_lh || (portal.settings || {}).meta_diesel_lh || $("dieselMeta").value || 25;
      const rawGroups = portal.kpi_groups && portal.kpi_groups.length ? [...portal.kpi_groups] : ["Todos los equipos", "Equipos de Barrenacion", "Equipos de Rezagado", "Acarreo", "Equipo Utilitario"];
      if(!rawGroups.includes("Acarreo")) rawGroups.splice(Math.min(rawGroups.length, 3), 0, "Acarreo");
      if(!rawGroups.includes("Equipo Utilitario")) rawGroups.splice(Math.min(rawGroups.length, 4), 0, "Equipo Utilitario");
      ["KPI Aceites", "KPI Llantas", "KPI Diesel"].forEach(group => { if(!rawGroups.includes(group)) rawGroups.push(group); });
      const groups = rawGroups.map(g => ({value:g, label:g}));
      const previousGroup = $("kpiGroup").value;
      $("kpiGroup").innerHTML = groups.map(g => `<option value="${esc(g.value)}">${esc(g.label)}</option>`).join("");
      $("kpiGroup").value = previousGroup && groups.some(g => g.value === previousGroup) ? previousGroup : groups[0]?.value || "";
      const equipmentOptions = portalEquipment().map(e => ({value:e.code || e.equipment_code, label:`${e.code || e.equipment_code} - ${e.description || e.family || ""}`}));
      setOptions("fichaEquipment", equipmentOptions, "Selecciona");
      if(!$("fichaEquipment").value && equipmentOptions.length) $("fichaEquipment").value = equipmentOptions[0].value;
      setOptions("prEquipment", equipmentOptions, "Todos");
      setOptions("srvEquipment", equipmentOptions, "Todos");
      setOptions("prevExecEquipment", equipmentOptions, "Selecciona");
      setOptions("woEquipment", equipmentOptions, "Selecciona");
      setOptions("woFilterEquipment", equipmentOptions, "Todos");
      setOptions("bitEquipment", equipmentOptions, "Todos");
      setOptions("spareEquipment", equipmentOptions, "Todos");
      const auditModules = [...new Set((portal.audit_log || []).map(row => row.module).filter(Boolean))].sort();
      setOptions("auditModule", auditModules.map(module => ({value:module, label:module})), "Todos");
      if(!$("capDate").value) $("capDate").value = today;
      setOptions("capEquipment", equipmentOptions, "Selecciona");
      if(!$("capEquipment").value && equipmentOptions.length) $("capEquipment").value = equipmentOptions[0].value;
      renderCaptureComponents();
      renderDieselSelectors();
      renderReqEquipmentOptions();
      if(!$("tireTrackDate").value) $("tireTrackDate").value = today;
      setOptions("tireTrackEquipment", tireTrackEquipmentOptions(), "Sin equipo");
    }
    function calculateKpiRows(groupOverride=null, startOverride=null, endOverride=null){
      const group = groupOverride || $("kpiGroup").value || "Todos los equipos";
      const start = startOverride || $("kpiStart").value;
      const end = endOverride || $("kpiEnd").value;
      const precomputed = desktopKpiReport(group, start, end);
      if(precomputed) return precomputed;
      const settings = portal.settings || {};
      const shiftHours = Number(settings.shift_hours || 9);
      const dailyHours = shiftHours * Number(settings.turns_per_day || 2);
      const missionHours = Number(settings.reliability_mission_hours || settings.mission_hours || 12);
      const days = Math.max(Math.round((parseIsoDate(end) - parseIsoDate(start)) / 86400000) + 1, 1);
      const captures = (portal.captures || []).filter(c => inRange(c.work_date, start, end));
      const grouped = {};
      const availabilityStatuses = availabilityStatusMap();
      portalEquipment().filter(eq => groupMatches(eq, group)).forEach(eq => {
        const code = eq.code || eq.equipment_code || "";
        const availabilityStatus = availabilityStatusForEquipment(eq, availabilityStatuses);
        grouped[code] = {code, description:eq.description || "", family:eq.family || "", status:availabilityStatus || eq.status || "Disponible", availabilityStatus, captureStatus:"", captureStatusOrder:"", period:days * dailyHours, worked:0, mp:0, mc:0, stops:0, unavailableCount:0};
      });
      captures.forEach(c => {
        const code = c.equipment_code || c.code || "";
        if(!grouped[code]) return;
        if(!kpiCaptureComponentMatches(code, c.component || c.component_name)) return;
        const row = grouped[code];
        const mp = Number(c.mp_hours || 0);
        let mc = Number(c.mc_hours || 0);
        const captureStatus = String(c.status || "").trim();
        if(captureStatus){
          const captureOrder = `${c.work_date || ""}-${String(c.id || "").padStart(10, "0")}`;
          if(!row.captureStatusOrder || captureOrder >= row.captureStatusOrder){
            row.captureStatus = captureStatus;
            row.captureStatusOrder = captureOrder;
            if(!row.availabilityStatus) row.status = captureStatus;
          }
        }
        if(unavailable(c.status)){
          const base = String(c.shift || "").toUpperCase() === "GENERAL" ? dailyHours : shiftHours;
          mc += Math.max(base - mp - mc, 0);
          row.unavailableCount += 1;
          if(!row.availabilityStatus) row.status = c.status || "FUERA";
        }
        row.worked += Number(c.worked_hours || 0);
        row.mp += mp;
        row.mc += mc;
        row.stops += Number(c.stops || 0);
      });
      const rows = Object.values(grouped).sort((a,b) => a.code.localeCompare(b.code)).map(row => {
        if(row.availabilityStatus) row.status = row.availabilityStatus;
        else if(row.captureStatus) row.status = row.captureStatus;
        const out = row.worked <= 0 && (row.unavailableCount > 0 || unavailable(row.status));
        const m = out ? {available:0, availability:0, utilization:0, tmef:0, tmpr:0, reliability:0} : metric(row.period, row.worked, row.mp, row.mc, row.stops, missionHours);
        return {...row, ...m, out, availabilityText: out ? "FUERA" : pct(m.availability), utilizationText: out ? "FUERA" : pct(m.utilization)};
      });
      const totals = rows.reduce((acc, row) => {
        acc.period += row.period; acc.worked += row.worked; acc.mp += row.mp; acc.mc += row.mc; acc.stops += row.stops; acc.available += row.available;
        return acc;
      }, {period:0, worked:0, mp:0, mc:0, stops:0, available:0});
      totals.availability = totals.period ? (totals.available / totals.period) * 100 : 0;
      totals.utilization = totals.available ? (totals.worked / totals.available) * 100 : 0;
      totals.tmef = totals.stops ? (totals.worked ? totals.worked / totals.stops : 0) : totals.worked;
      totals.tmpr = totals.stops ? totals.mc / totals.stops : 0;
      totals.reliability = totals.tmef && missionHours ? Math.max(Math.min(Math.exp(-(missionHours / totals.tmef)) * 100, 100), 0) : (totals.worked > 0 && !totals.stops ? 100 : 0);
      return {group, start, end, rows, totals};
    }
    const kpiMetricTabs = [
      {key:"availability", label:"% Disponibilidad"},
      {key:"utilization", label:"% Utilizacion"},
      {key:"reliability", label:"Confiabilidad"},
      {key:"tmef", label:"TMEF"},
      {key:"tmpr", label:"TMPR"},
    ];
    function kpiMetricValue(row, metric){
      if(metric === "utilization") return Number(row.utilization || 0);
      if(metric === "reliability") return Number(row.reliability || 0);
      if(metric === "tmef") return Number(row.tmef || 0);
      if(metric === "tmpr") return Number(row.tmpr || 0);
      return Number(row.availability || 0);
    }
    function kpiMetricText(row, metric){
      if((metric === "availability" || metric === "utilization" || metric === "reliability") && row.out) return "FUERA";
      const value = kpiMetricValue(row, metric);
      return (metric === "availability" || metric === "utilization" || metric === "reliability") ? pct(value) : one(value);
    }
    function kpiMetricAxisMax(metric, rows, target){
      if(metric === "availability" || metric === "utilization" || metric === "reliability") return 120;
      const peak = Math.max(Number(target || 0), ...rows.map(row => kpiMetricValue(row, metric)), 1);
      if(peak <= 5) return 5;
      if(peak <= 10) return 10;
      const step = peak <= 30 ? 5 : 10;
      return Math.ceil((peak * 1.18) / step) * step;
    }
    function kpiMetricChartHtml(report, settings){
      const metric = kpiMetricTabs.some(item => item.key === selectedKpiMetric) ? selectedKpiMetric : "availability";
      const targets = {
        availability: Number(settings.meta_availability || 85),
        utilization: Number(settings.meta_utilization || 75),
        reliability: Number(settings.meta_reliability || 80),
        tmef: Number(settings.meta_tmef || 8),
        tmpr: Number(settings.meta_tmpr || 4),
      };
      const axisMax = kpiMetricAxisMax(metric, report.rows, targets[metric]);
      const chartBars = report.rows.map(row => {
        const value = kpiMetricValue(row, metric);
        const h = Math.max(Math.min(value / Math.max(axisMax, 1), 1) * 210, 4);
        const outClass = row.out && (metric === "availability" || metric === "utilization" || metric === "reliability") ? "out" : "";
        const title = `${row.code} ${kpiMetricTabs.find(item => item.key === metric)?.label || ""}: ${kpiMetricText(row, metric)}`;
        return `<div class="chart-bar ${outClass}" title="${esc(title)}"><span>${esc(kpiMetricText(row, metric))}</span><i style="--h:${h}px"></i><b>${esc(row.code)}</b></div>`;
      }).join("") || `<p class="muted">Sin datos KPI para el periodo.</p>`;
      const tabs = kpiMetricTabs.map(item => `<button type="button" data-kpi-metric="${esc(item.key)}" class="${item.key === metric ? "active" : ""}">${esc(item.label)}</button>`).join("");
      const targetTop = Math.max(Math.min(100 - ((targets[metric] / Math.max(axisMax, 1)) * 100), 94), 6);
      const targetText = metric === "tmef" || metric === "tmpr" ? one(targets[metric]) : pct(targets[metric]);
      return `<div class="kpi-chart-head"><b>KPI</b><div class="kpi-mini-tabs">${tabs}</div></div><div class="chart-plot"><span class="chart-target" style="top:${targetTop}%"><b>Meta ${esc(targetText)}</b></span>${chartBars}</div>`;
    }
    function bindKpiMetricTabs(){
      document.querySelectorAll("#kpiChart [data-kpi-metric]").forEach(button => {
        button.addEventListener("click", () => {
          selectedKpiMetric = button.dataset.kpiMetric || "availability";
          renderDashboard();
        });
      });
    }
    function oilColumns(){
      const required = [
        {label:"15W40", key:"oil_motor_15w40"},
        {label:"ISO 68", key:"oil_hco_iso68"},
        {label:"SAE 30", key:"oil_trans_sae30"},
        {label:"SAE 50", key:"oil_sae50"},
        {label:"85W140", key:"oil_85w140"},
        {label:"ALMO", key:"almo_liters"},
        {label:"Refrigerante", key:"coolant_liters"},
        {label:"VG100", key:"oil_hyd_vg100"},
        {label:"ATF", key:"atf_liters"},
      ];
      const cols = portal.oil_kpi && Array.isArray(portal.oil_kpi.columns) ? [...portal.oil_kpi.columns] : [];
      required.forEach(item => {
        if(!cols.some(col => col.key === item.key)) cols.push(item);
      });
      return cols;
    }
    function oilGroupFor(eq){
      const code = String(eq.code || eq.equipment_code || "").toUpperCase();
      const text = `${code} ${eq.description || ""} ${eq.family || ""}`.toUpperCase();
      if(code.startsWith("JL") || code.startsWith("JA") || text.includes("JUMBO") || text.includes("ANCLADOR")) return "BARRENACION";
      if(code.startsWith("ST") || text.includes("SCOOP") || text.includes("CATERPILLAR") || text.includes("EPROC") || text.includes("R1300") || text.includes("R1600")) return "REZAGADO";
      return "UTILITARIO";
    }
    function oilMainColumns(){
      const available = oilColumns();
      return [
        {label:"Motor 15W40", key:"oil_motor_15w40"},
        {label:"ISO 68", key:"oil_hco_iso68"},
        {label:"SAE 30", key:"oil_trans_sae30"},
        {label:"SAE 50", key:"oil_sae50"},
        {label:"85W140", key:"oil_85w140"},
      ].map(item => available.find(col => col.key === item.key) || item);
    }
    function oilOrderColumns(){
      return [
        {label:"ALMO", key:"almo_liters"},
        {label:"85W140", key:"oil_85w140"},
        {label:"Compresor ISO 32", key:"oil_compressor_iso32"},
        {label:"HCO ISO 68", key:"oil_hco_iso68"},
        {label:"Motor 15W40", key:"oil_motor_15w40"},
        {label:"Trans. SAE 30", key:"oil_trans_sae30"},
        {label:"sin clasificar", key:"oil_liters"},
        {label:"Refrigerante", key:"coolant_liters"},
      ];
    }
    function oilDays(start, end){
      const a = parseIsoDate(start);
      const b = parseIsoDate(end);
      return Math.max(Math.round((b - a) / 86400000) + 1, 1);
    }
    function two(v){ return `${Number(v || 0).toFixed(2)}`; }
    function oilRowsForRange(start, end, cols=oilMainColumns()){
      const source = portal.oil_kpi || {};
      if(source.start === start && source.end === end && Array.isArray(source.rows) && source.rows.length){
        const rows = source.rows.map(row => {
          const out = {
            code: row.code || "",
            description: row.description || "",
            group: row.group || "UTILITARIO",
            period_hours: Number(row.period_hours || row.period || 0),
            worked_hours: Number(row.worked_hours || row.worked || 0),
            total_liters: 0,
          };
          cols.forEach(col => {
            out[col.key] = Number(row[col.key] || 0);
            out.total_liters += out[col.key];
          });
          return out;
        });
        const totals = rows.reduce((acc, row) => {
          acc.period += Number(row.period_hours || 0);
          acc.worked += Number(row.worked_hours || 0);
          acc.worked_hours += Number(row.worked_hours || 0);
          acc.total_liters += Number(row.total_liters || 0);
          cols.forEach(col => acc[col.key] = (acc[col.key] || 0) + Number(row[col.key] || 0));
          return acc;
        }, {period:0, worked:0, worked_hours:0, total_liters:0});
        return {start, end, cols, rows, totals, days:oilDays(start, end)};
      }
      const days = oilDays(start, end);
      const settings = portal.settings || {};
      const dailyHours = (Number(settings.shift_hours || 9) || 9) * (Number(settings.turns_per_day || 2) || 2);
      const grouped = {};
      portalEquipment().forEach(eq => {
        const code = eq.code || eq.equipment_code || "";
        grouped[code] = {code, description:eq.description || "", group:oilGroupFor(eq), period_hours:days * dailyHours, worked_hours:0, total_liters:0};
        cols.forEach(col => grouped[code][col.key] = 0);
      });
      (portal.captures || []).filter(row => inRange(row.work_date, start, end)).forEach(row => {
        const code = row.equipment_code || row.code || "";
        if(!grouped[code]) {
          grouped[code] = {code, description:"", group:"UTILITARIO", period_hours:days * dailyHours, worked_hours:0, total_liters:0};
          cols.forEach(col => grouped[code][col.key] = 0);
        }
        grouped[code].worked_hours += Number(row.worked_hours || 0);
        cols.forEach(col => {
          const value = Number(row[col.key] || 0);
          grouped[code][col.key] += value;
          grouped[code].total_liters += value;
        });
      });
      preventiveExecutionRows().filter(row => isPreventiveClosed(row) && inRange(row.close_date || row.service_date, start, end)).forEach(row => {
        const code = row.equipment_code || "";
        if(!code) return;
        if(!grouped[code]) {
          grouped[code] = {code, description:row.equipment_description || "", group:"UTILITARIO", period_hours:days * dailyHours, worked_hours:0, total_liters:0};
          cols.forEach(col => grouped[code][col.key] = 0);
        }
        cols.forEach(col => {
          const value = Number(row[col.key] || 0);
          grouped[code][col.key] += value;
          grouped[code].total_liters += value;
        });
      });
      const groupRank = {BARRENACION:1, REZAGADO:2, UTILITARIO:3};
      const rows = Object.values(grouped).sort((a,b) => (groupRank[a.group] || 9) - (groupRank[b.group] || 9) || a.code.localeCompare(b.code));
      const totals = rows.reduce((acc, row) => {
        acc.period += Number(row.period_hours || 0);
        acc.worked += row.worked_hours;
        acc.worked_hours += row.worked_hours;
        acc.total_liters += row.total_liters;
        cols.forEach(col => acc[col.key] = (acc[col.key] || 0) + Number(row[col.key] || 0));
        return acc;
      }, {period:0, worked:0, worked_hours:0, total_liters:0});
      return {start, end, cols, rows, totals, days};
    }
    function oilRowsForPeriod(){
      return oilRowsForRange($("kpiStart").value, $("kpiEnd").value);
    }
    function oilMonthLabel(value){
      const months = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
      const date = parseIsoDate(value);
      return `${months[date.getMonth()]}-${String(date.getFullYear()).slice(-2)}`;
    }
    function oilMetricSection(title, periodValue, accumulatedValue, tone="teal"){
      const periodLine = tone === "teal" ? "" : "oil-red";
      const totalLine = tone === "teal" ? "" : "oil-darkred";
      return `<div class="oil-metric-section">
        <div class="oil-metric-title">${esc(title)}</div>
        <div class="oil-metric-grid">
          <div class="oil-metric-cell"><strong>${two(periodValue)}</strong><div class="oil-metric-line ${periodLine}"><i></i></div><span>Consumo semanal</span></div>
          <div class="oil-metric-cell"><strong>${two(accumulatedValue)}</strong><div class="oil-metric-line ${totalLine}"><i></i></div><span>Consumo total acumulado</span></div>
        </div>
      </div>`;
    }
    function oilChartHtml(report){
      const cols = oilMainColumns();
      const colors = {
        oil_motor_15w40:"#4472c4",
        oil_hco_iso68:"#ed7d31",
        oil_trans_sae30:"#a5a5a5",
        oil_sae50:"#ffc000",
        oil_85w140:"#5b9bd5",
      };
      let chartRows = [...report.rows].filter(row => cols.some(col => Number(row[col.key] || 0) > 0));
      chartRows.sort((a,b) => cols.reduce((sum,col) => sum + Number(b[col.key] || 0), 0) - cols.reduce((sum,col) => sum + Number(a[col.key] || 0), 0));
      chartRows = chartRows.slice(0, 6);
      if(!chartRows.length) chartRows = report.rows.slice(0, 6);
      const peak = Math.max(1, ...chartRows.flatMap(row => cols.map(col => Number(row[col.key] || 0))));
      const axisMax = Math.max(100, Math.ceil((peak * 1.25) / 10) * 10);
      const ticks = Array.from({length:6}, (_, idx) => Math.round(axisMax - (axisMax / 5) * idx));
      const clusters = chartRows.map(row => `<div class="oil-cluster">
        <div class="oil-bars-stack">${cols.map(col => {
          const value = Number(row[col.key] || 0);
          const height = Math.max((value / axisMax) * 250, value ? 2 : 1);
          const label = Math.abs(value - Math.round(value)) < .01 ? String(Math.round(value)) : one(value);
          return `<div class="oil-series-bar" style="height:${height}px;background:${colors[col.key] || "#64748b"}"><b>${esc(label)}</b></div>`;
        }).join("")}</div>
        <div class="oil-cluster-label">${esc(row.code || "-")}</div>
      </div>`).join("") || `<p class="muted">Sin datos de aceites.</p>`;
      return `<div class="oil-chart-grid">
        <div class="oil-axis">${ticks.map(value => `<span>${esc(value)}</span>`).join("")}</div>
        <div class="oil-plot">${clusters}</div>
        <div class="oil-legend">${cols.map(col => `<span><i style="background:${colors[col.key] || "#64748b"}"></i>${esc(col.label)}</span>`).join("")}</div>
      </div>`;
    }
    function oilGroupTotals(report){
      const order = ["BARRENACION","REZAGADO","UTILITARIO"];
      const totals = {};
      order.forEach(group => {
        totals[group] = {period:0, worked:0};
        report.cols.forEach(col => totals[group][col.key] = 0);
      });
      report.rows.forEach(row => {
        const group = totals[row.group] ? row.group : "UTILITARIO";
        totals[group].period += Number(row.period_hours || 0);
        totals[group].worked += Number(row.worked_hours || 0);
        report.cols.forEach(col => totals[group][col.key] += Number(row[col.key] || 0));
      });
      return totals;
    }
    function oilValueCell(value, zeroClass=true){
      const numeric = Number(value || 0);
      return `<td class="${zeroClass && numeric === 0 ? "oil-zero" : ""}">${two(numeric)}</td>`;
    }
    function oilReportTableHtml(report){
      const groupOrder = ["BARRENACION","REZAGADO","UTILITARIO"];
      const groupLabels = {BARRENACION:"ACUMULADO EQ'S DE<br/>BARRENACION", REZAGADO:"EQUIPO REZAGADO", UTILITARIO:"EQUIPO UTILITARIO"};
      const groupTotals = oilGroupTotals(report);
      const body = [];
      groupOrder.forEach(group => {
        report.rows.filter(row => row.group === group).forEach(row => {
          body.push(`<tr><td>${esc(row.code)}</td><td>${esc(row.description || "")}</td><td>${one(row.period_hours)}</td><td>${Number(row.worked_hours || 0) ? one(row.worked_hours) : ""}</td>${report.cols.map(col => oilValueCell(row[col.key])).join("")}</tr>`);
        });
        const subtotal = groupTotals[group];
        if(report.rows.some(row => row.group === group)){
          body.push(`<tr class="oil-subtotal"><td></td><td><b>${groupLabels[group]}</b></td><td><b>${one(subtotal.period)}</b></td><td><b>${one(subtotal.worked)}</b></td>${report.cols.map(col => `<td><b>${two(subtotal[col.key])}</b></td>`).join("")}</tr>`);
        }
      });
      body.push(`<tr class="oil-total"><td></td><td><b>Total de Aceite Utilizado</b></td><td></td><td><b>${one(report.totals.worked_hours || report.totals.worked)}</b></td>${report.cols.map(col => `<td><b>${two(report.totals[col.key])}</b></td>`).join("")}</tr>`);
      return `<div class="oil-report-header"><h3>REPORTE SEMANAL CONSUMO DE ACEITES</h3><div class="oil-days"><span>Dia Inicial:<b>${Number(String(report.start).slice(-2))}</b></span><span>Dia Final:<b>${Number(String(report.end).slice(-2))}</b></span></div></div>
        <table class="oil-report-table"><thead><tr><th># Eco</th><th>Equipo</th><th>Hrs<br/>Periodo</th><th>Hrs<br/>Trab</th><th>Consumo<br/>Motor<br/>15W40</th><th>Consumo<br/>ISO 68</th><th>SAE30</th><th>SAE 50</th><th>85W140</th></tr></thead><tbody>${body.join("")}</tbody></table>`;
    }
    function oilTotalsForColumns(start, end, cols){
      const totals = {};
      cols.forEach(col => totals[col.key] = 0);
      (portal.captures || []).filter(row => inRange(row.work_date, start, end)).forEach(row => {
        cols.forEach(col => totals[col.key] += Number(row[col.key] || 0));
      });
      return totals;
    }
    function oilOrderPanelHtml(report){
      const cols = oilOrderColumns();
      const totals = oilTotalsForColumns(report.start, report.end, cols);
      report.cols.forEach(col => totals[col.key] = Number(report.totals[col.key] || totals[col.key] || 0));
      const days = Math.max(Number(report.days || 1), 1);
      const rows = cols.map(col => {
        const daily = Number(totals[col.key] || 0) / days;
        const need7 = daily * 7;
        return {
          label: col.label,
          need7,
          ped7: need7 * 1.05,
          ped15: daily * 15 * 1.05,
          ped30: daily * 30 * 1.05,
        };
      });
      const total7 = rows.reduce((sum,row) => sum + row.ped7, 0);
      const total15 = rows.reduce((sum,row) => sum + row.ped15, 0);
      const total30 = rows.reduce((sum,row) => sum + row.ped30, 0);
      return `<div class="oil-order-panel">
        <div class="oil-order-title">PEDIDO DE LUBRICANTES</div>
        <div class="oil-order-kpis">
          <div class="oil-order-kpi"><strong>${one(total7)} L</strong><span>Pedido 7d</span></div>
          <div class="oil-order-kpi"><strong>${one(total15)} L</strong><span>Pedido 15d</span></div>
          <div class="oil-order-kpi"><strong>${one(total30)} L</strong><span>Pedido 30d</span></div>
        </div>
        <table class="oil-order-table"><thead><tr><th>Lubricante</th><th>Nec. 7d</th><th>Ped. 7d</th><th>Ped. 15d</th><th>Ped. 30d</th></tr></thead><tbody>
          ${rows.map(row => `<tr><td>${esc(row.label)}</td><td>${one(row.need7)}</td><td>${one(row.ped7)}</td><td>${one(row.ped15)}</td><td class="oil-order-hot">${one(row.ped30)}</td></tr>`).join("")}
        </tbody></table>
      </div>`;
    }
    function renderOilDashboard(){
      setDashboardMode("oil");
      const report = oilRowsForPeriod();
      const accStart = `${String(report.end || report.start).slice(0,4)}-01-01`;
      const accumulated = oilRowsForRange(accStart, report.end, report.cols);
      const month = oilMonthLabel(report.end || report.start);
      $("portalUpdated").textContent = "";
      $("kpiTitle").innerHTML = `<span class="oil-month">${esc(month)}</span><span>Consumo de aceite de equipos "Providencia"</span><span class="oil-month">${esc(month)}</span>`;
      $("kpiCards").innerHTML = [
        oilMetricSection("Consumo de Aceite HCO", report.totals.oil_hco_iso68, accumulated.totals.oil_hco_iso68, "teal"),
        oilMetricSection("Consumo de Aceite SAE 30", report.totals.oil_trans_sae30, accumulated.totals.oil_trans_sae30, "red"),
      ].join("");
      $("kpiSideCards").innerHTML = [
        oilMetricSection("Consumo de Aceite de Motor", report.totals.oil_motor_15w40, accumulated.totals.oil_motor_15w40, "red"),
        oilMetricSection("Consumo de Aceite SAE 50", report.totals.oil_sae50, accumulated.totals.oil_sae50, "red"),
      ].join("");
      $("kpiChart").innerHTML = oilChartHtml(report);
      const tableWrap = $("kpiTable").closest(".table-wrap");
      if(tableWrap) tableWrap.classList.add("oil-bottom-wrap");
      $("kpiTable").className = "oil-bottom-grid";
      $("kpiTable").innerHTML = `<tbody><tr><td class="oil-report-cell">${oilReportTableHtml(report)}</td><td class="oil-order-cell">${oilOrderPanelHtml(report)}</td></tr></tbody>`;
    }
    function renderTireDashboard(){
      setDashboardMode("special");
      const tire = portal.tire_kpi || {};
      const rows = Array.isArray(tire.rows) ? tire.rows : [];
      const summary = tire.summary || {};
      $("portalUpdated").textContent = portal.updated_at || portal.generated_at ? `Actualizado ${portal.updated_at || portal.generated_at}` : "Sin sincronizar";
      $("kpiTitle").textContent = "KPI Llantas";
      $("kpiCards").innerHTML = [
        ["Llantas", `${summary.total || rows.length || 0}`, "registradas", 100, false],
        ["Vida prom.", pct(summary.avg_life || 0), "igual a % piso", Number(summary.avg_life || 0), false],
        ["Criticas", `${summary.critical || 0}`, "cambio requerido", Number(summary.critical || 0) ? 100 : 0, Number(summary.critical || 0) > 0],
        ["Proximas", `${summary.soon || 0}`, "seguimiento", Number(summary.soon || 0) ? 70 : 0, false],
      ].map(([label, value, note, width, bad]) => metricCardHtml(label, value, note, width, bad)).join("");
      const chartRows = [...rows].sort((a,b) => Number(a.life_percent || 0) - Number(b.life_percent || 0)).slice(0,24);
      $("kpiChart").innerHTML = chartRows.map(row => {
        const value = Math.max(Math.min(Number(row.life_percent || row.tread_remaining_percent || 0), 100), 0);
        const out = ["CRITICA","BAJA"].includes(String(row.control_status || "").toUpperCase());
        return `<div class="chart-bar ${out ? "out" : ""}" title="${esc(row.equipment_code)} ${esc(row.tire_code)} ${pct(value)}"><span>${pct(value)}</span><i style="--h:${Math.max(value * 2.1, 4)}px"></i><b>${esc(row.equipment_code || "-")}</b></div>`;
      }).join("") || `<p class="muted">Sin llantas registradas.</p>`;
      $("kpiTable").innerHTML = `<thead><tr><th>Equipo</th><th>Llanta</th><th>Pos.</th><th>Marca</th><th>Hrs uso</th><th>Hrs rest.</th><th>% Vida</th><th>% Piso</th><th>KPI</th><th>Recomendacion</th></tr></thead><tbody>` +
        rows.map(row => {
          const value = Number(row.life_percent || row.tread_remaining_percent || 0);
          const status = String(row.control_status || "");
          const cls = status === "OK" ? "ok" : (status === "PROXIMA" || status === "REVISION" ? "warn" : "bad");
          return `<tr><td>${esc(row.equipment_code)}</td><td>${esc(row.tire_code)}</td><td>${esc(row.position)}</td><td>${esc(row.brand)}</td><td>${one(row.hours_used)}</td><td>${one(row.life_remaining_hours)}</td><td>${pct(value)}</td><td>${pct(value)}</td><td><span class="pill ${cls}">${esc(status || "S/D")}</span></td><td>${esc(row.recommendation || "")}</td></tr>`;
        }).join("") + `</tbody>`;
    }
    function tireRows(){
      const tire = portal.tire_kpi || {};
      return Array.isArray(tire.rows) ? tire.rows : [];
    }
    function tireEvents(){
      const tracking = portal.tire_tracking || {};
      return Array.isArray(tracking.events) ? tracking.events : [];
    }
    function tireStatusClass(status){
      const text = normalizedText(status);
      if(text === "OK" || text === "ALMACEN") return "ok";
      if(text === "PROXIMA" || text === "REVISION") return "warn";
      return "bad";
    }
    function tireCodeLabel(row){
      const code = row.tire_code || "";
      const details = [row.equipment_code, row.position].filter(Boolean).join(" ");
      return details ? `${code} - ${details}` : code;
    }
    function tireTrackEquipmentOptions(){
      const options = portalEquipment().map(e => ({value:e.code || e.equipment_code, label:`${e.code || e.equipment_code} - ${e.description || e.family || ""}`}));
      const seen = new Set(options.map(item => item.value));
      tireRows().forEach(row => {
        const code = row.equipment_code || "";
        if(code && !seen.has(code)){
          seen.add(code);
          options.push({value:code, label:code});
        }
      });
      return options;
    }
    function resetTireTrackForm(){
      $("tireTrackDate").value = toIsoDate(new Date());
      $("tireTrackType").value = "INSPECCION";
      $("tireTrackCode").value = "";
      $("tireTrackEquipment").value = "";
      $("tireTrackPosition").value = "";
      $("tireTrackMountStatus").value = "MONTADA";
      ["tireTrackBrand","tireTrackModel","tireTrackSize","tireTrackTechnician","tireTrackNotes"].forEach(id => $(id).value = "");
      ["tireTrackInstallMeter","tireTrackCurrentMeter","tireTrackTargetHours","tireTrackTreadInitial","tireTrackTreadCurrent","tireTrackPressure"].forEach(id => $(id).value = "0");
      $("tireTrackStatus").textContent = "";
    }
    function fillTireTrackForm(row){
      if(!row) return;
      $("tireTrackCode").value = row.tire_code || "";
      $("tireTrackEquipment").value = row.equipment_code || "";
      $("tireTrackPosition").value = row.position || "";
      $("tireTrackMountStatus").value = row.status || "MONTADA";
      $("tireTrackBrand").value = row.brand || "";
      $("tireTrackModel").value = row.model || "";
      $("tireTrackSize").value = row.size || "";
      $("tireTrackInstallMeter").value = row.install_meter || 0;
      $("tireTrackCurrentMeter").value = row.current_meter || 0;
      $("tireTrackTargetHours").value = row.target_life_hours || 0;
      $("tireTrackTreadInitial").value = row.tread_initial || 0;
      $("tireTrackTreadCurrent").value = row.tread_current || 0;
      $("tireTrackPressure").value = row.pressure_current || 0;
      $("tireTrackNotes").value = row.notes || "";
      $("tireTrackStatus").textContent = `${row.control_status || "S/D"} | ${row.recommendation || ""}`;
    }
    function renderTireTracking(){
      const rows = tireRows();
      const events = tireEvents();
      const summary = (portal.tire_kpi || {}).summary || {};
      $("tireTrackCount").textContent = `${rows.length} llantas`;
      $("tireEventCount").textContent = `${events.length} evento(s)`;
      $("tireTrackSummary").innerHTML = [
        ["Total", summary.total || rows.length || 0],
        ["Criticas", summary.critical || 0],
        ["Proximas", summary.soon || 0],
        ["Vida prom.", pct(summary.avg_life || 0)],
      ].map(([label, value]) => `<span>${esc(label)}<b>${esc(value)}</b></span>`).join("");
      $("tireTrackCodes").innerHTML = rows.map(row => `<option value="${esc(row.tire_code || "")}">${esc(tireCodeLabel(row))}</option>`).join("");
      setOptions("tireTrackEquipment", tireTrackEquipmentOptions(), "Sin equipo");
      $("tireTrackTable").innerHTML = `<thead><tr><th>Equipo</th><th>Llanta</th><th>Pos.</th><th>Marca</th><th>Hor.</th><th>Piso</th><th>Vida</th><th>KPI</th><th>Accion</th></tr></thead><tbody>` +
        rows.map(row => {
          const status = row.control_status || "S/D";
          return `<tr><td>${esc(row.equipment_code || "")}</td><td>${esc(row.tire_code || "")}</td><td>${esc(row.position || "")}</td><td>${esc(row.brand || "")}</td><td>${one(row.current_meter)}</td><td>${one(row.tread_current)}</td><td>${pct(row.life_percent || 0)}</td><td><span class="pill ${tireStatusClass(status)}">${esc(status)}</span></td><td><button type="button" class="btn secondary small" data-tire-load="${esc(row.tire_code || "")}">Cargar</button></td></tr>`;
        }).join("") + `</tbody>`;
      $("tireEventTable").innerHTML = `<thead><tr><th>Fecha</th><th>Tipo</th><th>Llanta</th><th>Equipo</th><th>Pos.</th><th>Hor.</th><th>Piso</th><th>PSI</th><th>KPI</th><th>Notas</th><th>Accion</th></tr></thead><tbody>` +
        events.map(event => `<tr><td>${esc(event.event_date || "")}</td><td>${esc(event.event_type || "")}</td><td>${esc(event.tire_code || "")}</td><td>${esc(event.equipment_code || "")}</td><td>${esc(event.position || "")}</td><td>${one(event.meter)}</td><td>${one(event.tread_mm)}</td><td>${one(event.pressure_psi)}</td><td><span class="pill ${tireStatusClass(event.control_status)}">${esc(event.control_status || "")}</span></td><td>${esc(shortText(event.notes || "", 90))}</td><td><button type="button" class="btn danger small" data-tire-event-delete="${esc(event.id || "")}">Eliminar</button></td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-tire-load]").forEach(button => button.addEventListener("click", () => {
        const row = tireRows().find(item => String(item.tire_code || "") === String(button.dataset.tireLoad || ""));
        fillTireTrackForm(row);
      }));
      document.querySelectorAll("[data-tire-event-delete]").forEach(button => button.addEventListener("click", () => deleteTireTrackEvent(button.dataset.tireEventDelete).catch(showError)));
    }
    async function saveTireTrackEvent(){
      if(!hasApiKey()) return;
      const payload = {
        event_date:$("tireTrackDate").value,
        event_type:$("tireTrackType").value,
        tire_code:$("tireTrackCode").value,
        equipment_code:$("tireTrackEquipment").value,
        position:$("tireTrackPosition").value,
        status:$("tireTrackMountStatus").value,
        brand:$("tireTrackBrand").value,
        model:$("tireTrackModel").value,
        size:$("tireTrackSize").value,
        install_meter:$("tireTrackInstallMeter").value,
        current_meter:$("tireTrackCurrentMeter").value,
        target_life_hours:$("tireTrackTargetHours").value,
        tread_initial:$("tireTrackTreadInitial").value,
        tread_current:$("tireTrackTreadCurrent").value,
        pressure_current:$("tireTrackPressure").value,
        technician:$("tireTrackTechnician").value,
        event_notes:$("tireTrackNotes").value,
        notes:$("tireTrackNotes").value,
      };
      const response = await fetch("/api/tire-tracking/events", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      portal = result.portal || portal;
      $("tireTrackStatus").textContent = "Seguimiento guardado y KPI actualizado.";
      renderPortalSelectors();
      renderTireTracking();
      renderDashboard();
    }
    async function deleteTireTrackEvent(id){
      if(!hasApiKey()) return;
      if(!confirm("Eliminar este evento de llanta?")) return;
      const response = await fetch("/api/tire-tracking/events/delete", {method:"POST", headers:headers(true), body:JSON.stringify({id})});
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      portal = result.portal || portal;
      $("tireTrackStatus").textContent = "Evento eliminado.";
      renderTireTracking();
      renderDashboard();
    }
    function dieselCaptureHoursFromBitacora(capture){
      const hi = Number(capture?.hi || capture?.horometer_initial || 0);
      const hf = Number(capture?.hf || capture?.horometer_final || 0);
      const worked = Number(capture?.worked_hours || 0);
      if(worked > 0) return worked;
      if(hi > 0 && hf >= hi) return hf - hi;
      return 0;
    }
    function dieselBitacoraHoursByEquipment(start, end, aliasMap=dieselBaseEquipmentAliasMap()){
      const grouped = new Map();
      (portal.captures || []).filter(capture => inRange(capture.work_date, start, end)).forEach(capture => {
        const equipment = dieselCanonicalEquipment(capture.equipment_code || capture.equipment || capture.code || capture.eco || "", aliasMap);
        if(!equipment) return;
        addDieselEquipmentAlias(aliasMap, equipment, [capture.equipment_code, capture.equipment, capture.code, capture.eco]);
        if(!grouped.has(equipment)) grouped.set(equipment, {hi_values:[], hf_values:[], worked_hours:0});
        const row = grouped.get(equipment);
        const hi = Number(capture.hi || capture.horometer_initial || 0);
        const hf = Number(capture.hf || capture.horometer_final || 0);
        if(hi > 0) row.hi_values.push(hi);
        if(hf > 0) row.hf_values.push(hf);
        row.worked_hours += dieselCaptureHoursFromBitacora(capture);
      });
      return grouped;
    }
    function dieselApplyBitacoraHours(row, bitacoraHours){
      const bitacora = bitacoraHours.get(row.equipment);
      const bitacoraWorked = bitacora ? Number(bitacora.worked_hours || 0) : 0;
      const hiValues = bitacora && bitacora.hi_values.length ? bitacora.hi_values : (row.hi_values || row.hi || []);
      const hfValues = bitacora && bitacora.hf_values.length ? bitacora.hf_values : (row.hf_values || row.hf || []);
      return {
        ...row,
        horometer_initial: hiValues.length ? Math.min(...hiValues) : 0,
        horometer_final: hfValues.length ? Math.max(...hfValues) : 0,
        worked_hours: bitacoraWorked > 0 ? bitacoraWorked : Number(row.worked_hours || 0),
        hours_source: bitacoraWorked > 0 ? "bitacora" : (row.hours_source || "diesel"),
      };
    }
    function dieselKpiRowsForPeriod(){
      const start = $("kpiStart").value;
      const end = $("kpiEnd").value;
      const meta = Number($("dieselMeta").value || diesel.meta_lh || (portal.settings || {}).meta_diesel_lh || 25);
      const grouped = {};
      const aliasMap = dieselBaseEquipmentAliasMap();
      const bitacoraHours = dieselBitacoraHoursByEquipment(start, end, aliasMap);
      const periodRecords = (diesel.records || []).filter(row => inRange(row.work_date, start, end));
      periodRecords.forEach(record => {
        const equipment = dieselCanonicalEquipment(record.equipment, aliasMap);
        if(!equipment) return;
        addDieselEquipmentAlias(aliasMap, equipment, [record.equipment]);
        if(!grouped[equipment]){
          grouped[equipment] = {equipment, condition:record.condition || "DISPONIBLE", hi:[], hf:[], worked_hours:0, diesel_liters:0};
        }
        const row = grouped[equipment];
        row.condition = record.condition || row.condition;
        const hi = Number(record.horometer_initial || 0);
        const hf = Number(record.horometer_final || 0);
        if(hi > 0) row.hi.push(hi);
        if(hf > 0) row.hf.push(hf);
        row.worked_hours += Number(record.worked_hours || 0);
        row.diesel_liters += Number(record.diesel_liters || 0);
      });
      const rows = Object.values(grouped).map(row => {
        const merged = dieselApplyBitacoraHours(row, bitacoraHours);
        const rendimiento = merged.worked_hours > 0 ? merged.diesel_liters / merged.worked_hours : null;
        let status = "OK";
        if(merged.diesel_liters <= 0) status = "SIN CONSUMO";
        else if(merged.worked_hours <= 0) status = "SIN HORAS";
        else if(rendimiento !== null && rendimiento > meta) status = "ALTO";
        return {
          ...merged,
          rendimiento_lh: rendimiento,
          status,
        };
      }).sort((a,b) => b.diesel_liters - a.diesel_liters || a.equipment.localeCompare(b.equipment));
      const totals = rows.reduce((acc, row) => {
        acc.diesel_liters += Number(row.diesel_liters || 0);
        acc.worked_hours += Number(row.worked_hours || 0);
        if(["ALTO","SIN HORAS"].includes(row.status)) acc.critical += 1;
        return acc;
      }, {diesel_liters:0, worked_hours:0, critical:0});
      totals.rendimiento_lh = totals.worked_hours > 0 ? totals.diesel_liters / totals.worked_hours : null;
      Object.assign(totals, dieselSupplierTotals(periodRecords));
      Object.assign(totals, dieselInventoryTotals(start, end));
      return {start, end, meta, rows, totals};
    }
    function dieselPctValue(value){ return Math.max(Math.min(Number(value || 0), 100), 0); }
    function dieselPeriodDays(report){
      const start = parseIsoDate(report.start);
      const end = parseIsoDate(report.end);
      return Math.max(Math.round((end - start) / 86400000) + 1, 1);
    }
    function dieselKpiCardHtml(label, value, note, width, bad=false){
      return `<article class="diesel-card ${bad ? "is-bad" : ""}"><span>${esc(label)}</span><strong>${esc(value)}</strong><small>${esc(note)}</small><div class="diesel-meter"><i style="width:${dieselPctValue(width)}%"></i></div></article>`;
    }
    function dieselCardsHtml(report){
      const totals = report.totals || {};
      const avg = totals.rendimiento_lh;
      const liters = Number(totals.diesel_liters || 0);
      const totalStock = Number(totals.total_stock || 0);
      const mga = Number(totals.mga_stock || 0);
      const prosermin = Number(totals.prosermin_stock || 0);
      const supplierTotal = Math.max(totalStock || (mga + prosermin), 1);
      const days = dieselPeriodDays(report);
      return [
        dieselKpiCardHtml("Equipos", `${report.rows.length}`, "con captura diesel", Math.min((report.rows.length / 18) * 100, 100)),
        dieselKpiCardHtml("Consumo total", `${one(liters)} L`, `${days} dias analizados`, liters > 0 ? 100 : 0),
        dieselKpiCardHtml("Diesel MGA", `${one(mga)} L`, "disponible MGA", (mga / supplierTotal) * 100),
        dieselKpiCardHtml("Diesel PROSERMIN", `${one(prosermin)} L`, "disponible PROSERMIN", (prosermin / supplierTotal) * 100),
        dieselKpiCardHtml("Horas trabajadas", `${one(totals.worked_hours)} h`, "horas de bitacora", Math.min(Number(totals.worked_hours || 0) / Math.max(report.rows.length * 10, 1) * 100, 100)),
        dieselKpiCardHtml("Rendimiento", avg == null ? "S/H" : `${one(avg)} L/H`, `Meta ${one(report.meta)} L/H`, avg == null ? 0 : Math.min((avg / Math.max(report.meta, 1)) * 100, 100), avg != null && avg > report.meta),
      ].join("");
    }
    function dieselConsumptionPanelHtml(report){
      const chartRows = report.rows.filter(row => Number(row.diesel_liters || 0) > 0 || Number(row.worked_hours || 0) > 0).slice(0, 12);
      const maxLiters = Math.max(1, ...chartRows.map(row => Number(row.diesel_liters || 0)));
      const body = chartRows.map(row => {
        const liters = Number(row.diesel_liters || 0);
        const width = Math.max((liters / maxLiters) * 100, 2);
        const status = dieselStatusClass(row.status);
        const rend = row.rendimiento_lh == null ? "S/H" : `${one(row.rendimiento_lh)} L/H`;
        return `<div class="diesel-bar-row ${status === "bad" ? "is-bad" : ""}" title="${esc(row.equipment)} ${one(liters)} L | ${esc(rend)}"><div class="diesel-bar-label"><b>${esc(row.equipment)}</b><span>${esc(row.status)}</span></div><div class="diesel-bar-track"><i style="width:${dieselPctValue(width)}%"></i></div><strong>${one(liters)} L</strong><em>${esc(rend)}</em></div>`;
      }).join("");
      return `<section class="diesel-panel"><div class="diesel-panel-head"><span>Consumo por equipo</span><b>Top ${chartRows.length || 0}</b></div>${body ? `<div class="diesel-bars-list">${body}</div>` : `<div class="diesel-empty">Sin capturas diesel en el periodo.</div>`}</section>`;
    }
    function dieselPerformancePanelHtml(report){
      const totals = report.totals || {};
      const avg = totals.rendimiento_lh;
      const maxScale = Math.max(report.meta * 1.6, 1);
      const avgPct = avg == null ? 0 : dieselPctValue((avg / maxScale) * 100);
      const targetPct = dieselPctValue((report.meta / maxScale) * 100);
      const totalStock = Number(totals.total_stock || 0);
      const mga = Number(totals.mga_stock || 0);
      const prosermin = Number(totals.prosermin_stock || 0);
      const supplierTotal = Math.max(totalStock || (mga + prosermin), 1);
      const mgaPct = (mga / supplierTotal) * 100;
      const proPct = (prosermin / supplierTotal) * 100;
      const watch = report.rows.filter(row => ["ALTO","SIN HORAS"].includes(String(row.status || "").toUpperCase())).slice(0, 4);
      const watchHtml = watch.length
        ? watch.map(row => `<div class="diesel-watch-item"><b>${esc(row.equipment)}</b><span>${esc(row.status)} | ${row.rendimiento_lh == null ? "S/H" : `${one(row.rendimiento_lh)} L/H`}</span></div>`).join("")
        : `<div class="diesel-watch-item"><b>Sin alertas</b><span>Dentro de meta</span></div>`;
      return `<section class="diesel-panel diesel-performance-grid">
        <div class="diesel-target">
          <div class="diesel-target-top"><div><span>Rendimiento promedio</span><strong>${avg == null ? "S/H" : `${one(avg)} L/H`}</strong></div><small>Meta ${one(report.meta)} L/H</small></div>
          <div class="diesel-target-meter"><i style="width:${avgPct}%"></i><span class="diesel-target-marker" style="--target:${targetPct}%"></span></div>
        </div>
        <div class="diesel-split">
          <h4>Existencia disponible</h4>
          <div class="diesel-split-track"><i class="mga" style="width:${dieselPctValue(mgaPct)}%"></i><i class="pro" style="width:${dieselPctValue(proPct)}%"></i></div>
          <div class="diesel-split-legend"><span><em><i class="mga"></i>MGA</em><b>${one(mga)} L</b></span><span><em><i class="pro"></i>PROSERMIN</em><b>${one(prosermin)} L</b></span></div>
        </div>
        <div class="diesel-watch-list">${watchHtml}</div>
      </section>`;
    }
    function dieselVisualHtml(report){
      return `<div class="diesel-visual-grid">${dieselConsumptionPanelHtml(report)}${dieselPerformancePanelHtml(report)}</div>`;
    }
    function dieselTableHtml(report){
      const body = report.rows.map(row => {
        const cls = dieselStatusClass(row.status);
        return `<tr><td class="diesel-eq">${esc(row.equipment)}</td><td>${esc(row.condition)}</td><td class="diesel-number">${one(row.horometer_initial)}</td><td class="diesel-number">${one(row.horometer_final)}</td><td class="diesel-number">${one(row.worked_hours)}</td><td class="diesel-number">${one(row.diesel_liters)}</td><td class="diesel-number">${row.rendimiento_lh == null ? "S/H" : one(row.rendimiento_lh)}</td><td class="diesel-number">${one(report.meta)}</td><td><span class="pill ${cls}">${esc(row.status)}</span></td></tr>`;
      }).join("") || `<tr><td colspan="9">Sin capturas diesel en el periodo.</td></tr>`;
      return `<thead><tr><th>Equipo</th><th>Condicion</th><th>HI</th><th>HF</th><th>Hrs Trab</th><th>Diesel L</th><th>Rend. L/H</th><th>Meta</th><th>KPI</th></tr></thead><tbody>${body}<tr class="diesel-total"><td>Total</td><td></td><td></td><td></td><td class="diesel-number">${one(report.totals.worked_hours)}</td><td class="diesel-number">${one(report.totals.diesel_liters)}</td><td class="diesel-number">${report.totals.rendimiento_lh == null ? "S/H" : one(report.totals.rendimiento_lh)}</td><td class="diesel-number">${one(report.meta)}</td><td>${report.totals.critical} revision</td></tr></tbody>`;
    }
    function renderDieselDashboard(){
      setDashboardMode("diesel");
      const report = dieselKpiRowsForPeriod();
      $("portalUpdated").textContent = diesel.updated_at ? `Actualizado ${diesel.updated_at}` : (portal.updated_at || portal.generated_at ? `Actualizado ${portal.updated_at || portal.generated_at}` : "Sin sincronizar");
      $("kpiTitle").textContent = `KPI Diesel | ${report.start} a ${report.end}`;
      $("kpiCards").innerHTML = dieselCardsHtml(report);
      $("kpiChart").innerHTML = dieselVisualHtml(report);
      $("kpiTable").className = "diesel-table";
      $("kpiTable").innerHTML = dieselTableHtml(report);
    }
    function renderDashboard(){
      const selectedGroup = $("kpiGroup").value || "";
      const commandReport = simulatedKpiReport(calculateKpiRows("Todos los equipos", $("kpiStart").value, $("kpiEnd").value));
      renderKpiCommandCenter(commandReport);
      renderKpiMainSummary(commandReport, currentKpiSettings());
      if(selectedGroup === "KPI Aceites") {
        renderOilDashboard();
        return;
      }
      if(selectedGroup === "KPI Llantas") {
        renderTireDashboard();
        return;
      }
      if(selectedGroup === "KPI Diesel") {
        renderDieselDashboard();
        return;
      }
      setDashboardMode("format");
      const report = simulatedKpiReport(calculateKpiRows());
      renderKpiCommandCenter(report);
      const settings = currentKpiSettings();
      renderKpiMainSummary(report, settings);
      $("kpiPrintArea").classList.toggle("simulation", Boolean(report.simulation?.enabled));
      $("portalUpdated").textContent = portal.updated_at || portal.generated_at ? `Actualizado ${portal.updated_at || portal.generated_at}` : "Sin sincronizar";
      $("kpiTitle").innerHTML = `${esc(report.group)} | ${esc(report.start)} a ${esc(report.end)}${report.simulation?.enabled ? `<span class="kpi-simulation-badge">SIMULACION: ${esc(report.simulation.name || "Escenario KPI")}</span>` : ""}`;
      const metaAvailability = Number(settings.meta_availability || 85);
      const metaUtilization = Number(settings.meta_utilization || 75);
      const metaReliability = Number(settings.meta_reliability || 80);
      const metaTmef = Number(settings.meta_tmef || 8);
      const metaTmpr = Number(settings.meta_tmpr || 4);
      $("kpiCards").innerHTML = [
        metricCardHtml("% Disponibilidad", pct(report.totals.availability), `Meta ${pct(metaAvailability)}`, report.totals.availability, report.totals.availability < metaAvailability),
        metricCardHtml("Meta", pct(metaAvailability), `${one(report.totals.availability - metaAvailability)}%`, metaAvailability, false),
        metricCardHtml("% Utilizacion", pct(report.totals.utilization), `Meta ${pct(metaUtilization)}`, report.totals.utilization, report.totals.utilization < metaUtilization),
        metricCardHtml("Meta", pct(metaUtilization), `${one(report.totals.utilization - metaUtilization)}%`, metaUtilization, report.totals.utilization < metaUtilization),
      ].join("");
      $("kpiSideCards").innerHTML = [
        metricCardHtml("Confiabilidad", pct(report.totals.reliability), `Meta ${pct(metaReliability)}`, report.totals.reliability, report.totals.reliability < metaReliability),
        metricCardHtml("TMEF", `${one(report.totals.tmef)} h`, `Meta ${one(metaTmef)} h`, Math.min((report.totals.tmef / Math.max(metaTmef, 1)) * 100, 100), report.totals.tmef < metaTmef),
        metricCardHtml("TMPR", `${one(report.totals.tmpr)} h`, `Meta ${one(metaTmpr)} h`, Math.min((report.totals.tmpr / Math.max(metaTmpr, 1)) * 100, 100), report.totals.tmpr > metaTmpr),
        metricCardHtml("Meta Conf.", pct(metaReliability), `${one(report.totals.reliability - metaReliability)}%`, metaReliability, report.totals.reliability < metaReliability),
      ].join("");
      $("kpiChart").innerHTML = kpiMetricChartHtml(report, settings);
      bindKpiMetricTabs();
      $("kpiTable").innerHTML = `<thead><tr><th># Eco</th><th>Equipo</th><th>Hrs periodo</th><th>Hrs MP</th><th>Hrs MC</th><th>Hrs trab</th><th># Paradas</th><th>% Disp</th><th>% Util</th><th>Confiabilidad</th><th>TMEF</th><th>TMPR</th><th>Estatus</th></tr></thead><tbody>` +
        report.rows.map(row => `<tr><td>${esc(row.code)}</td><td>${esc(row.description)}</td><td>${one(row.period)}</td><td>${one(row.mp)}</td><td>${one(row.mc)}</td><td>${one(row.worked)}</td><td>${num(row.stops)}</td><td>${esc(row.availabilityText)}</td><td>${esc(row.utilizationText)}</td><td>${pct(row.reliability)}</td><td>${one(row.tmef)}</td><td>${one(row.tmpr)}</td><td>${esc(row.out ? "FUERA" : row.status)}</td></tr>`).join("") +
        `<tr><td></td><td><b>Total ${esc(report.group)}</b></td><td><b>${one(report.totals.period)}</b></td><td><b>${one(report.totals.mp)}</b></td><td><b>${one(report.totals.mc)}</b></td><td><b>${one(report.totals.worked)}</b></td><td><b>${num(report.totals.stops)}</b></td><td><b>${pct(report.totals.availability)}</b></td><td><b>${pct(report.totals.utilization)}</b></td><td><b>${pct(report.totals.reliability)}</b></td><td><b>${one(report.totals.tmef)}</b></td><td><b>${one(report.totals.tmpr)}</b></td><td></td></tr></tbody>`;
    }
    const preventiveServiceHours = {PM1:250, PM2:500, PM3:750, PM4:1000};
    const preventiveClosedStates = new Set(["CERRADO","CERRADA","TERMINADO","TERMINADA","FINALIZADO","FINALIZADA"]);
    const preventiveOilInputs = [
      ["prevExecOil15w40", "oil_motor_15w40", "15W40"],
      ["prevExecOilHco68", "oil_hco_iso68", "HCO ISO 68"],
      ["prevExecOilSae30", "oil_trans_sae30", "SAE 30"],
      ["prevExecOil85w140", "oil_85w140", "85W140"],
      ["prevExecAlmo", "almo_liters", "ALMO"],
      ["prevExecCoolant", "coolant_liters", "Refrigerante"],
      ["prevExecOilVg100", "oil_hyd_vg100", "VG100"],
      ["prevExecAtf", "atf_liters", "ATF"],
    ];
    const preventiveChecklistInputs = [
      ["prevChkInspection", "inspection", "Inspeccion"],
      ["prevChkFilters", "filters", "Filtros/refacciones"],
      ["prevChkLubrication", "lubrication", "Lubricacion"],
      ["prevChkElectrical", "electrical", "Revision electrica"],
      ["prevChkTest", "test", "Prueba final"],
      ["prevChkSupervisor", "supervisor", "Supervisor"],
    ];
    function preventiveExecutionRows(){
      const payload = portal.preventive_execution || {};
      return Array.isArray(payload.records) ? payload.records : [];
    }
    function isPreventiveClosed(row){
      return preventiveClosedStates.has(String(row.status || "").toUpperCase());
    }
    function preventiveIntervalHours(value){
      const text = String(value || "").toUpperCase().replace(/\s+/g, "");
      if(preventiveServiceHours[text]) return preventiveServiceHours[text];
      const match = text.match(/(\d+)/);
      return match ? Number(match[1]) : 0;
    }
    function preventiveRowsWithWebClosures(rows){
      const closed = preventiveExecutionRows().filter(isPreventiveClosed);
      if(!closed.length) return rows;
      return rows.map(row => {
        const code = row.equipment_code || "";
        const service = String(row.service_interval || row.service_name || "").toUpperCase().replace(/\s+/g, "");
        const interval = preventiveIntervalHours(service || row.service_interval || row.service_name);
        const match = closed
          .filter(item => {
            const itemService = String(item.service_type || "").toUpperCase().replace(/\s+/g, "");
            return item.equipment_code === code && (!service || itemService === service || preventiveServiceHours[itemService] === interval);
          })
          .sort((a,b) => String(b.close_date || b.service_date || "").localeCompare(String(a.close_date || a.service_date || "")) || Number(b.completed_meter || 0) - Number(a.completed_meter || 0))[0];
        if(!match || !Number(match.completed_meter || 0)) return row;
        const current = Number(row.current_meter || 0);
        const next = Number(match.completed_meter || 0) + (interval || preventiveServiceHours[String(match.service_type || "").toUpperCase()] || 0);
        if(!next || next <= Number(row.next_service_meter || 0)) return row;
        const updated = {...row, last_service_meter:Number(match.completed_meter || 0), next_service_meter:next, hours_remaining:next-current, last_web_service_date:match.close_date || match.service_date || ""};
        if(updated.hours_remaining < 0) updated.status = "VENCIDO";
        else if(updated.hours_remaining <= Math.max((interval || 250) * 0.1, 25)) updated.status = "URGENTE";
        else if(updated.hours_remaining <= Math.max((interval || 250) * 0.25, 50)) updated.status = "PROXIMO";
        else updated.status = "PROGRAMADO";
        return updated;
      });
    }
    function preventiveOilTotal(row){
      return preventiveOilInputs.reduce((sum, item) => sum + Number(row?.[item[1]] || 0), 0);
    }
    function preventiveOilsText(row){
      return preventiveOilInputs
        .map(item => ({label:item[2], value:Number(row?.[item[1]] || 0)}))
        .filter(item => item.value > 0)
        .map(item => `${item.label}: ${one(item.value)} L`)
        .join("; ");
    }
    function preventiveChecklistPayload(){
      const checklist = {};
      preventiveChecklistInputs.forEach(item => { checklist[item[1]] = Boolean($(item[0]).checked); });
      return checklist;
    }
    function preventiveChecklistCount(row){
      const checklist = row?.checklist || {};
      return preventiveChecklistInputs.filter(item => Boolean(checklist[item[1]])).length;
    }
    function filteredPreventives(){
      const [start, end] = periodRange($("prPeriod").value, $("prBase").value);
      const selected = $("prEquipment").value;
      const search = ($("prSearch").value || "").toUpperCase();
      const rows = preventiveRowsWithWebClosures(portal.preventives || []).filter(row => {
        const dateOk = inRange(row.projected_date, start, end) || ["VENCIDO", "URGENTE"].includes(String(row.status || "").toUpperCase());
        const eqOk = !selected || row.equipment_code === selected;
        const text = [row.equipment_code,row.equipment_description,row.component,row.meter_type,row.status].join(" ").toUpperCase();
        return dateOk && eqOk && (!search || text.includes(search));
      }).sort((a,b) => String(a.projected_date || "").localeCompare(String(b.projected_date || "")) || String(a.equipment_code || "").localeCompare(String(b.equipment_code || "")));
      return {start, end, rows};
    }
    function renderPreventives(){
      const result = filteredPreventives();
      $("prTitle").textContent = `PR Preventivos | ${result.start} a ${result.end}`;
      $("prCount").textContent = `${result.rows.length} preventivo(s)`;
      const period = $("prPeriod").value.toLowerCase();
      if(period.startsWith("a")){
        const months = Array.from({length:12}, (_, idx) => {
          const month = idx + 1;
          const count = result.rows.filter(r => String(r.projected_date || "").slice(5,7) === String(month).padStart(2,"0")).length;
          return `<div class="schedule-cell"><strong>${["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"][idx]}</strong><em>${count} servicios</em></div>`;
        });
        $("prCalendar").innerHTML = months.join("");
      } else {
        $("prCalendar").innerHTML = dateList(result.start, result.end, 35).map(day => {
          const chips = result.rows.filter(r => r.projected_date === day).slice(0,3).map(r => `<span class="schedule-chip ${r.status === "VENCIDO" ? "late" : (r.status === "URGENTE" ? "near" : "")}">${esc(r.equipment_code)} ${esc(r.component)}</span>`).join("");
          return `<div class="schedule-cell"><strong>${esc(day.slice(8,10))}</strong><em>${esc(day.slice(5,7))}</em>${chips}</div>`;
        }).join("");
      }
      $("prTable").innerHTML = `<thead><tr><th>Equipo</th><th>Descripcion</th><th>Componente</th><th>Tipo hor.</th><th>Horometro</th><th>Ultimo serv.</th><th>Prox. serv.</th><th>Hrs restantes</th><th>Fecha prog.</th><th>Estado</th></tr></thead><tbody>` +
        result.rows.map(row => `<tr><td>${esc(row.equipment_code)}</td><td>${esc(row.equipment_description)}</td><td>${esc(row.component)}</td><td>${esc(row.meter_type)}</td><td>${one(row.current_meter)}</td><td>${one(row.last_service_meter)}</td><td>${one(row.next_service_meter)}</td><td>${one(row.hours_remaining)}</td><td>${esc(row.projected_date || "")}</td><td><span class="pill ${row.status === "PROGRAMADO" ? "ok" : (row.status === "PROXIMO" ? "warn" : "bad")}">${esc(row.status)}</span></td></tr>`).join("") +
        `</tbody>`;
    }
    function calculateBacklogRows(start, end){
      const rows = [];
      const addRow = (row) => rows.push({...row, level: row.level || backlogLevel(Number(row.score || 0))});
      const published = (portal.backlog && Array.isArray(portal.backlog.items)) ? portal.backlog.items : [];
      if(published.some(row => row.key || row.flow_status || row.responsible || row.work_order_id)){
        return published
          .map(row => ({
            ...row,
            level: row.level || backlogLevel(Number(row.score || 0)),
            flow_status: row.flow_status || "Pendiente",
          }))
          .sort((a,b) => {
            const rank = {Pendiente:0, "En proceso":1, Atendido:2, Cancelado:3};
            return (rank[a.flow_status || "Pendiente"] ?? 0) - (rank[b.flow_status || "Pendiente"] ?? 0)
              || Number(b.score || 0) - Number(a.score || 0)
              || String(a.date || "9999-12-31").localeCompare(String(b.date || "9999-12-31"));
          });
      }
      (portal.preventives || []).forEach(row => {
        const status = String(row.status || "").toUpperCase();
        const projected = row.projected_date || "";
        const inPeriod = inRange(projected, start, end);
        if(!["VENCIDO", "URGENTE", "PROXIMO"].includes(status) && !inPeriod) return;
        const remaining = Number(row.hours_remaining || 0);
        let score = status === "VENCIDO" ? 92 + Math.min(Math.abs(remaining) / Math.max(Number(row.service_interval || 250), 1) * 18, 18) : status === "URGENTE" ? 76 : status === "PROXIMO" ? 50 : 35;
        addRow({
          score, source:"Preventivo", equipment_code:row.equipment_code || "", equipment_description:row.equipment_description || "",
          component:row.component || "", system:"Preventivo", date:projected, due_date:projected, hours_remaining:remaining,
          detail:`${row.service_interval || "PM"} ${status}: faltan ${one(remaining)} h.`,
          action: status === "VENCIDO" ? "Ejecutar preventivo y registrar horometro real." : "Programar en plan semanal y preparar filtros/refacciones.",
        });
      });
      (portal.captures || []).filter(row => inRange(row.work_date, start, end)).forEach(row => {
        const stops = Number(row.stops || 0);
        const mc = Number(row.mc_hours || 0);
        const fault = String(row.fault || "").trim();
        const wear = String(row.wear || "").trim();
        const noDisp = unavailable(row.status);
        if(!(stops || mc > 0 || fault || wear || noDisp)) return;
        let score = 26 + stops * 10 + mc * 5 + (fault ? 18 : 0) + (wear ? 9 : 0) + (noDisp ? 25 : 0);
        const system = classifyStoppage(row.component, row.fault, row.wear, row.observations, row.status);
        const detail = [
          stops ? `${stops} parada(s)` : "",
          mc ? `${one(mc)} h MC` : "",
          noDisp ? row.status : "",
          fault ? `Falla: ${fault}` : "",
          wear ? `Desgaste: ${wear}` : "",
        ].filter(Boolean).join(" | ");
        addRow({
          score, source:"Captura", equipment_code:row.equipment_code || "", equipment_description:row.equipment_description || "",
          component:row.component || "", system, date:row.work_date || "", due_date:"", hours_remaining:null,
          detail: detail || "Captura con condicion de revision.", action:"Generar OT correctiva o validar cierre en bitacora.",
        });
      });
      published.filter(row => ["OT", "Requisicion"].includes(String(row.source || ""))).forEach(row => addRow(row));
      workOrderRows().filter(row => !workOrderClosed(row)).forEach(row => {
        const priority = String(row.priority || "").toUpperCase();
        addRow({
          score: priority === "URGENTE" ? 96 : (priority === "ALTA" ? 82 : priority === "MEDIA" ? 58 : 38),
          source:"OT",
          equipment_code:row.equipment_code || "",
          equipment_description:row.equipment_description || "",
          component:row.origin || "OT",
          system:classifyStoppage(row.description, row.action, row.origin),
          date:row.date || "",
          due_date:"",
          hours_remaining:null,
          work_order_id:row.folio || row.id || "",
          responsible:row.responsible || row.mechanic || "",
          flow_status:String(row.status || "ABIERTA").replace("ABIERTA","Pendiente").replace("EN PROCESO","En proceso"),
          detail:row.description || "",
          action:row.action || "Dar seguimiento y cerrar OT.",
        });
      });
      (requisitions || []).filter(row => ["ABIERTA", "AUTORIZADA"].includes(String(row.status || "").toUpperCase())).forEach(row => {
        const items = Array.isArray(row.items) ? row.items.length : Number(row.items || 0);
        const urgent = String(row.priority || "").toUpperCase().includes("URG");
        addRow({
          score: urgent ? 52 : 42, source:"Requisicion", equipment_code:row.equipment || "", equipment_description:"",
          component:"Refacciones", system:"Refacciones", date:row.request_date || "", due_date:"", hours_remaining:null,
          detail:`${row.folio || ""} ${row.status || ""} con ${items || 0} partida(s).`,
          action:"Dar seguimiento a compra/surtido para liberar trabajos.",
        });
      });
      return rows.sort((a,b) => Number(b.score || 0) - Number(a.score || 0) || String(a.date || "9999-12-31").localeCompare(String(b.date || "9999-12-31")));
    }
    function renderBacklog(){
      const start = $("backlogStart").value || (portal.period || {}).start || toIsoDate(new Date());
      const end = $("backlogEnd").value || (portal.period || {}).end || start;
      const level = $("backlogLevel").value;
      const source = $("backlogSource").value;
      const status = $("backlogStatus").value;
      const search = normalizedText($("backlogSearch").value);
      let rows = calculateBacklogRows(start, end).filter(row => {
        const dateValue = row.date || row.due_date || start;
        const dateOk = !dateValue || inRange(dateValue, start, end) || ["VENCIDO", "URGENTE"].includes(String(row.status || row.detail || "").toUpperCase());
        const levelOk = !level || row.level === level;
        const sourceOk = !source || row.source === source;
        const flowOk = !status || (row.flow_status || "Pendiente") === status;
        const text = normalizedText([row.flow_status,row.level,row.source,row.equipment_code,row.equipment_description,row.component,row.system,row.detail,row.action,row.responsible].join(" "));
        return dateOk && levelOk && sourceOk && flowOk && (!search || text.includes(search));
      });
      const stats = {
        total: rows.length,
        high: rows.filter(r => r.level === "ALTA").length,
        medium: rows.filter(r => r.level === "MEDIA").length,
        low: rows.filter(r => r.level === "BAJA").length,
        pending: rows.filter(r => (r.flow_status || "Pendiente") === "Pendiente").length,
        process: rows.filter(r => (r.flow_status || "Pendiente") === "En proceso").length,
        attended: rows.filter(r => (r.flow_status || "Pendiente") === "Atendido").length,
        cancelled: rows.filter(r => (r.flow_status || "Pendiente") === "Cancelado").length,
      };
      $("backlogTitle").textContent = `Backlog priorizado | ${start} a ${end}`;
      $("backlogCount").textContent = `${rows.length} registro(s)`;
      $("backlogStats").innerHTML = [
        ["Total", stats.total], ["Alta", stats.high], ["Media", stats.medium], ["Baja", stats.low],
        ["Pend.", stats.pending], ["Proceso", stats.process], ["Atend.", stats.attended], ["Canc.", stats.cancelled],
      ].map(([k,v]) => `<div class="stat"><strong>${v}</strong>${esc(k)}</div>`).join("");
      const systems = {};
      rows.forEach(row => {
        const key = row.system || "Sin clasificar";
        systems[key] = systems[key] || {system:key, count:0, high:0, score:0};
        systems[key].count += 1; systems[key].score += Number(row.score || 0); if(row.level === "ALTA") systems[key].high += 1;
      });
      const systemRows = Object.values(systems).sort((a,b) => b.high - a.high || b.score - a.score || a.system.localeCompare(b.system));
      $("backlogTable").innerHTML = `<thead><tr><th>Estado</th><th>Nivel</th><th>Puntaje</th><th>Origen</th><th>Equipo</th><th>Componente</th><th>Sistema</th><th>Fecha</th><th>Vence</th><th>Hrs rest.</th><th>OT</th><th>Responsable</th><th>Cierre</th><th>Detalle</th><th>Accion</th></tr></thead><tbody>` +
        rows.map(row => `<tr><td><span class="pill ${flowClass(row.flow_status)}">${esc(row.flow_status || "Pendiente")}</span></td><td><span class="pill ${levelClass(row.level)}">${esc(row.level)}</span></td><td>${one(row.score)}</td><td>${esc(row.source)}</td><td>${esc(row.equipment_code)}</td><td>${esc(row.component)}</td><td>${esc(row.system)}</td><td>${esc(row.date || "")}</td><td>${esc(row.due_date || "")}</td><td>${row.hours_remaining == null ? "" : one(row.hours_remaining)}</td><td>${esc(row.work_order_id || "")}</td><td>${esc(row.responsible || "")}</td><td>${esc(row.closed_at || "")}</td><td>${esc(shortText(row.detail, 160))}</td><td>${esc(shortText(row.action, 140))}</td></tr>`).join("") +
        `</tbody>`;
      $("backlogSystemTable").innerHTML = `<thead><tr><th>Sistema</th><th>Total</th><th>Altas</th><th>Puntaje</th></tr></thead><tbody>` +
        systemRows.map(row => `<tr><td>${esc(row.system)}</td><td>${row.count}</td><td>${row.high}</td><td>${one(row.score)}</td></tr>`).join("") +
        `</tbody>`;
    }
    function filteredServiceHistory(){
      const selected = $("srvEquipment").value;
      const interval = $("srvInterval").value;
      const type = $("srvType").value;
      const start = $("srvStart").value;
      const end = $("srvEnd").value;
      const search = ($("srvSearch").value || "").toUpperCase();
      const rows = (Array.isArray(portal.service_history) ? portal.service_history : []).filter(row => {
        const service = [row.service_interval, row.service_name, row.stage].join(" ").toUpperCase();
        const serviceType = String(row.service_type || "Programado").toUpperCase();
        const eqOk = !selected || row.equipment_code === selected;
        const intervalOk = !interval || service === interval.toUpperCase() || service.includes(interval.toUpperCase());
        const typeOk = !type || serviceType === type.toUpperCase();
        const dateOk = !start || !end || inRange(row.completed_date, start, end);
        const text = [row.equipment_code,row.equipment_description,row.component,row.service_type,row.stage,row.service_name,row.service_interval,row.order_number,row.document_name,row.notes,row.status].join(" ").toUpperCase();
        return eqOk && intervalOk && typeOk && dateOk && (!search || text.includes(search));
      }).sort((a,b) => String(b.completed_date || "").localeCompare(String(a.completed_date || "")) || String(a.equipment_code || "").localeCompare(String(b.equipment_code || "")));
      return {start, end, rows};
    }
    function renderServiceHistory(){
      const result = filteredServiceHistory();
      $("srvTitle").textContent = `Servicios realizados | ${result.start || ""} a ${result.end || ""}`;
      $("srvCount").textContent = `${result.rows.length} servicio(s)`;
      const byInterval = result.rows.reduce((acc, row) => {
        const key = String(row.service_interval || row.service_name || "S/D").toUpperCase();
        acc[key] = (acc[key] || 0) + 1;
        return acc;
      }, {});
      const programmed = result.rows.filter(row => String(row.service_type || "Programado").toUpperCase() !== "NO PROGRAMADO").length;
      const unplanned = result.rows.filter(row => String(row.service_type || "").toUpperCase() === "NO PROGRAMADO").length;
      const late = result.rows.filter(row => String(row.status || "").toUpperCase() === "TARDIO").length;
      $("srvStats").innerHTML = [
        `<div class="stat"><strong>${result.rows.length}</strong>Realizados</div>`,
        `<div class="stat"><strong>${programmed}</strong>Programados</div>`,
        `<div class="stat"><strong>${unplanned}</strong>No programados</div>`,
        `<div class="stat"><strong>${byInterval["250H"] || 0}</strong>250H</div>`,
        `<div class="stat"><strong>${byInterval["500H"] || 0}</strong>500H</div>`,
        `<div class="stat"><strong>${byInterval["750H"] || 0}</strong>750H</div>`,
        `<div class="stat"><strong>${byInterval["1000H"] || 0}</strong>1000H</div>`,
        `<div class="stat"><strong>${late}</strong>Tardios</div>`,
      ].join("");
      const meter = (value) => {
        const number = Number(value || 0);
        return number > 0 ? one(number) : "";
      };
      $("srvTable").innerHTML = `<thead><tr><th>Fecha</th><th>Tipo</th><th>Etapa</th><th>Equipo</th><th>Descripcion</th><th>Componente</th><th>Servicio</th><th>Programado</th><th>Realizado</th><th>Fecha prog.</th><th>Estado</th><th>OT</th><th>Carta/gama</th><th>Filtros usados</th><th>Aceites</th><th>Detalle</th></tr></thead><tbody>` +
        result.rows.map(row => {
          const status = String(row.status || "");
          const cls = status === "A TIEMPO" ? "ok" : (status === "TARDIO" ? "bad" : "warn");
          const service = [row.service_name, row.service_interval].filter(Boolean).join(" / ");
          const documentText = row.document_name || (row.document_path ? "Registrada" : "");
          return `<tr><td>${esc(row.completed_date || "")}</td><td>${esc(row.service_type || "Programado")}</td><td>${esc(row.stage || "Cerrado")}</td><td>${esc(row.equipment_code || "")}</td><td>${esc(row.equipment_description || "")}</td><td>${esc(row.component || "")}</td><td>${esc(service)}</td><td>${esc(meter(row.scheduled_meter))}</td><td>${esc(meter(row.completed_meter))}</td><td>${esc(row.due_date || "")}</td><td><span class="pill ${cls}">${esc(status || "SIN FECHA")}</span></td><td>${esc(row.order_number || "")}</td><td>${esc(documentText)}</td><td>${esc(shortText(serviceFiltersText(row), 100))}</td><td>${esc(shortText(serviceOilsText(row), 100))}</td><td>${esc(shortText(row.notes || ""))}</td></tr>`;
        }).join("") +
        `</tbody>`;
    }
    function resetPreventiveExecutionForm(){
      currentPreventiveExecutionRecord = null;
      $("prevExecId").value = "";
      $("prevExecDate").value = toIsoDate(new Date());
      if(!$("prevExecEquipment").value && $("prevExecEquipment").options.length > 1) $("prevExecEquipment").selectedIndex = 1;
      $("prevExecSupervisor").value = "";
      $("prevExecMechanic").value = "";
      $("prevExecServiceType").value = "PM1";
      $("prevExecAttribute").value = "GENERAL";
      $("prevExecMeter").value = "0";
      $("prevExecState").value = "ABIERTO";
      $("prevExecParts").value = "";
      preventiveOilInputs.forEach(item => { $(item[0]).value = "0"; });
      preventiveChecklistInputs.forEach(item => { $(item[0]).checked = false; });
      updatePreventiveOilTotal();
      $("prevExecEvidence").value = "";
      $("prevExecNotes").value = "";
      $("prevExecStatus").textContent = "";
    }
    function updatePreventiveOilTotal(){
      const total = preventiveOilInputs.reduce((sum, item) => sum + captureNumber(item[0]), 0);
      $("prevExecOilTotal").value = total ? String(Math.round(total * 100) / 100) : "0";
    }
    function preventiveExecutionPayload(){
      const equipmentCode = $("prevExecEquipment").value || "";
      const equipment = portalEquipment().find(item => (item.code || item.equipment_code || "") === equipmentCode) || {};
      const payload = {
        id: $("prevExecId").value || "",
        service_date: $("prevExecDate").value || toIsoDate(new Date()),
        equipment_code: equipmentCode,
        equipment_description: equipment.description || equipment.family || "",
        supervisor: $("prevExecSupervisor").value.trim(),
        mechanic: $("prevExecMechanic").value.trim(),
        service_type: $("prevExecServiceType").value,
        attribute_type: $("prevExecAttribute").value,
        completed_meter: Number($("prevExecMeter").value || 0),
        status: $("prevExecState").value,
        parts_used: $("prevExecParts").value.trim(),
        lubricants_used: "",
        checklist: preventiveChecklistPayload(),
        evidence_note: $("prevExecEvidence").value.trim(),
        notes: $("prevExecNotes").value.trim(),
      };
      preventiveOilInputs.forEach(item => { payload[item[1]] = captureNumber(item[0]); });
      payload.oil_liters = preventiveOilTotal(payload);
      payload.lubricants_used = preventiveOilsText(payload);
      return payload;
    }
    function fillPreventiveExecutionForm(row){
      currentPreventiveExecutionRecord = row || null;
      $("prevExecId").value = row.id || "";
      $("prevExecDate").value = row.service_date || row.close_date || toIsoDate(new Date());
      $("prevExecEquipment").value = row.equipment_code || "";
      $("prevExecSupervisor").value = row.supervisor || "";
      $("prevExecMechanic").value = row.mechanic || "";
      $("prevExecServiceType").value = row.service_type || "PM1";
      $("prevExecAttribute").value = row.attribute_type || "GENERAL";
      $("prevExecMeter").value = Number(row.completed_meter || 0);
      $("prevExecState").value = row.status || "ABIERTO";
      $("prevExecParts").value = row.parts_used || "";
      preventiveOilInputs.forEach(item => { $(item[0]).value = formatCaptureNumber(row[item[1]]); });
      preventiveChecklistInputs.forEach(item => { $(item[0]).checked = Boolean((row.checklist || {})[item[1]]); });
      updatePreventiveOilTotal();
      $("prevExecEvidence").value = row.evidence_note || "";
      $("prevExecNotes").value = row.notes || "";
      $("prevExecStatus").textContent = `Editando ${row.id || ""}`;
      document.querySelector('[data-tab="ejecucionPreventivos"]')?.click();
    }
    function renderPreventiveExecution(){
      const rows = preventiveExecutionRows().slice().sort((a,b) => String(b.service_date || "").localeCompare(String(a.service_date || "")) || String(b.updated_at || "").localeCompare(String(a.updated_at || "")));
      const open = rows.filter(row => !isPreventiveClosed(row) && String(row.status || "").toUpperCase() !== "CANCELADO");
      const closed = rows.filter(isPreventiveClosed);
      $("prevExecOpenCount").textContent = `${open.length} abierto(s)`;
      $("prevExecClosedCount").textContent = `${closed.length} cerrado(s)`;
      const openBody = open.map(row => `<tr data-prev-exec-id="${esc(row.id)}"><td>${esc(row.service_date || "")}</td><td>${esc(row.equipment_code || "")}</td><td>${esc(row.service_type || "")}</td><td>${esc(row.attribute_type || "")}</td><td>${esc(row.supervisor || "")}</td><td>${esc(row.mechanic || "")}</td><td>${preventiveChecklistCount(row)}/6</td><td><span class="pill warn">${esc(row.status || "")}</span></td><td>${esc(shortText(row.notes || "", 90))}</td></tr>`).join("") || `<tr><td colspan="9">Sin servicios preventivos abiertos.</td></tr>`;
      $("prevExecOpenTable").innerHTML = `<thead><tr><th>Fecha</th><th>Equipo</th><th>PM</th><th>Atributo</th><th>Supervisor</th><th>Mecanico</th><th>Checklist</th><th>Estatus</th><th>Notas</th></tr></thead><tbody>${openBody}</tbody>`;
      const closedBody = closed.map(row => `<tr data-prev-exec-id="${esc(row.id)}"><td>${esc(row.close_date || row.service_date || "")}</td><td>${esc(row.equipment_code || "")}</td><td>${esc(row.service_type || "")}</td><td>${esc(row.attribute_type || "")}</td><td>${one(row.completed_meter || 0)}</td><td>${preventiveChecklistCount(row)}/6</td><td>${esc(shortText(row.parts_used || "", 110))}</td><td>${esc(shortText(preventiveOilsText(row) || row.lubricants_used || "", 110))}</td><td><span class="pill ok">${esc(row.status || "CERRADO")}</span></td></tr>`).join("") || `<tr><td colspan="9">Sin servicios cerrados desde esta pestaña.</td></tr>`;
      $("prevExecClosedTable").innerHTML = `<thead><tr><th>Fecha cierre</th><th>Equipo</th><th>PM</th><th>Atributo</th><th>Horometro</th><th>Checklist</th><th>Refacciones</th><th>Lubricantes</th><th>Estatus</th></tr></thead><tbody>${closedBody}</tbody>`;
      document.querySelectorAll("[data-prev-exec-id]").forEach(row => row.addEventListener("click", () => {
        const record = rows.find(item => String(item.id || "") === String(row.dataset.prevExecId || ""));
        if(record) fillPreventiveExecutionForm(record);
      }));
    }
    async function savePreventiveExecution(close=false){
      if(!hasApiKey()) return;
      const payload = preventiveExecutionPayload();
      if(close) payload.status = "CERRADO";
      if(!payload.equipment_code) return alert("Selecciona un equipo.");
      if(close && Object.values(payload.checklist || {}).filter(Boolean).length < preventiveChecklistInputs.length && !confirm("El checklist de cierre no esta completo. ¿Cerrar servicio de todos modos?")) return;
      if(!payload.supervisor && !payload.mechanic && !confirm("No capturaste supervisor ni mecanico. ¿Guardar asi?")) return;
      const response = await fetch("/api/preventive-execution/records", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      if(result.portal) portal = result.portal;
      renderPortalSelectors();
      renderPreventiveExecution();
      renderServiceHistory();
      renderPreventives();
      renderBacklog();
      renderExecutiveBoard();
      $("prevExecStatus").textContent = close ? "Servicio cerrado y enviado a Servicios realizados." : "Servicio guardado.";
      if(result.record) fillPreventiveExecutionForm(result.record);
    }
    async function deletePreventiveExecution(){
      if(!hasApiKey()) return;
      const id = $("prevExecId").value || (currentPreventiveExecutionRecord || {}).id || "";
      if(!id) return alert("Selecciona un servicio para eliminar.");
      if(!confirm("¿Eliminar este servicio preventivo? Si ya estaba cerrado tambien se retirara de Servicios realizados.")) return;
      const response = await fetch("/api/preventive-execution/records/delete", {method:"POST", headers:headers(true), body:JSON.stringify({id})});
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      if(result.portal) portal = result.portal;
      renderPortalSelectors();
      renderPreventiveExecution();
      renderServiceHistory();
      renderPreventives();
      renderBacklog();
      renderExecutiveBoard();
      resetPreventiveExecutionForm();
      $("prevExecStatus").textContent = "Servicio eliminado.";
    }
    function renderSpareParts(){
      const payload = portal.parts_manuals || {rows: [], summary: {}};
      const selected = $("spareEquipment").value || "";
      const status = $("spareStatus").value || "";
      const search = normalizedText($("spareSearch").value || "");
      const rows = (Array.isArray(payload.rows) ? payload.rows : []).filter(row => {
        const eqOk = !selected || row.equipment_code === selected;
        const statusOk = !status || row.inventory_status === status;
        const text = normalizedText([row.equipment_code,row.equipment_description,row.system,row.component,row.service_interval,row.part_number,row.equivalent_part,row.description,row.brand,row.criticality,row.manual_title,row.notes].join(" "));
        return eqOk && statusOk && (!search || text.includes(search));
      }).sort((a,b) => String(a.equipment_code || "").localeCompare(String(b.equipment_code || "")) || String(a.system || "").localeCompare(String(b.system || "")) || String(a.description || "").localeCompare(String(b.description || "")));
      const summary = payload.summary || {};
      const missing = rows.filter(row => row.inventory_status === "SIN INVENTARIO").length;
      const shortage = rows.filter(row => row.inventory_status === "FALTANTE").length;
      const ok = rows.filter(row => row.inventory_status === "DISPONIBLE").length;
      $("spareCount").textContent = `${rows.length} refaccion(es)`;
      $("spareStats").innerHTML = [
        `<div class="stat"><strong>${summary.manuals || 0}</strong>Manuales</div>`,
        `<div class="stat"><strong>${rows.length}</strong>Mostradas</div>`,
        `<div class="stat"><strong>${ok}</strong>Disponibles</div>`,
        `<div class="stat"><strong>${shortage}</strong>Faltantes</div>`,
        `<div class="stat"><strong>${missing}</strong>Sin inventario</div>`,
        `<div class="stat"><strong>${summary.unique_parts || 0}</strong>Partes unicas</div>`,
      ].join("");
      $("spareTable").innerHTML = `<thead><tr><th>Estado</th><th>Equipo</th><th>Descripcion equipo</th><th>Sistema</th><th>Componente</th><th>Servicio</th><th>No. parte</th><th>Equiv.</th><th>Descripcion</th><th>Cant.</th><th>Exist.</th><th>Falt.</th><th>Unidad</th><th>Crit.</th><th>Manual</th></tr></thead><tbody>` +
        rows.map(row => {
          const cls = row.inventory_status === "DISPONIBLE" ? "ok" : "bad";
          return `<tr><td><span class="pill ${cls}">${esc(row.inventory_status || "")}</span></td><td>${esc(row.equipment_code || "")}</td><td>${esc(row.equipment_description || "")}</td><td>${esc(row.system || "")}</td><td>${esc(row.component || "")}</td><td>${esc(row.service_interval || "")}</td><td>${esc(row.part_number || "")}</td><td>${esc(row.equivalent_part || "")}</td><td>${esc(row.description || "")}</td><td>${one(row.quantity)}</td><td>${row.available == null ? "" : one(row.available)}</td><td>${one(row.shortage)}</td><td>${esc(row.unit || "")}</td><td>${esc(row.criticality || "")}</td><td>${esc(row.manual_title || "")}</td></tr>`;
        }).join("") +
        `</tbody>`;
    }
    function renderAudit(){
      const start = $("auditStart").value || "";
      const end = $("auditEnd").value || "";
      const module = $("auditModule").value || "";
      const search = normalizedText($("auditSearch").value || "");
      const rows = (Array.isArray(portal.audit_log) ? portal.audit_log : []).filter(row => {
        const created = String(row.created_at || "").slice(0, 10);
        const dateOk = (!start || created >= start) && (!end || created <= end);
        const moduleOk = !module || row.module === module;
        const text = normalizedText([row.user_name,row.role,row.module,row.action,row.entity_type,row.entity_id,row.summary,row.payload_json].join(" "));
        return dateOk && moduleOk && (!search || text.includes(search));
      }).sort((a,b) => String(b.created_at || "").localeCompare(String(a.created_at || "")));
      $("auditCount").textContent = `${rows.length} movimiento(s)`;
      $("auditTable").innerHTML = `<thead><tr><th>Fecha</th><th>Usuario</th><th>Rol</th><th>Modulo</th><th>Accion</th><th>Entidad</th><th>ID</th><th>Resumen</th></tr></thead><tbody>` +
        rows.map(row => `<tr><td>${esc(row.created_at || "")}</td><td>${esc(row.user_name || "")}</td><td>${esc(row.role || "")}</td><td>${esc(row.module || "")}</td><td>${esc(row.action || "")}</td><td>${esc(row.entity_type || "")}</td><td>${esc(row.entity_id || "")}</td><td>${esc(shortText(row.summary || "", 180))}</td></tr>`).join("") +
        `</tbody>`;
    }
    function renderBitacora(){
      const selected = $("bitEquipment").value;
      const start = $("bitStart").value;
      const end = $("bitEnd").value;
      const search = ($("bitSearch").value || "").toUpperCase();
      const rows = (portal.captures || []).filter(row => {
        const eqOk = !selected || row.equipment_code === selected;
        const text = [row.component,row.fault,row.wear,row.status,row.observations].join(" ").toUpperCase();
        return eqOk && inRange(row.work_date, start, end) && (!search || text.includes(search));
      });
      $("bitTable").innerHTML = `<thead><tr><th>Fecha</th><th>Turno</th><th>Equipo</th><th>Componente</th><th>HI</th><th>HF</th><th>Hrs Trab</th><th>MP</th><th>MC</th><th>Stand By</th><th>Paradas</th><th>Aceite L</th><th>Estatus</th><th>Falla / observaciones</th><th>Fotos</th></tr></thead><tbody>` +
        rows.map(row => `<tr><td>${esc(row.work_date)}</td><td>${esc(row.shift)}</td><td>${esc(row.equipment_code)}</td><td>${esc(row.component)}</td><td>${one(row.hi)}</td><td>${one(row.hf)}</td><td>${one(row.worked_hours)}</td><td>${one(row.mp_hours)}</td><td>${one(row.mc_hours)}</td><td>${one(row.standby_hours)}</td><td>${num(row.stops)}</td><td>${one(row.oil_liters)}</td><td>${esc(row.status)}</td><td>${esc([row.fault,row.observations].filter(Boolean).join(" | "))}</td><td>${num(row.evidence_count)}</td></tr>`).join("") +
        `</tbody>`;
    }
    function sameCaptureEquipment(rowCode, selected){
      if(!selected) return true;
      const selectedKeys = equipmentKeys(selected);
      return equipmentKeys(rowCode).some(key => selectedKeys.includes(key));
    }
    function defaultCaptureComponents(code){
      const normalized = String(code || "").toUpperCase();
      if(normalized.startsWith("JL") || normalized.startsWith("JA")) return ["ELECT", "DIESEL", "COMPRESOR", "PER"];
      if(normalized.startsWith("ST") || normalized.startsWith("RET")) return ["DIESEL"];
      return ["MOTOR", "DIESEL", "ELECT"];
    }
    function componentsForCaptureEquipment(code){
      const values = [];
      const add = value => {
        const clean = String(value || "").trim().toUpperCase();
        if(clean && !values.includes(clean)) values.push(clean);
      };
      const eq = portalEquipment().find(item => sameCaptureEquipment(item.code || item.equipment_code, code));
      if(eq && Array.isArray(eq.components)) eq.components.forEach(component => add(component.name || component.component || component));
      (portal.captures || []).forEach(row => {
        if(sameCaptureEquipment(row.equipment_code || row.equipment, code)) add(row.component || row.component_name);
      });
      defaultCaptureComponents(code).forEach(add);
      return values;
    }
    function renderCaptureComponents(){
      const select = $("capComponent");
      const previous = select.value;
      const options = componentsForCaptureEquipment($("capEquipment").value);
      select.innerHTML = options.map(value => `<option value="${esc(value)}">${esc(value)}</option>`).join("");
      if(options.includes(previous)) select.value = previous;
      else if(options.length) select.value = options[0];
      applyPreviousHi(false);
    }
    function captureShiftOrder(value){
      const text = normalizedText(value).replace(/\s/g, "");
      if(["2","T2","TURNO2","SEGUNDO","NOCHE"].includes(text)) return 2;
      if(["1","T1","TURNO1","PRIMERO","DIA"].includes(text)) return 1;
      return 0;
    }
    function sortedCaptureRows(rows){
      return [...rows].sort((a,b) => {
        const aKey = `${a.work_date || ""}-${String(captureShiftOrder(a.shift)).padStart(2,"0")}-${String(a.id || 0).padStart(10,"0")}`;
        const bKey = `${b.work_date || ""}-${String(captureShiftOrder(b.shift)).padStart(2,"0")}-${String(b.id || 0).padStart(10,"0")}`;
        return bKey.localeCompare(aKey);
      });
    }
    function sameCaptureComponent(rowComponent, selected){
      if(!selected) return true;
      return normalizedText(rowComponent) === normalizedText(selected);
    }
    function previousCaptureForHi(){
      const equipment = $("capEquipment").value;
      const component = $("capComponent").value;
      const workDate = $("capDate").value;
      if(!equipment || !component || !workDate) return null;
      const currentShiftOrder = captureShiftOrder($("capShift").value);
      const previousRows = (portal.captures || []).filter(row =>
        sameCaptureEquipment(row.equipment_code || row.equipment, equipment) &&
        (
          String(row.work_date || "") < workDate ||
          (String(row.work_date || "") === workDate && captureShiftOrder(row.shift) < currentShiftOrder)
        ) &&
        Number(row.hf || 0) > 0
      );
      const exactComponent = previousRows.filter(row => sameCaptureComponent(row.component || row.component_name, component));
      return sortedCaptureRows(exactComponent)[0] || sortedCaptureRows(previousRows)[0] || null;
    }
    function formatCaptureNumber(value){
      const number = Number(value || 0);
      return Number.isInteger(number) ? String(number) : number.toFixed(1).replace(/\.0$/, "");
    }
    function applyPreviousHi(force=false){
      if(currentCaptureRecord && !force) return;
      const currentHi = String($("capHi").value || "").trim();
      if(!force && currentHi && Number(currentHi) > 0) return;
      const previous = previousCaptureForHi();
      if(previous){
        $("capHi").value = formatCaptureNumber(previous.hf);
        $("capHiHint").textContent = `HI automatico: HF ${formatCaptureNumber(previous.hf)} del ${previous.work_date || ""}`;
      } else {
        $("capHiHint").textContent = "Sin HF anterior para este equipo/componente.";
        if(force && (!currentHi || Number(currentHi) <= 0)) $("capHi").value = "0";
      }
      updateCaptureWorkedHours();
    }
    function normalizeCaptureShiftValue(value){
      const text = normalizedText(value);
      if(text.includes("2")) return "Turno 2";
      if(text.includes("GENERAL")) return "General";
      return "Turno 1";
    }
    function normalizeCaptureStatusValue(value){
      const text = normalizedText(value);
      if(text.includes("STAND")) return "Stand By";
      if(text.includes("NO DISP") || text.includes("FUERA")) return "No Disponible";
      if(text.includes("OPERATIVA")) return "Operativa";
      return "Disponible";
    }
    function setCaptureComponent(value){
      const clean = String(value || "").trim().toUpperCase();
      renderCaptureComponents();
      if(clean && ![...$("capComponent").options].some(option => option.value === clean)){
        $("capComponent").insertAdjacentHTML("beforeend", `<option value="${esc(clean)}">${esc(clean)}</option>`);
      }
      if(clean) $("capComponent").value = clean;
    }
    function resetCaptureForm(){
      currentCaptureRecord = null;
      if(!$("capDate").value) $("capDate").value = toIsoDate(new Date());
      $("capShift").value = "Turno 1";
      ["capHi","capHf","capWorked","capMp","capMc","capStandby","capStops","capOil","capOil15w40","capOilHco68","capOilSae30","capOil85w140","capAlmo","capCoolant","capOilVg100","capAtf"].forEach(id => { $(id).value = "0"; });
      ["capFault","capWear","capObservations"].forEach(id => { $(id).value = ""; });
      $("capCaptureStatus").value = "Disponible";
      $("capStatus").textContent = "";
      $("capSaveBtn").textContent = "Guardar captura";
      renderCaptureComponents();
      applyPreviousHi(true);
      renderCaptureRecent();
    }
    function captureNumber(id){
      const value = Number($(id).value || 0);
      return Number.isFinite(value) && value > 0 ? value : 0;
    }
    function updateCaptureWorkedHours(){
      const hi = Number($("capHi").value || 0);
      const hf = Number($("capHf").value || 0);
      if(Number.isFinite(hi) && Number.isFinite(hf) && hf >= hi){
        $("capWorked").value = one(hf - hi).replace(/\.0$/, "");
      }
    }
    function updateCaptureOilTotal(){
      const detailIds = ["capOil15w40","capOilHco68","capOilSae30","capOil85w140","capAlmo","capCoolant","capOilVg100","capAtf"];
      const total = detailIds.reduce((sum, id) => sum + captureNumber(id), 0);
      $("capOil").value = total ? String(Math.round(total * 100) / 100) : "0";
    }
    function captureMobileId(record){
      return ["web", record.work_date, record.shift, record.equipment_code, record.component_name]
        .map(value => normalizedText(value).replace(/[^A-Z0-9]/g, ""))
        .join(":");
    }
    function capturePayload(){
      const workDate = $("capDate").value;
      const equipment = $("capEquipment").value;
      const component = $("capComponent").value;
      if(!workDate) throw new Error("Selecciona la fecha de captura.");
      if(!equipment) throw new Error("Selecciona el equipo.");
      if(!component) throw new Error("Selecciona el componente.");
      const oilMotor = captureNumber("capOil15w40");
      const oilHco = captureNumber("capOilHco68");
      const oilSae = captureNumber("capOilSae30");
      const oil85w140 = captureNumber("capOil85w140");
      const almo = captureNumber("capAlmo");
      const coolant = captureNumber("capCoolant");
      const oilVg100 = captureNumber("capOilVg100");
      const atf = captureNumber("capAtf");
      const record = {
        work_date: workDate,
        shift: normalizeCaptureShiftValue($("capShift").value),
        equipment_code: equipment,
        equipment: equipment,
        component_name: component,
        component: component,
        hi: captureNumber("capHi"),
        hf: captureNumber("capHf"),
        worked_hours: captureNumber("capWorked"),
        mp_hours: captureNumber("capMp"),
        mc_hours: captureNumber("capMc"),
        standby_hours: captureNumber("capStandby"),
        stops: Math.round(captureNumber("capStops")),
        oil_liters: captureNumber("capOil") || oilMotor + oilHco + oilSae + oil85w140 + almo + coolant + oilVg100 + atf,
        oil_motor_15w40: oilMotor,
        oil_hco_iso68: oilHco,
        oil_trans_sae30: oilSae,
        oil_85w140: oil85w140,
        almo_liters: almo,
        coolant_liters: coolant,
        oil_hyd_vg100: oilVg100,
        atf_liters: atf,
        fault: $("capFault").value.trim(),
        wear: $("capWear").value.trim(),
        status: $("capCaptureStatus").value || "Disponible",
        observations: $("capObservations").value.trim(),
        captured_at: new Date().toISOString(),
        user_name: "Portal web",
        source: "web",
        photos: [],
      };
      record.mobile_id = currentCaptureRecord?.mobile_id || captureMobileId(record);
      if(currentCaptureRecord){
        record.original_work_date = currentCaptureRecord.work_date || "";
        record.original_shift = currentCaptureRecord.shift || "";
        record.original_equipment_code = currentCaptureRecord.equipment_code || "";
        record.original_component_name = currentCaptureRecord.component || currentCaptureRecord.component_name || "";
        record.original_capture_id = currentCaptureRecord.id || "";
      }
      return record;
    }
    function captureValidationWarnings(){
      const warnings = [];
      ["capDate","capEquipment","capComponent","capHi","capHf","capWorked"].forEach(id => $(id).classList.remove("input-warning"));
      if(!$("capDate").value){ warnings.push("Falta fecha."); $("capDate").classList.add("input-warning"); }
      if(!$("capEquipment").value){ warnings.push("Falta equipo."); $("capEquipment").classList.add("input-warning"); }
      if(!$("capComponent").value){ warnings.push("Falta componente."); $("capComponent").classList.add("input-warning"); }
      const hi = Number($("capHi").value || 0);
      const hf = Number($("capHf").value || 0);
      const worked = Number($("capWorked").value || 0);
      if(hf && hi && hf < hi){ warnings.push("Horometro final menor al inicial."); $("capHf").classList.add("input-warning"); }
      if(!worked && hf > hi){ warnings.push("Horas trabajadas en cero; revisa antes de guardar."); $("capWorked").classList.add("input-warning"); }
      if(Number($("capMp").value || 0) + Number($("capMc").value || 0) > worked && worked > 0) warnings.push("Hrs MP + MC superan las horas trabajadas.");
      $("capQuickAlerts").innerHTML = warnings.map(item => `<div class="quick-alert">${esc(item)}</div>`).join("");
      return warnings;
    }
    async function saveDailyCapture(resetAfter=false){
      if(!hasApiKey(true)) return;
      const warnings = captureValidationWarnings();
      if(warnings.some(text => text.includes("Falta") || text.includes("menor")) && !confirm("Hay alertas de captura. ¿Guardar de todos modos?")) return;
      const record = capturePayload();
      $("capStatus").textContent = "Guardando...";
      const response = await fetch("/api/sync", {
        method: "POST",
        headers: headers(true),
        body: JSON.stringify({device: "portal-web", user: "Portal web", records: [record]}),
      });
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      if(!result.ok) throw new Error(`No se pudo guardar la captura. Errores: ${result.errors || 0}`);
      const wasEditing = !!currentCaptureRecord;
      upsertLocalCapture(record, result);
      currentCaptureRecord = null;
      $("capSaveBtn").textContent = "Guardar captura";
      $("capStatus").textContent = wasEditing || result.updated ? "Captura actualizada." : "Captura guardada.";
      renderDashboard();
      renderBitacora();
      renderCaptureRecent();
      renderExecutiveBoard();
      if(resetAfter) resetCaptureForm();
    }
    function editDailyCapture(index){
      const row = currentCaptureRows[index];
      if(!row) return;
      currentCaptureRecord = row;
      $("capDate").value = row.work_date || toIsoDate(new Date());
      $("capShift").value = normalizeCaptureShiftValue(row.shift);
      $("capEquipment").value = row.equipment_code || row.equipment || "";
      setCaptureComponent(row.component || row.component_name || "");
      $("capHi").value = formatCaptureNumber(row.hi);
      $("capHf").value = formatCaptureNumber(row.hf);
      $("capWorked").value = formatCaptureNumber(row.worked_hours);
      $("capMp").value = formatCaptureNumber(row.mp_hours);
      $("capMc").value = formatCaptureNumber(row.mc_hours);
      $("capStandby").value = formatCaptureNumber(row.standby_hours);
      $("capStops").value = formatCaptureNumber(row.stops);
      $("capOil").value = formatCaptureNumber(row.oil_liters);
      $("capOil15w40").value = formatCaptureNumber(row.oil_motor_15w40);
      $("capOilHco68").value = formatCaptureNumber(row.oil_hco_iso68);
      $("capOilSae30").value = formatCaptureNumber(row.oil_trans_sae30);
      $("capOil85w140").value = formatCaptureNumber(row.oil_85w140);
      $("capAlmo").value = formatCaptureNumber(row.almo_liters);
      $("capCoolant").value = formatCaptureNumber(row.coolant_liters);
      $("capOilVg100").value = formatCaptureNumber(row.oil_hyd_vg100);
      $("capAtf").value = formatCaptureNumber(row.atf_liters);
      $("capFault").value = row.fault || "";
      $("capWear").value = row.wear || "";
      $("capCaptureStatus").value = normalizeCaptureStatusValue(row.status);
      $("capObservations").value = row.observations || "";
      $("capHiHint").textContent = "Editando captura existente.";
      $("capSaveBtn").textContent = "Guardar cambios";
      $("capStatus").textContent = `Editando ${row.work_date || ""} ${row.equipment_code || ""} ${row.component || ""}`;
      renderCaptureRecent();
    }
    function captureRowKey(row){
      return [
        row?.work_date || "",
        normalizedText(row?.shift || ""),
        normalizedText(row?.equipment_code || row?.equipment || ""),
        normalizedText(row?.component || row?.component_name || ""),
      ].join("|");
    }
    function captureDeletePayload(row){
      return {
        id: row?.id || "",
        mobile_id: row?.mobile_id || "",
        work_date: row?.work_date || "",
        shift: row?.shift || "",
        equipment_code: row?.equipment_code || row?.equipment || "",
        component: row?.component || row?.component_name || "",
      };
    }
    function captureRowsSame(a, b){
      if(!a || !b) return false;
      const aMobile = String(a.mobile_id || "").trim();
      const bMobile = String(b.mobile_id || "").trim();
      if(aMobile && bMobile && aMobile === bMobile) return true;
      const aId = Number(a.id || 0);
      const bId = Number(b.id || 0);
      if(aId && bId && aId === bId && captureRowKey(a) === captureRowKey(b)) return true;
      return captureRowKey(a) === captureRowKey(b);
    }
    function removeLocalCapture(row){
      portal.captures = (portal.captures || []).filter(item => !captureRowsSame(item, row));
    }
    function upsertLocalCapture(record, result){
      const first = Array.isArray(result?.results) ? result.results[0] || {} : {};
      const saved = {
        id: first.capture_id || currentCaptureRecord?.id || Date.now(),
        mobile_id: record.mobile_id,
        source: "web",
        received_at: new Date().toISOString(),
        work_date: record.work_date,
        shift: record.shift,
        equipment_code: record.equipment_code,
        equipment_description: "",
        component: record.component_name || record.component || "",
        hi: record.hi,
        hf: record.hf,
        worked_hours: record.worked_hours,
        mp_hours: record.mp_hours,
        mc_hours: record.mc_hours,
        standby_hours: record.standby_hours,
        stops: record.stops,
        oil_liters: record.oil_liters,
        oil_motor_15w40: record.oil_motor_15w40,
        oil_hco_iso68: record.oil_hco_iso68,
        oil_trans_sae30: record.oil_trans_sae30,
        oil_85w140: record.oil_85w140,
        almo_liters: record.almo_liters,
        coolant_liters: record.coolant_liters,
        oil_hyd_vg100: record.oil_hyd_vg100,
        atf_liters: record.atf_liters,
        fault: record.fault,
        wear: record.wear,
        status: record.status,
        observations: record.observations,
        evidence_count: 0,
      };
      if(currentCaptureRecord) removeLocalCapture(currentCaptureRecord);
      removeLocalCapture(saved);
      portal.captures = [saved, ...(portal.captures || [])];
    }
    async function deleteDailyCapture(index){
      const row = currentCaptureRows[index];
      if(!row) return;
      if(!hasApiKey(true)) return;
      const label = `${row.work_date || ""} ${row.shift || ""} ${row.equipment_code || ""} ${row.component || ""}`.trim();
      if(!confirm(`Se eliminara la captura ${label}.`)) return;
      $("capStatus").textContent = "Eliminando...";
      const response = await fetch("/api/portal/captures/delete", {
        method: "POST",
        headers: headers(true),
        body: JSON.stringify(captureDeletePayload(row)),
      });
      if(!response.ok) throw new Error(await apiError(response));
      const result = await response.json();
      if(currentCaptureRecord && captureRowKey(currentCaptureRecord) === captureRowKey(row)){
        currentCaptureRecord = null;
        $("capSaveBtn").textContent = "Guardar captura";
      }
      removeLocalCapture(row);
      $("capStatus").textContent = result.deleted_mobile || result.deleted_portal ? "Captura eliminada." : "No se encontro la captura para eliminar.";
      renderDashboard();
      renderBitacora();
      renderCaptureRecent();
    }
    function renderCaptureRecent(){
      const selected = $("capEquipment").value;
      const rows = sortedCaptureRows((portal.captures || [])
        .filter(row => sameCaptureEquipment(row.equipment_code || row.equipment, selected))
      ).slice(0, 25);
      currentCaptureRows = rows;
      $("capRecentCount").textContent = `${rows.length} registros`;
      $("capRecentTable").innerHTML = `<thead><tr><th>Fecha</th><th>Turno</th><th>Equipo</th><th>Comp.</th><th>HI</th><th>HF</th><th>Hrs</th><th>MP</th><th>MC</th><th>Paradas</th><th>Estatus</th><th>Accion</th></tr></thead><tbody>` +
        rows.map((row, idx) => `<tr class="${currentCaptureRecord === row ? "warn" : ""}"><td>${esc(row.work_date)}</td><td>${esc(row.shift)}</td><td>${esc(row.equipment_code)}</td><td>${esc(row.component)}</td><td>${one(row.hi)}</td><td>${one(row.hf)}</td><td>${one(row.worked_hours)}</td><td>${one(row.mp_hours)}</td><td>${one(row.mc_hours)}</td><td>${num(row.stops)}</td><td>${esc(row.status)}</td><td><button type="button" class="btn secondary small" data-cap-edit="${idx}">Editar</button> <button type="button" class="btn danger small" data-cap-delete="${idx}">Eliminar</button></td></tr>`).join("") +
        `</tbody>`;
      document.querySelectorAll("[data-cap-edit]").forEach(button => button.addEventListener("click", () => editDailyCapture(Number(button.dataset.capEdit))));
      document.querySelectorAll("[data-cap-delete]").forEach(button => button.addEventListener("click", () => deleteDailyCapture(Number(button.dataset.capDelete)).catch(showError)));
    }
    function conditionClass(condition){
      const text = String(condition || "").toUpperCase();
      if(text.includes("FUERA") || text.includes("NO DISP")) return "cond-out";
      if(text.includes("OPERATIVA") || text.includes("REPARACION") || text.includes("STAND")) return "cond-warn";
      if(text.includes("DISPONIBLE")) return "cond-ok";
      return "";
    }
    function meterText(value){
      const number = Number(value || 0);
      if(!number || number <= 0) return "";
      return Number.isInteger(number) ? String(number) : `${number}`;
    }
    function availabilityShiftOrder(value){
      const text = normalizedText(value).replace(/\s/g, "");
      if(["2","T2","TURNO2","SEGUNDO","NOCHE"].includes(text)) return 2;
      if(["1","T1","TURNO1","PRIMERO","DIA"].includes(text)) return 1;
      return 0;
    }
    function putAvailabilityLookup(map, keys, value, order){
      const hf = Number(value || 0);
      if(!hf || hf <= 0) return;
      keys.forEach(key => {
        if(!key) return;
        const current = map.get(key);
        if(!current || order >= current.order) map.set(key, {order, hf});
      });
    }
    function availabilityLastHfLookup(){
      const map = new Map();
      (portal.captures || []).forEach(row => {
        const keys = [...equipmentKeys(row.equipment_code || row.code || row.equipment), ...equipmentKeys(row.equipment_description || "")];
        const order = `${row.work_date || ""}-${String(availabilityShiftOrder(row.shift)).padStart(2, "0")}-${String(row.id || 0).padStart(10, "0")}`;
        putAvailabilityLookup(map, keys, row.hf || row.horometer_final, order);
      });
      (diesel.records || []).forEach(row => {
        const order = `${row.work_date || ""}-${String(availabilityShiftOrder(row.shift)).padStart(2, "0")}-${String(row.id || 0).padStart(10, "0")}`;
        putAvailabilityLookup(map, equipmentKeys(row.equipment), row.horometer_final, order);
      });
      return map;
    }
    function availabilityKpiLookup(){
      const start = $("kpiStart").value || (portal.period || {}).start || toIsoDate(new Date());
      const end = $("kpiEnd").value || (portal.period || {}).end || start;
      const report = calculateKpiRows("Todos los equipos", start, end);
      const map = new Map();
      report.rows.forEach(row => {
        const value = {availability: row.availabilityText || pct(row.availability), utilization: row.utilizationText || pct(row.utilization)};
        [...equipmentKeys(row.code), ...equipmentKeys(row.description)].forEach(key => {
          if(key && !map.has(key)) map.set(key, value);
        });
      });
      return map;
    }
    function availabilityValueForRow(row, map){
      for(const key of [...equipmentKeys(row.eco), ...equipmentKeys(row.equipment)]){
        if(map.has(key)) return map.get(key);
      }
      return null;
    }
    function renderDisponibilidad(){
      const search = ($("dispSearch").value || "").toUpperCase();
      const status = $("dispStatus").value;
      const hfLookup = availabilityLastHfLookup();
      const kpiLookup = availabilityKpiLookup();
      const rows = (portal.availability || []).filter(row => {
        const text = [row.category,row.equipment,row.eco,row.condition,row.observations].join(" ").toUpperCase();
        return (!status || String(row.condition || "").toUpperCase().includes(status)) && (!search || text.includes(search));
      });
      $("dispTable").innerHTML = `<thead><tr><th>Categoria</th><th>Equipo</th><th>No ECO</th><th>Ultimo HF</th><th>% Disp</th><th>% Util</th><th>Condicion</th><th>Observaciones</th></tr></thead><tbody>` +
        rows.map(row => {
          const hf = availabilityValueForRow(row, hfLookup);
          const kpi = availabilityValueForRow(row, kpiLookup) || {};
          return `<tr><td>${esc(row.category)}</td><td>${esc(row.equipment)}</td><td>${esc(row.eco)}</td><td>${esc(meterText(hf?.hf))}</td><td>${esc(kpi.availability || "")}</td><td>${esc(kpi.utilization || "")}</td><td class="condition-cell ${conditionClass(row.condition)}">${esc(row.condition)}</td><td class="${Number(row.highlight_observation || 0) ? "highlight" : ""}">${esc(row.observations)}</td></tr>`;
        }).join("") +
        `</tbody>`;
    }
    function reqFormData(){
      return {
        id: currentReqId,
        folio: $("reqFolio").value,
        request_date: $("reqDate").value,
        authorization_date: $("reqAuthDate").value,
        equipment: $("reqEquipment").value,
        cost_center: $("reqCostCenter").value,
        request_area: $("reqArea").value,
        location: $("reqLocation").value,
        requesting_unit: $("reqRequestingUnit").value,
        operating_unit: $("reqOperatingUnit").value,
        priority: $("reqPriority").value,
        recommendation: $("reqRecommendation").value,
        status: $("reqReqStatus").value,
        notes: $("reqNotes").value,
        items: currentReqItems
      };
    }
    function setReqForm(row){
      currentReqId = row?.id || null;
      currentReqItems = (row?.items || []).map(item => ({...item}));
      currentReqItemIndex = null;
      $("reqFolio").value = row?.folio || "";
      $("reqDate").value = row?.request_date || toIsoDate(new Date());
      $("reqAuthDate").value = row?.authorization_date || $("reqDate").value;
      $("reqEquipment").value = row?.equipment || "PARA STOCK";
      $("reqCostCenter").value = row?.cost_center || "";
      $("reqArea").value = row?.request_area || "MTTO";
      $("reqLocation").value = row?.location || "PROVIDENCIA";
      $("reqRequestingUnit").value = row?.requesting_unit || "TALLER CENTRAL";
      $("reqOperatingUnit").value = row?.operating_unit || "PROVIDENCIA";
      $("reqPriority").value = row?.priority || "URGENTE";
      $("reqRecommendation").value = row?.recommendation || "ORIGINAL";
      $("reqReqStatus").value = row?.status || "Abierta";
      $("reqNotes").value = row?.notes || "";
      clearReqItem();
      renderReqItems();
      $("reqStatus").textContent = currentReqId ? `Editando ${$("reqFolio").value}` : "Nueva requisicion";
    }
    function newRequisition(nextFolio=""){
      setReqForm({ folio: nextFolio || "", request_date: toIsoDate(new Date()), authorization_date: toIsoDate(new Date()), items: [] });
    }
    function clearReqItem(){
      currentReqItemIndex = null;
      $("reqItemQty").value = "1";
      $("reqItemUnit").value = "PZA";
      $("reqItemPart").value = "";
      $("reqItemDesc").value = "";
    }
    function addReqItem(){
      const item = { quantity:Number($("reqItemQty").value || 1), unit:$("reqItemUnit").value || "PZA", part_number:$("reqItemPart").value.trim(), description:$("reqItemDesc").value.trim() };
      if(!item.part_number && !item.description) return alert("Captura No. parte o descripcion.");
      if(currentReqItemIndex == null) currentReqItems.push(item); else currentReqItems[currentReqItemIndex] = item;
      clearReqItem();
      renderReqItems();
    }
    function editReqItem(index){
      const item = currentReqItems[index];
      if(!item) return;
      currentReqItemIndex = index;
      $("reqItemQty").value = item.quantity || 1;
      $("reqItemUnit").value = item.unit || "PZA";
      $("reqItemPart").value = item.part_number || "";
      $("reqItemDesc").value = item.description || "";
    }
    function deleteReqItem(){
      if(currentReqItemIndex == null) return alert("Selecciona una partida.");
      currentReqItems.splice(currentReqItemIndex, 1);
      clearReqItem();
      renderReqItems();
    }
    function renderReqItems(){
      $("reqItemsTable").innerHTML = `<thead><tr><th>Cant.</th><th>Unidad</th><th>No. parte</th><th>Descripcion</th></tr></thead><tbody>` +
        currentReqItems.map((item, idx) => `<tr data-req-item="${idx}" style="cursor:pointer"><td>${num(item.quantity)}</td><td>${esc(item.unit)}</td><td>${esc(item.part_number)}</td><td>${esc(item.description)}</td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-req-item]").forEach(row => row.addEventListener("click", () => editReqItem(Number(row.dataset.reqItem))));
    }
    async function loadReq(rowId){
      const r = await fetch(`/api/requisitions/${rowId}`, {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      const payload = await r.json();
      setReqForm(payload.requisition);
    }
    async function saveReq(){
      if(!hasApiKey(true)) return;
      const r = await fetch("/api/requisitions", {method:"POST", headers:headers(true), body:JSON.stringify(reqFormData())});
      if(!r.ok) return alert(await apiError(r));
      const payload = await r.json();
      setReqForm(payload.requisition);
      await load();
    }
    async function deleteReq(){
      if(!currentReqId) return alert("Selecciona una requisicion.");
      await deleteReqById(currentReqId, $("reqFolio").value);
    }
    async function deleteReqById(rowId, folio=""){
      if(!rowId) return alert("Selecciona una requisicion.");
      if(!hasApiKey(true)) return false;
      const label = folio || `ID ${rowId}`;
      if(!confirm(`Eliminar requisicion ${label}? Esta accion no se puede deshacer.`)) return false;
      const r = await fetch(`/api/requisitions/${rowId}`, {method:"DELETE", headers:headers()});
      if(!r.ok) return alert(await apiError(r));
      if(String(currentReqId || "") === String(rowId)){
        currentReqId = null;
        currentReqItems = [];
        currentReqItemIndex = null;
        $("reqFolio").value = "";
      }
      if(String(currentTrackId || "") === String(rowId)){
        clearTrackForm();
      }
      await load();
      return true;
    }
    async function ensureReqForOutput(){
      if(currentReqId) return currentReqId;
      await saveReq();
      return currentReqId;
    }
    async function reqPdf(print=false){
      const id = await ensureReqForOutput();
      if(!id) return;
      if(!currentReqItems.length) return alert("Agrega al menos una partida antes de generar PDF.");
      const url = `/api/requisitions/${id}/pdf`;
      if(print){
        const win = window.open(url, "_blank");
        if(!win) alert("No se pudo abrir la ventana de impresion.");
        return;
      }
      const a = document.createElement("a");
      a.href = url;
      a.download = `Requisicion_${$("reqFolio").value || id}.pdf`;
      a.click();
    }
    function renderReqProducts(){
      const q = ($("reqProductSearch").value || "").toUpperCase();
      const rows = (products || []).filter(row => !q || [row.code,row.product].join(" ").toUpperCase().includes(q)).slice(0,120);
      $("reqProductsTable").innerHTML = `<thead><tr><th>Clave</th><th>Producto</th></tr></thead><tbody>` +
        rows.map(row => `<tr data-product-code="${esc(row.code)}" data-product-name="${esc(row.product)}" style="cursor:pointer"><td>${esc(row.code)}</td><td>${esc(row.product)}</td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-product-code]").forEach(row => row.addEventListener("click", () => {
        $("reqItemPart").value = row.dataset.productCode || "";
        $("reqItemDesc").value = row.dataset.productName || "";
      }));
    }
    function renderReqList(){
      $("reqItemUnit").innerHTML = ["PZA","JGO","KIT","SERV","LT","L","GAL","ML","TAMBO","TAMBOR","CUBETA","BOTE","LATA","CAJA","PAQUETE","BOLSA","MTS","M2","M3","KG","GR","TON","ROLLO"].map(unit => `<option>${esc(unit)}</option>`).join("");
      const rows = requisitions || [];
      $("reqListTable").innerHTML = `<thead><tr><th>Folio</th><th>Fecha</th><th>Equipo</th><th>Estatus</th><th>Partidas</th><th>Acciones</th></tr></thead><tbody>` +
        rows.map(row => `<tr data-req-id="${row.id}" style="cursor:pointer"><td>${esc(row.folio)}</td><td>${esc(row.request_date)}</td><td>${esc(row.equipment)}</td><td>${esc(row.status)}</td><td>${num(row.items_count)}</td><td><button type="button" class="btn danger small" data-req-delete="${row.id}" data-req-folio="${esc(row.folio)}">Eliminar</button></td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-req-id]").forEach(row => row.addEventListener("click", () => loadReq(row.dataset.reqId).catch(showError)));
      document.querySelectorAll("[data-req-delete]").forEach(btn => btn.addEventListener("click", event => {
        event.stopPropagation();
        deleteReqById(btn.dataset.reqDelete, btn.dataset.reqFolio).catch(showError);
      }));
    }
    function renderRequisiciones(){
      renderReqProducts();
      renderReqList();
      renderReqItems();
    }
    function trackingState(row){
      const status = String(row.purchase_status || "").toUpperCase();
      const oc = String(row.purchase_order || "").trim();
      if(status.includes("RECIB") || row.received_date) return {label:"Recibido", cls:"ok"};
      if(oc) return {label:"Con OC", cls:"ok"};
      if(status.includes("CANCEL")) return {label:"Cancelado", cls:"bad"};
      if(status) return {label:status, cls:"warn"};
      return {label:"Sin seguimiento", cls:"bad"};
    }
    function trackingRows(){
      const search = ($("trackSearch").value || "").toUpperCase();
      const status = $("trackStatus").value;
      return (requisitions || []).filter(row => {
        const state = trackingState(row);
        const hasOc = String(row.purchase_order || "").trim() !== "";
        const received = state.label === "Recibido";
        const statusOk = !status || (status === "sin_oc" && !hasOc && !received) || (status === "con_oc" && hasOc && !received) || (status === "recibido" && received);
        const text = [row.folio,row.equipment,row.status,row.purchase_status,row.purchase_order,row.supplier,row.buyer,row.tracking_notes].join(" ").toUpperCase();
        return statusOk && (!search || text.includes(search));
      });
    }
    function renderTracking(){
      const rows = trackingRows();
      const total = requisitions.length || 0;
      const withOc = (requisitions || []).filter(row => String(row.purchase_order || "").trim()).length;
      const received = (requisitions || []).filter(row => trackingState(row).label === "Recibido").length;
      $("trackSummary").textContent = `${rows.length} visibles | ${withOc}/${total} con OC | ${received} recibido(s)`;
      $("trackTable").innerHTML = `<thead><tr><th>Estado</th><th>Folio</th><th>Req.</th><th>Equipo</th><th>OC</th><th>Proveedor</th><th>Promesa</th><th>Actualizado</th><th>Acciones</th></tr></thead><tbody>` +
        rows.map(row => {
          const state = trackingState(row);
          return `<tr data-track-id="${row.id}" style="cursor:pointer"><td><span class="pill ${state.cls}">${esc(state.label)}</span></td><td>${esc(row.folio)}</td><td>${esc(row.request_date)}</td><td>${esc(row.equipment)}</td><td>${esc(row.purchase_order)}</td><td>${esc(row.supplier)}</td><td>${esc(row.expected_date || row.received_date || "")}</td><td>${esc(row.tracking_updated_at || "")}</td><td><button type="button" class="btn danger small" data-track-delete="${row.id}" data-track-folio="${esc(row.folio)}">Eliminar</button></td></tr>`;
        }).join("") + `</tbody>`;
      document.querySelectorAll("[data-track-id]").forEach(row => row.addEventListener("click", () => loadTrack(row.dataset.trackId).catch(showError)));
      document.querySelectorAll("[data-track-delete]").forEach(btn => btn.addEventListener("click", event => {
        event.stopPropagation();
        deleteReqById(btn.dataset.trackDelete, btn.dataset.trackFolio).catch(showError);
      }));
      if(currentTrackId && !rows.some(row => String(row.id) === String(currentTrackId))) clearTrackForm();
    }
    function clearTrackForm(){
      currentTrackId = null;
      currentTrackItems = [];
      ["trackFolio","trackReqDate","trackEquipment","trackPurchaseStatus","trackPurchaseOrder","trackPurchaseOrderDate","trackSupplier","trackBuyer","trackExpectedDate","trackReceivedDate","trackNotes"].forEach(id => $(id).value = "");
      $("trackSelected").textContent = "Selecciona una requisicion.";
      renderTrackItems();
    }
    function setTrackForm(row){
      currentTrackId = row?.id || null;
      currentTrackItems = row?.items || [];
      $("trackFolio").value = row?.folio || "";
      $("trackReqDate").value = row?.request_date || "";
      $("trackEquipment").value = row?.equipment || "";
      $("trackPurchaseStatus").value = row?.purchase_status || "";
      $("trackPurchaseOrder").value = row?.purchase_order || "";
      $("trackPurchaseOrderDate").value = row?.purchase_order_date || "";
      $("trackSupplier").value = row?.supplier || "";
      $("trackBuyer").value = row?.buyer || "";
      $("trackExpectedDate").value = row?.expected_date || "";
      $("trackReceivedDate").value = row?.received_date || "";
      $("trackNotes").value = row?.tracking_notes || "";
      $("trackSelected").textContent = currentTrackId ? `${row.folio} | ${row.items_count || currentTrackItems.length} partida(s)` : "Selecciona una requisicion.";
      renderTrackItems();
    }
    function renderTrackItems(){
      $("trackItemsTable").innerHTML = `<thead><tr><th>Cant.</th><th>Unidad</th><th>No. parte</th><th>Descripcion</th></tr></thead><tbody>` +
        (currentTrackItems || []).map(item => `<tr><td>${num(item.quantity)}</td><td>${esc(item.unit)}</td><td>${esc(item.part_number)}</td><td>${esc(item.description)}</td></tr>`).join("") + `</tbody>`;
    }
    async function loadTrack(rowId){
      const r = await fetch(`/api/requisitions/${rowId}`, {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      const payload = await r.json();
      setTrackForm(payload.requisition);
    }
    async function saveTrack(){
      if(!currentTrackId) return alert("Selecciona una requisicion.");
      if(!hasApiKey(true)) return;
      const base = (requisitions || []).find(row => String(row.id) === String(currentTrackId)) || {};
      const payload = {
        ...base,
        id: currentTrackId,
        purchase_status: $("trackPurchaseStatus").value,
        purchase_order: $("trackPurchaseOrder").value,
        purchase_order_date: $("trackPurchaseOrderDate").value,
        supplier: $("trackSupplier").value,
        buyer: $("trackBuyer").value,
        expected_date: $("trackExpectedDate").value,
        received_date: $("trackReceivedDate").value,
        tracking_notes: $("trackNotes").value,
        items: currentTrackItems
      };
      const r = await fetch("/api/requisitions", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!r.ok) return alert(await apiError(r));
      const saved = await r.json();
      await load();
      setTrackForm(saved.requisition);
    }
    async function importTracking(){
      const file = $("trackImportFile").files[0]; if(!file) return alert("Selecciona un Excel .xlsx o .xlsm.");
      if(!hasApiKey(true)) return;
      const dataUrl = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload=()=>res(fr.result); fr.onerror=rej; fr.readAsDataURL(file); });
      const r = await fetch("/api/requisition-tracking/import", {method:"POST", headers:headers(true), body:JSON.stringify({file_name:file.name, data:String(dataUrl), create_missing:$("trackCreateMissing").checked})});
      const payload = await r.json().catch(() => ({}));
      $("trackImportResult").textContent = JSON.stringify(payload, null, 2);
      if(!r.ok) return alert(payload.detail || "No se pudo importar seguimiento.");
      requisitions = payload.requisitions || requisitions;
      renderTracking();
    }
    async function exportTracking(){
      if(!hasApiKey(true)) return;
      const r = await fetch("/api/requisition-tracking/export", {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      const blob = await r.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `MTTO_PROVIDENCIA_SEGUIMIENTO_REQ_${toIsoDate(new Date())}.xlsx`;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1500);
    }
    function openTrackedRequisition(){
      if(!currentTrackId) return alert("Selecciona una requisicion.");
      loadReq(currentTrackId).then(() => {
        document.querySelector('[data-tab="requisiciones"]').click();
      }).catch(showError);
    }
    function hoseEquipmentOptions(){
      const set = new Set();
      portalEquipment().forEach(eq => { if(eq.code || eq.equipment_code) set.add(String(eq.code || eq.equipment_code).trim().toUpperCase()); });
      (hoses.records || []).forEach(row => { if(row.equipment) set.add(String(row.equipment).trim().toUpperCase()); });
      return [...set].filter(Boolean).sort();
    }
    function renderHoseSelectors(){
      const codes = hoseEquipmentOptions();
      const currentFilter = $("hoseFilterEquipment").value || "Todos";
      $("hoseFilterEquipment").innerHTML = `<option>Todos</option>` + codes.map(code => `<option>${esc(code)}</option>`).join("");
      $("hoseFilterEquipment").value = codes.includes(currentFilter) || currentFilter === "Todos" ? currentFilter : "Todos";
      const currentEquipment = $("hoseEquipment").value || "";
      $("hoseEquipment").innerHTML = `<option value=""></option>` + codes.map(code => `<option>${esc(code)}</option>`).join("");
      if(codes.includes(currentEquipment)) $("hoseEquipment").value = currentEquipment;
    }
    function filteredHoseRecords(){
      const selected = $("hoseFilterEquipment").value || "Todos";
      if(selected === "Todos") return hoses.records || [];
      return (hoses.records || []).filter(row => String(row.equipment || "").toUpperCase() === String(selected).toUpperCase());
    }
    function renderHoses(){
      $("hoseBase").value = $("hoseBase").value || toIsoDate(new Date());
      $("hoseStart").value = hoses.start || $("hoseStart").value || toIsoDate(new Date());
      $("hoseEnd").value = hoses.end || $("hoseEnd").value || $("hoseStart").value;
      $("hoseDate").value = $("hoseDate").value || toIsoDate(new Date());
      renderHoseSelectors();
      const totals = hoses.totals || {};
      $("hosePeriodLabel").textContent = `${hoses.start || $("hoseStart").value} a ${hoses.end || $("hoseEnd").value} | ${hoses.period_days || 0} dias`;
      $("hoseStats").innerHTML = [
        ["Cambios", totals.changes || 0],
        ["Consumo real", `${one(totals.real_qty)} pzas`],
        ["Consumo aprox.", `${one(totals.estimated_qty)} pzas`],
        ["Variacion", `${Number(totals.variance || 0) >= 0 ? "+" : ""}${one(totals.variance)} pzas`],
        ["Equipo critico", totals.critical_equipment || "S/D"],
      ].map(([k,v]) => `<div class="stat"><strong>${esc(v)}</strong>${esc(k)}</div>`).join("");
      $("hoseSummaryTable").innerHTML = `<thead><tr><th>Equipo</th><th>Tipo</th><th>Sistemas</th><th>Cambios</th><th>Real</th><th>Aprox.</th><th>Var.</th></tr></thead><tbody>` +
        (hoses.summary || []).map(row => `<tr><td>${esc(row.equipment)}</td><td>${esc(row.part_type)}</td><td>${esc(row.systems)}</td><td>${num(row.changes)}</td><td>${one(row.quantity)}</td><td>${one(row.estimated_qty)}</td><td>${Number(row.variance || 0) >= 0 ? "+" : ""}${one(row.variance)}</td></tr>`).join("") +
        `</tbody>`;
      const records = filteredHoseRecords();
      $("hoseRecordsTable").innerHTML = `<thead><tr><th>Fecha</th><th>Equipo</th><th>Sistema</th><th>Tipo</th><th>Diam.</th><th>Largo</th><th>Cant.</th><th>Causa</th><th>Tecnico</th><th>Accion</th></tr></thead><tbody>` +
        records.map((row, idx) => `<tr data-hose-index="${idx}" style="cursor:pointer"><td>${esc(row.change_date)}</td><td>${esc(row.equipment)}</td><td>${esc(row.system)}</td><td>${esc(row.part_type)}</td><td>${esc(row.diameter)}</td><td>${one(row.length_m)}</td><td>${one(row.quantity)}</td><td>${esc(row.failure_reason)}</td><td>${esc(row.technician)}</td><td><button type="button" class="btn danger small" data-hose-delete="${idx}">Eliminar</button></td></tr>`).join("") +
        `</tbody>`;
      document.querySelectorAll("[data-hose-index]").forEach(tr => tr.addEventListener("click", () => editHoseRecord(Number(tr.dataset.hoseIndex))));
      document.querySelectorAll("[data-hose-delete]").forEach(button => button.addEventListener("click", event => {
        event.stopPropagation();
        deleteHoseRecordAt(Number(button.dataset.hoseDelete)).catch(showError);
      }));
    }
    function clearHoseRecord(){
      currentHoseId = null;
      currentHoseRecord = null;
      $("hoseEditStatus").textContent = "Nuevo cambio";
      $("hoseDate").value = toIsoDate(new Date());
      $("hoseEquipment").value = "";
      $("hoseSystem").value = "HIDRAULICO";
      $("hoseType").value = "MANGUERA";
      ["hoseDiameter","hoseReason","hoseTech","hoseNotes"].forEach(id => $(id).value = "");
      $("hoseLength").value = "0";
      $("hoseQty").value = "1";
      $("hoseLifeDays").value = "30";
      $("hoseEstimatedWeekly").value = "0";
    }
    function editHoseRecord(index){
      const row = filteredHoseRecords()[index];
      if(!row) return;
      currentHoseId = Number(row.id || 0) || null;
      currentHoseRecord = row;
      $("hoseEditStatus").textContent = currentHoseId ? `Editando cambio #${currentHoseId}` : "Editando cambio";
      $("hoseDate").value = row.change_date || "";
      $("hoseEquipment").value = row.equipment || "";
      $("hoseSystem").value = row.system || "HIDRAULICO";
      $("hoseType").value = row.part_type || "MANGUERA";
      $("hoseDiameter").value = row.diameter || "";
      $("hoseLength").value = row.length_m || 0;
      $("hoseQty").value = row.quantity || 1;
      $("hoseLifeDays").value = row.estimated_life_days || 30;
      $("hoseEstimatedWeekly").value = row.estimated_weekly_qty || 0;
      $("hoseReason").value = row.failure_reason || "";
      $("hoseTech").value = row.technician || "";
      $("hoseNotes").value = row.notes || "";
    }
    function hosePayload(){
      return {
        id: currentHoseId,
        change_date: $("hoseDate").value,
        equipment: $("hoseEquipment").value,
        system: $("hoseSystem").value,
        part_type: $("hoseType").value,
        diameter: $("hoseDiameter").value,
        length_m: $("hoseLength").value,
        quantity: $("hoseQty").value,
        unit_cost: 0,
        estimated_life_days: $("hoseLifeDays").value,
        estimated_weekly_qty: $("hoseEstimatedWeekly").value,
        failure_reason: $("hoseReason").value,
        technician: $("hoseTech").value,
        notes: $("hoseNotes").value,
      };
    }
    async function refreshHoses(){
      const params = new URLSearchParams({start:$("hoseStart").value, end:$("hoseEnd").value, equipment:$("hoseFilterEquipment").value || ""});
      const r = await fetch(`/api/hose-changes?${params}`, {headers: headers(), cache:"no-store"});
      if(!r.ok) return alert(await apiError(r));
      hoses = await r.json();
      renderHoses();
    }
    function applyHosePeriod(){
      const [start, end] = periodRange($("hosePeriod").value, $("hoseBase").value || $("hoseStart").value);
      $("hoseStart").value = start;
      $("hoseEnd").value = end;
      refreshHoses().catch(showError);
    }
    async function saveHoseRecord(){
      if(!hasApiKey(true)) return;
      const r = await fetch("/api/hose-changes", {method:"POST", headers:headers(true), body:JSON.stringify(hosePayload())});
      if(!r.ok) return alert(await apiError(r));
      clearHoseRecord();
      await refreshHoses();
    }
    async function deleteHoseRecordAt(index){
      const row = filteredHoseRecords()[index];
      if(!row) return alert("Selecciona un cambio para eliminar.");
      if(!hasApiKey(true)) return;
      if(!confirm(`Se eliminara el cambio de ${row.part_type || "MANGUERA"} en ${row.equipment || ""}.`)) return;
      const r = await fetch("/api/hose-changes/delete", {method:"POST", headers:headers(true), body:JSON.stringify({id:row.id})});
      if(!r.ok) return alert(await apiError(r));
      clearHoseRecord();
      await refreshHoses();
    }
    async function deleteHoseRecord(){
      if(!currentHoseRecord) return alert("Selecciona un cambio para eliminar.");
      if(!hasApiKey(true)) return;
      if(!confirm("Se eliminara el cambio de manguera/conexion seleccionado.")) return;
      const r = await fetch("/api/hose-changes/delete", {method:"POST", headers:headers(true), body:JSON.stringify({id:currentHoseRecord.id})});
      if(!r.ok) return alert(await apiError(r));
      clearHoseRecord();
      await refreshHoses();
    }
    function dieselEquipmentKey(value){
      return String(value || "").trim().toUpperCase().replace(/\([^)]*\)/g, " ").replace(/[^A-Z0-9]+/g, "");
    }
    function dieselEquipmentCodeKey(value){
      const match = String(value || "").trim().toUpperCase().match(/([A-Z]{1,4})[- ]?0*(\d{1,4})/);
      return match ? `${match[1]}${Number(match[2])}` : "";
    }
    function dieselEquipmentCandidates(value){
      const text = String(value || "").trim().toUpperCase();
      if(!text) return [];
      const withoutParentheses = text.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
      const firstToken = text.split(/[\s(]+/)[0].trim();
      const codeMatches = text.match(/[A-Z]{1,4}[- ]?\d{2,4}/g) || [];
      return [...new Set([text, withoutParentheses, firstToken, ...codeMatches].flatMap(item => [dieselEquipmentKey(item), dieselEquipmentCodeKey(item)]).filter(Boolean))];
    }
    function dieselFallbackEquipment(value){
      const text = String(value || "").trim().toUpperCase();
      if(!text) return "";
      const withoutParentheses = text.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
      const firstToken = text.split(/[\s(]+/)[0].trim();
      return firstToken && /\d/.test(firstToken) ? firstToken : (withoutParentheses || text);
    }
    function addDieselEquipmentAlias(aliasMap, canonical, aliases=[]){
      const clean = String(canonical || "").trim().toUpperCase();
      if(!clean) return;
      [clean, ...aliases].forEach(alias => {
        dieselEquipmentCandidates(alias).forEach(key => aliasMap.set(key, clean));
      });
    }
    function dieselBaseEquipmentAliasMap(){
      const aliasMap = new Map();
      [...portalEquipment(), ...(Array.isArray(data.equipment) ? data.equipment : [])].forEach(eq => {
        const code = String(eq.code || eq.equipment_code || "").trim().toUpperCase();
        const description = String(eq.description || eq.family || "").trim().toUpperCase();
        if(code) addDieselEquipmentAlias(aliasMap, code, [description ? `${code} ${description}` : "", description ? `${code} (${description})` : ""]);
      });
      (diesel.equipment || []).forEach(value => {
        const canonical = dieselCanonicalEquipment(value, aliasMap) || dieselFallbackEquipment(value);
        addDieselEquipmentAlias(aliasMap, canonical, [value]);
      });
      return aliasMap;
    }
    function dieselCanonicalEquipment(value, aliasMap=dieselBaseEquipmentAliasMap()){
      const text = String(value || "").trim().toUpperCase();
      if(!text) return "";
      for(const key of dieselEquipmentCandidates(text)){
        if(aliasMap.has(key)) return aliasMap.get(key);
      }
      return dieselFallbackEquipment(text);
    }
    function dieselEquipmentOptions(){
      const aliasMap = dieselBaseEquipmentAliasMap();
      const set = new Set();
      function addOption(value){
        const canonical = dieselCanonicalEquipment(value, aliasMap);
        if(!canonical) return;
        addDieselEquipmentAlias(aliasMap, canonical, [value]);
        set.add(canonical);
      }
      (diesel.equipment || []).forEach(addOption);
      portalEquipment().forEach(eq => addOption(eq.code || eq.equipment_code || ""));
      (diesel.records || []).forEach(row => addOption(row.equipment));
      return [...set].sort();
    }
    function renderDieselSelectors(){
      const codes = dieselEquipmentOptions();
      const rawFilter = $("dieselFilterEquipment").value || "Todos";
      const currentFilter = rawFilter === "Todos" ? "Todos" : dieselCanonicalEquipment(rawFilter);
      $("dieselFilterEquipment").innerHTML = `<option>Todos</option>` + codes.map(code => `<option>${esc(code)}</option>`).join("");
      $("dieselFilterEquipment").value = codes.includes(currentFilter) || currentFilter === "Todos" ? currentFilter : "Todos";
      const currentEquipment = dieselCanonicalEquipment($("dieselEquipment").value);
      $("dieselEquipment").innerHTML = `<option value=""></option>` + codes.map(code => `<option>${esc(code)}</option>`).join("");
      if(codes.includes(currentEquipment)) $("dieselEquipment").value = currentEquipment;
    }
    function dieselRendText(value){ return value == null || value === "" ? "S/H" : one(value); }
    function dieselStatusClass(status){
      const text = String(status || "").toUpperCase();
      if(text.includes("ALTO") || text.includes("SIN HORAS")) return "bad";
      if(text.includes("SIN CONSUMO")) return "warn";
      return "ok";
    }
    function groupedDieselRows(rows){
      const aliasMap = dieselBaseEquipmentAliasMap();
      const bitacoraHours = dieselBitacoraHoursByEquipment(diesel.start || $("dieselStart").value, diesel.end || $("dieselEnd").value, aliasMap);
      const grouped = new Map();
      (rows || []).forEach(raw => {
        const equipment = dieselCanonicalEquipment(raw.equipment, aliasMap);
        if(!equipment) return;
        addDieselEquipmentAlias(aliasMap, equipment, [raw.equipment]);
        if(!grouped.has(equipment)){
          grouped.set(equipment, {equipment, condition:raw.condition || "DISPONIBLE", hi_values:[], hf_values:[], worked_hours:0, diesel_liters:0});
        }
        const row = grouped.get(equipment);
        row.condition = raw.condition || row.condition;
        const hi = Number(raw.horometer_initial || 0);
        const hf = Number(raw.horometer_final || 0);
        if(hi > 0) row.hi_values.push(hi);
        if(hf > 0) row.hf_values.push(hf);
        row.worked_hours += Number(raw.worked_hours || 0);
        row.diesel_liters += Number(raw.diesel_liters || 0);
      });
      const meta = Number($("dieselMeta").value || diesel.meta_lh || 25);
      return [...grouped.values()].map(row => {
        const merged = dieselApplyBitacoraHours(row, bitacoraHours);
        const rendimiento = merged.worked_hours > 0 ? merged.diesel_liters / merged.worked_hours : null;
        let status = "OK";
        if(merged.diesel_liters <= 0) status = "SIN CONSUMO";
        else if(merged.worked_hours <= 0) status = "SIN HORAS";
        else if(rendimiento !== null && rendimiento > meta) status = "ALTO";
        return {
          equipment: merged.equipment,
          condition: merged.condition,
          horometer_initial: merged.horometer_initial,
          horometer_final: merged.horometer_final,
          worked_hours: merged.worked_hours,
          hours_source: merged.hours_source,
          diesel_liters: merged.diesel_liters,
          rendimiento_lh: rendimiento,
          status,
        };
      }).sort((a,b) => b.diesel_liters - a.diesel_liters || a.equipment.localeCompare(b.equipment));
    }
    function filteredDieselRows(){
      const selected = $("dieselFilterEquipment").value || "Todos";
      const rows = groupedDieselRows(diesel.rows || []);
      if(selected === "Todos") return rows;
      const selectedKey = dieselEquipmentKey(dieselCanonicalEquipment(selected));
      return rows.filter(row => dieselEquipmentKey(row.equipment) === selectedKey);
    }
    function filteredDieselRecords(){
      const selected = $("dieselFilterEquipment").value || "Todos";
      const rows = diesel.records || [];
      if(selected === "Todos") return rows;
      const aliasMap = dieselBaseEquipmentAliasMap();
      const selectedKey = dieselEquipmentKey(dieselCanonicalEquipment(selected, aliasMap));
      return rows.filter(row => dieselEquipmentKey(dieselCanonicalEquipment(row.equipment, aliasMap)) === selectedKey);
    }
    function dieselDayOwner(row){
      const text = String(row?.supplier || "").toUpperCase();
      return text.includes("PROSERMIN") ? "PROSERMIN" : "MGA";
    }
    function dieselSupplierTotals(records){
      const ownerByDate = {};
      (diesel.days || []).forEach(row => {
        if(row.work_date) ownerByDate[row.work_date] = row.supplier_owner || dieselDayOwner(row);
      });
      return (records || []).reduce((acc, row) => {
        const owner = ownerByDate[row.work_date] || "MGA";
        const liters = Number(row.diesel_liters || 0);
        if(owner === "PROSERMIN") acc.prosermin_liters += liters;
        else acc.mga_liters += liters;
        return acc;
      }, {mga_liters:0, prosermin_liters:0});
    }
    function dieselInventoryTotals(start, end){
      const rows = [...(diesel.days || [])].filter(row => !start || inRange(row.work_date, start, end)).sort((a,b) => String(a.work_date || "").localeCompare(String(b.work_date || "")));
      const candidates = rows.filter(row => Number(row.final_stock || 0) > 0 || Number(row.initial_stock || 0) > 0 || Number(row.diesel_received || 0) > 0 || Number(row.diesel_liters || 0) > 0);
      const last = candidates[candidates.length - 1] || {};
      const totalStock = Number(last.final_stock || 0);
      let proserminStock = Number(last.prosermin_stock || 0);
      if(!proserminStock && diesel.totals) proserminStock = Number(diesel.totals.prosermin_stock || 0);
      proserminStock = Math.max(Math.min(proserminStock, totalStock), 0);
      return {
        total_stock: totalStock,
        mga_stock: Math.max(totalStock - proserminStock, 0),
        prosermin_stock: proserminStock,
      };
    }
    function dieselTotalsForRows(rows, records){
      const liters = rows.reduce((sum,row) => sum + Number(row.diesel_liters || 0), 0);
      const hours = rows.reduce((sum,row) => sum + Number(row.worked_hours || 0), 0);
      const split = dieselSupplierTotals(records || []);
      const stock = dieselInventoryTotals(diesel.start || $("dieselStart").value, diesel.end || $("dieselEnd").value);
      return {
        diesel_liters: liters,
        worked_hours: hours,
        rendimiento_lh: hours > 0 ? liters / hours : null,
        critical: rows.filter(row => ["ALTO","SIN HORAS"].includes(String(row.status || "").toUpperCase())).length,
        mga_liters: split.mga_liters,
        prosermin_liters: split.prosermin_liters,
        total_stock: stock.total_stock,
        mga_stock: stock.mga_stock,
        prosermin_stock: stock.prosermin_stock,
      };
    }
    function dieselDayConsumption(date){
      const day = (diesel.days || []).find(row => row.work_date === date);
      if(day) return Number(day.diesel_liters || 0);
      return (diesel.records || [])
        .filter(row => row.work_date === date)
        .reduce((sum, row) => sum + Number(row.diesel_liters || 0), 0);
    }
    function hasDieselTankEntry(row){
      return Number(row?.diesel_received || 0) > 0 ||
        Number(row?.initial_stock || 0) > 0 ||
        Number(row?.final_stock || 0) > 0 ||
        Boolean(String(row?.supplier || row?.notes || "").trim());
    }
    function dieselPreviousFinalStock(date, fallbackInitial=Number($("dieselInitial").value || 0)){
      const previous = (diesel.days || [])
        .filter(row => row.work_date && row.work_date < date && hasDieselTankEntry(row))
        .sort((a,b) => String(b.work_date).localeCompare(String(a.work_date)))[0];
      return previous ? Number(previous.final_stock || 0) : fallbackInitial;
    }
    function setDieselNumberInput(id, value){
      $(id).value = Number.isFinite(value) ? one(Math.max(value, 0)) : "0.0";
    }
    function updateDieselTankCalc(usePreviousInitial=true, fallbackInitial=Number($("dieselInitial").value || 0)){
      const date = $("dieselDayDate").value;
      if(!date) return;
      if(usePreviousInitial) setDieselNumberInput("dieselInitial", dieselPreviousFinalStock(date, fallbackInitial));
      const initial = Number($("dieselInitial").value || 0);
      const received = Number($("dieselReceived").value || 0);
      const consumed = dieselDayConsumption(date);
      setDieselNumberInput("dieselFinal", initial + received - consumed);
    }
    function renderDiesel(){
      renderDieselSelectors();
      const rows = filteredDieselRows();
      const records = filteredDieselRecords();
      const totals = dieselTotalsForRows(rows, records);
      $("dieselTitle").textContent = `Rendimiento diesel | ${diesel.start || $("dieselStart").value} a ${diesel.end || $("dieselEnd").value}`;
      $("dieselUpdated").textContent = diesel.updated_at ? `Actualizado ${diesel.updated_at}` : "";
      $("dieselStats").innerHTML = [
        ["Consumo diesel", `${one(totals.diesel_liters)} L`],
        ["Remanente total", `${one(totals.total_stock)} L`],
        ["Diesel MGA", `${one(totals.mga_stock)} L`],
        ["Diesel PROSERMIN", `${one(totals.prosermin_stock)} L`],
        ["Horas trabajadas", `${one(totals.worked_hours)} h`],
        ["Rendimiento prom.", dieselRendText(totals.rendimiento_lh) + (totals.rendimiento_lh == null ? "" : " L/H")],
        ["Equipos revision", totals.critical],
        ["Llegada diesel", `${one((diesel.totals || {}).received || 0)} L`],
      ].map(([k,v]) => `<div class="stat"><strong>${esc(v)}</strong>${esc(k)}</div>`).join("");
      $("dieselReportTable").innerHTML = `<thead><tr><th>Equipo</th><th>Condicion actual</th><th>Horometro inicial</th><th>Horometro final</th><th>Horas trabajadas</th><th>Consumo diesel</th><th>Rendimiento L/H</th><th>KPI</th></tr></thead><tbody>` +
        rows.map(row => `<tr><td>${esc(row.equipment)}</td><td>${esc(row.condition)}</td><td>${one(row.horometer_initial)}</td><td>${one(row.horometer_final)}</td><td>${one(row.worked_hours)}</td><td>${one(row.diesel_liters)}</td><td>${dieselRendText(row.rendimiento_lh)}</td><td><span class="pill ${dieselStatusClass(row.status)}">${esc(row.status)}</span></td></tr>`).join("") +
        `</tbody>`;
      $("dieselDailyTable").innerHTML = `<thead><tr><th>Fecha</th><th>Consumo L</th><th>Origen</th><th>Llegada L</th><th>Inicial L</th><th>Final L</th><th>PROSERMIN disp.</th><th>Proveedor</th><th>Accion</th></tr></thead><tbody>` +
        (diesel.days || []).map((row, idx) => `<tr><td>${esc(row.work_date)}</td><td>${one(row.diesel_liters)}</td><td>${esc(row.supplier_owner || dieselDayOwner(row))}</td><td>${one(row.diesel_received)}</td><td>${one(row.initial_stock)}</td><td>${one(row.final_stock)}</td><td>${one(row.prosermin_stock)}</td><td>${esc(row.supplier)}</td><td><button type="button" class="btn danger small" data-diesel-day-delete="${idx}">Eliminar</button></td></tr>`).join("") +
        `</tbody>`;
      const recordAliasMap = dieselBaseEquipmentAliasMap();
      $("dieselRecordsTable").innerHTML = `<thead><tr><th>Fecha</th><th>Equipo</th><th>Turno</th><th>HI</th><th>HF</th><th>Hrs</th><th>Diesel L</th><th>Origen</th><th>Accion</th></tr></thead><tbody>` +
        records.map((row, idx) => `<tr data-diesel-index="${idx}" style="cursor:pointer"><td>${esc(row.work_date)}</td><td>${esc(dieselCanonicalEquipment(row.equipment, recordAliasMap))}</td><td>${esc(row.shift)}</td><td>${one(row.horometer_initial)}</td><td>${one(row.horometer_final)}</td><td>${one(row.worked_hours)}</td><td>${one(row.diesel_liters)}</td><td>${esc(row.source || "desktop")}</td><td><button type="button" class="btn danger small" data-diesel-record-delete="${idx}">Eliminar</button></td></tr>`).join("") +
        `</tbody>`;
      document.querySelectorAll("[data-diesel-index]").forEach(tr => tr.addEventListener("click", () => editDieselRecord(Number(tr.dataset.dieselIndex))));
      document.querySelectorAll("[data-diesel-record-delete]").forEach(button => button.addEventListener("click", event => {
        event.stopPropagation();
        deleteDieselRecordAt(Number(button.dataset.dieselRecordDelete)).catch(showError);
      }));
      document.querySelectorAll("[data-diesel-day-delete]").forEach(button => button.addEventListener("click", event => {
        event.stopPropagation();
        deleteDieselDayAt(Number(button.dataset.dieselDayDelete)).catch(showError);
      }));
    }
    function clearDieselRecord(){
      currentDieselId = null;
      currentDieselRecord = null;
      $("dieselEditStatus").textContent = "Nueva captura";
      $("dieselDate").value = toIsoDate(new Date());
      $("dieselEquipment").value = "";
      $("dieselCondition").value = "DISPONIBLE";
      $("dieselShift").value = "1";
      ["dieselHi","dieselHf","dieselHours","dieselLiters"].forEach(id => $(id).value = "0");
      ["dieselOperator","dieselDispatcher","dieselSupervisor","dieselNotes"].forEach(id => $(id).value = "");
    }
    function editDieselRecord(index){
      const row = filteredDieselRecords()[index];
      if(!row) return;
      currentDieselId = row.source === "web" ? Number(row.id || 0) || null : null;
      currentDieselRecord = row;
      $("dieselEditStatus").textContent = currentDieselId ? `Editando captura web #${currentDieselId}` : "Editando captura sincronizada";
      $("dieselDate").value = row.work_date || "";
      $("dieselEquipment").value = dieselCanonicalEquipment(row.equipment) || "";
      $("dieselCondition").value = row.condition || "DISPONIBLE";
      $("dieselShift").value = row.shift || "1";
      $("dieselHi").value = row.horometer_initial || 0;
      $("dieselHf").value = row.horometer_final || 0;
      $("dieselHours").value = row.worked_hours || 0;
      $("dieselLiters").value = row.diesel_liters || 0;
      $("dieselOperator").value = row.operator || "";
      $("dieselDispatcher").value = row.dispatcher || "";
      $("dieselSupervisor").value = row.supervisor || "";
      $("dieselNotes").value = row.notes || "";
    }
    function dieselRecordPayload(){
      return {
        id: currentDieselId,
        work_date: $("dieselDate").value,
        equipment: $("dieselEquipment").value,
        condition: $("dieselCondition").value,
        shift: $("dieselShift").value,
        horometer_initial: $("dieselHi").value,
        horometer_final: $("dieselHf").value,
        worked_hours: $("dieselHours").value,
        diesel_liters: $("dieselLiters").value,
        operator: $("dieselOperator").value,
        dispatcher: $("dieselDispatcher").value,
        supervisor: $("dieselSupervisor").value,
        notes: $("dieselNotes").value,
      };
    }
    async function refreshDiesel(){
      const params = new URLSearchParams({start:$("dieselStart").value, end:$("dieselEnd").value, meta_lh:$("dieselMeta").value || "25"});
      const r = await fetch(`/api/diesel?${params}`, {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      diesel = await r.json();
      renderDiesel();
    }
    function applyDieselPeriod(){
      const [start, end] = periodRange($("dieselPeriod").value, $("dieselBase").value || $("dieselStart").value);
      $("dieselStart").value = start;
      $("dieselEnd").value = end;
      refreshDiesel().catch(showError);
    }
    async function saveDieselRecord(){
      if(!hasApiKey(true)) return;
      const r = await fetch("/api/diesel/records", {method:"POST", headers:headers(true), body:JSON.stringify(dieselRecordPayload())});
      if(!r.ok) return alert(await apiError(r));
      clearDieselRecord();
      await refreshDiesel();
    }
    function dieselRecordDeletePayload(row){
      return {
        id: Number(row?.id || 0) || "",
        work_date: row?.work_date || "",
        equipment: dieselCanonicalEquipment(row?.equipment || "") || row?.equipment || "",
        shift: row?.shift || "1",
      };
    }
    async function deleteDieselRecordAt(index){
      const row = filteredDieselRecords()[index];
      if(!row) return alert("Selecciona una captura diesel para eliminar.");
      if(!hasApiKey(true)) return;
      const label = `${row.work_date || ""} ${dieselCanonicalEquipment(row.equipment || "") || row.equipment || ""} turno ${row.shift || "1"}`.trim();
      if(!confirm(`Se eliminara la captura diesel ${label}.`)) return;
      const r = await fetch("/api/diesel/records/delete", {method:"POST", headers:headers(true), body:JSON.stringify(dieselRecordDeletePayload(row))});
      if(!r.ok) return alert(await apiError(r));
      clearDieselRecord();
      await refreshDiesel();
    }
    async function deleteDieselRecord(){
      if(!currentDieselRecord) return alert("Selecciona una captura diesel para eliminar.");
      if(!hasApiKey(true)) return;
      if(!confirm("Se eliminara la captura diesel seleccionada.")) return;
      const r = await fetch("/api/diesel/records/delete", {method:"POST", headers:headers(true), body:JSON.stringify(dieselRecordDeletePayload(currentDieselRecord))});
      if(!r.ok) return alert(await apiError(r));
      clearDieselRecord();
      await refreshDiesel();
    }
    async function saveDieselDay(){
      if(!hasApiKey(true)) return;
      updateDieselTankCalc(true);
      const payload = {
        work_date:$("dieselDayDate").value,
        diesel_received:$("dieselReceived").value,
        initial_stock:$("dieselInitial").value,
        final_stock:$("dieselFinal").value,
        prosermin_stock:$("dieselProserminStock").value,
        supplier:$("dieselSupplier").value,
        notes:$("dieselDayNotes").value,
      };
      const r = await fetch("/api/diesel/days", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!r.ok) return alert(await apiError(r));
      await refreshDiesel();
    }
    async function deleteDieselDayAt(index){
      const row = (diesel.days || [])[index];
      if(!row || !row.work_date) return alert("Selecciona un dia para eliminar.");
      if(!hasApiKey(true)) return;
      if(!confirm(`Se eliminara el consumo diario del ${row.work_date}.`)) return;
      const r = await fetch("/api/diesel/days/delete", {method:"POST", headers:headers(true), body:JSON.stringify({work_date:row.work_date})});
      if(!r.ok) return alert(await apiError(r));
      if($("dieselDayDate").value === row.work_date){
        $("dieselReceived").value = "0";
        $("dieselInitial").value = "0";
        $("dieselFinal").value = "0";
        $("dieselProserminStock").value = "0";
        $("dieselSupplier").value = "";
        $("dieselDayNotes").value = "";
      }
      await refreshDiesel();
    }
    async function downloadDieselExcel(){
      const params = new URLSearchParams({start:$("dieselStart").value, end:$("dieselEnd").value, meta_lh:$("dieselMeta").value || "25"});
      const r = await fetch(`/api/diesel/export?${params}`, {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      const blob = await r.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `Control_Diesel_${$("dieselStart").value}_${$("dieselEnd").value}.xlsx`;
      a.click();
      URL.revokeObjectURL(a.href);
    }
    async function downloadMonthlyPowerPoint(){
      const month = $("monthlyMonth").value || String((new Date()).getMonth() + 1);
      const year = $("monthlyYear").value || String((new Date()).getFullYear());
      $("monthlyStatus").textContent = "Generando...";
      const params = new URLSearchParams({month, year});
      const r = await fetch(`/api/monthly-report/powerpoint?${params}`, {headers: headers(), cache:"no-store"});
      if(!r.ok){
        $("monthlyStatus").textContent = "";
        return alert(await apiError(r));
      }
      const blob = await r.blob();
      const a = document.createElement("a");
      const label = monthNames[Number(month) - 1] || "Mes";
      a.href = URL.createObjectURL(blob);
      a.download = `Reporte_Mensual_${label}_${year}.pptx`;
      a.click();
      URL.revokeObjectURL(a.href);
      $("monthlyStatus").textContent = "PowerPoint generado";
    }
    function applyWeeklyPeriod(updateStatus=true){
      const base = $("weeklyBase").value || $("weeklyStart").value || toIsoDate(new Date());
      const [start, end] = periodRange("Semana", base);
      $("weeklyStart").value = start;
      $("weeklyEnd").value = end;
      if(updateStatus) $("weeklyStatus").textContent = `Semana ${start} a ${end}`;
    }
    async function downloadWeeklyPowerPoint(){
      if(!$("weeklyStart").value || !$("weeklyEnd").value) applyWeeklyPeriod(false);
      const start = $("weeklyStart").value;
      const end = $("weeklyEnd").value;
      $("weeklyStatus").textContent = "Generando reporte semanal...";
      const params = new URLSearchParams({start, end});
      const r = await fetch(`/api/weekly-report/powerpoint?${params}`, {headers: headers(), cache:"no-store"});
      if(!r.ok){
        $("weeklyStatus").textContent = "";
        return alert(await apiError(r));
      }
      const blob = await r.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `Reporte_Semanal_${start}_${end}.pptx`;
      a.click();
      URL.revokeObjectURL(a.href);
      $("weeklyStatus").textContent = "PowerPoint semanal generado";
    }
    function renderAll(){
      renderStats();
      renderSelectors();
      renderPortalSelectors();
      renderExecutiveBoard();
      renderEquipmentProfile();
      renderDashboard();
      renderPreventives();
      renderBacklog();
      renderWorkOrders();
      renderServiceHistory();
      renderPreventiveExecution();
      renderAudit();
      renderBitacora();
      renderCaptureRecent();
      renderDisponibilidad();
      renderRequisiciones();
      renderTracking();
      renderHoses();
      renderDiesel();
      renderTireTracking();
      renderSpareParts();
      renderFilters();
      renderInventory();
      renderMovements();
    }
    document.querySelectorAll(".tabs button").forEach(btn => btn.addEventListener("click", () => {
      activateTab(btn.dataset.tab);
    }));
    ["kpiGroup","kpiStart","kpiEnd"].forEach(id => $(id).addEventListener("change", renderDashboard));
    ["fichaEquipment","fichaStart","fichaEnd"].forEach(id => $(id).addEventListener("change", renderEquipmentProfile));
    $("renderFichaBtn").addEventListener("click", renderEquipmentProfile);
    $("fichaGoCaptureBtn").addEventListener("click", () => {
      const code = selectedFichaCode();
      if(code) $("capEquipment").value = code;
      renderCaptureComponents();
      applyPreviousHi(true);
      activateTab("captura");
    });
    $("fichaGoServiceBtn").addEventListener("click", () => {
      const code = selectedFichaCode();
      resetPreventiveExecutionForm();
      if(code) $("prevExecEquipment").value = code;
      activateTab("ejecucionPreventivos");
    });
    ["kpiSimEnabled","kpiSimName","kpiSimMetaAvailability","kpiSimMetaUtilization","kpiSimMetaReliability","kpiSimMetaTmef","kpiSimMetaTmpr","kpiSimPeriod","kpiSimWorked","kpiSimMp","kpiSimMc","kpiSimStops","kpiSimMission"].forEach(id => {
      const el = $(id);
      if(!el) return;
      el.addEventListener(id === "kpiSimEnabled" ? "change" : "input", renderDashboard);
    });
    $("renderKpiBtn").addEventListener("click", renderDashboard);
    $("printKpiBtn").addEventListener("click", printExactKpi);
    $("kpiImageBtn").addEventListener("click", () => downloadKpiImage().catch(showError));
    $("kpiExcelBtn").addEventListener("click", () => downloadKpiExcel().catch(showError));
    $("meetingModeBtn").addEventListener("click", () => {
      document.body.classList.toggle("meeting-mode");
      $("meetingModeBtn").textContent = document.body.classList.contains("meeting-mode") ? "Salir reunion" : "Modo reunion";
    });
    $("monthlyPptBtn").addEventListener("click", () => downloadMonthlyPowerPoint().catch(showError));
    $("weeklyBase").addEventListener("change", () => applyWeeklyPeriod(true));
    $("weeklyApplyBtn").addEventListener("click", () => applyWeeklyPeriod(true));
    $("weeklyPptBtn").addEventListener("click", () => downloadWeeklyPowerPoint().catch(showError));
    ["prPeriod","prBase","prEquipment"].forEach(id => $(id).addEventListener("change", renderPreventives));
    $("prSearch").addEventListener("input", renderPreventives);
    $("renderPrBtn").addEventListener("click", renderPreventives);
    ["backlogStart","backlogEnd","backlogLevel","backlogSource","backlogStatus"].forEach(id => $(id).addEventListener("change", renderBacklog));
    $("backlogSearch").addEventListener("input", renderBacklog);
    $("renderBacklogBtn").addEventListener("click", renderBacklog);
    ["woFilterEquipment","woFilterStatus","woFilterPriority"].forEach(id => $(id).addEventListener("change", renderWorkOrders));
    $("woSearch").addEventListener("input", renderWorkOrders);
    $("woRefreshBtn").addEventListener("click", renderWorkOrders);
    $("woNewBtn").addEventListener("click", () => newWorkOrder());
    $("woSaveBtn").addEventListener("click", () => saveWorkOrder(false).catch(showError));
    $("woCloseBtn").addEventListener("click", () => saveWorkOrder(true).catch(showError));
    $("woDeleteBtn").addEventListener("click", () => deleteWorkOrder().catch(showError));
    ["srvEquipment","srvInterval","srvType","srvStart","srvEnd"].forEach(id => $(id).addEventListener("change", renderServiceHistory));
    $("srvSearch").addEventListener("input", renderServiceHistory);
    $("renderSrvBtn").addEventListener("click", renderServiceHistory);
    $("prevExecNewBtn").addEventListener("click", resetPreventiveExecutionForm);
    $("prevExecSaveBtn").addEventListener("click", () => savePreventiveExecution(false).catch(showError));
    $("prevExecCloseBtn").addEventListener("click", () => savePreventiveExecution(true).catch(showError));
    $("prevExecDeleteBtn").addEventListener("click", () => deletePreventiveExecution().catch(showError));
    $("prevExecViewHistoryBtn").addEventListener("click", () => document.querySelector('[data-tab="servicios"]')?.click());
    preventiveOilInputs.forEach(item => $(item[0]).addEventListener("input", updatePreventiveOilTotal));
    ["spareEquipment","spareStatus"].forEach(id => $(id).addEventListener("change", renderSpareParts));
    $("spareSearch").addEventListener("input", renderSpareParts);
    $("renderSpareBtn").addEventListener("click", renderSpareParts);
    ["auditStart","auditEnd","auditModule"].forEach(id => $(id).addEventListener("change", renderAudit));
    $("auditSearch").addEventListener("input", renderAudit);
    $("auditRefreshBtn").addEventListener("click", renderAudit);
    ["bitEquipment","bitStart","bitEnd"].forEach(id => $(id).addEventListener("change", renderBitacora));
    $("bitSearch").addEventListener("input", renderBitacora);
    $("renderBitBtn").addEventListener("click", renderBitacora);
    $("capDate").addEventListener("change", () => applyPreviousHi(true));
    $("capShift").addEventListener("change", () => { applyPreviousHi(true); renderCaptureRecent(); });
    $("capEquipment").addEventListener("change", () => { renderCaptureComponents(); applyPreviousHi(true); renderCaptureRecent(); });
    $("capComponent").addEventListener("change", () => applyPreviousHi(true));
    ["capHi","capHf"].forEach(id => $(id).addEventListener("input", updateCaptureWorkedHours));
    ["capOil15w40","capOilHco68","capOilSae30","capOil85w140","capAlmo","capCoolant","capOilVg100","capAtf"].forEach(id => $(id).addEventListener("input", updateCaptureOilTotal));
    $("capNewBtn").addEventListener("click", resetCaptureForm);
    $("capSaveBtn").addEventListener("click", () => saveDailyCapture().catch(showError));
    $("capSaveNewBtn").addEventListener("click", () => saveDailyCapture(true).catch(showError));
    $("capRefreshBtn").addEventListener("click", () => load().catch(showError));
    ["capDate","capEquipment","capComponent","capHi","capHf","capWorked","capMp","capMc"].forEach(id => $(id).addEventListener("input", captureValidationWarnings));
    $("dispSearch").addEventListener("input", renderDisponibilidad);
    $("dispStatus").addEventListener("change", renderDisponibilidad);
    $("renderDispBtn").addEventListener("click", renderDisponibilidad);
    $("reqProductSearch").addEventListener("input", renderReqProducts);
    $("reqNewBtn").addEventListener("click", () => newRequisition());
    $("reqSaveBtn").addEventListener("click", () => saveReq().catch(showError));
    $("reqDeleteBtn").addEventListener("click", () => deleteReq().catch(showError));
    $("reqAddItemBtn").addEventListener("click", addReqItem);
    $("reqClearItemBtn").addEventListener("click", clearReqItem);
    $("reqDeleteItemBtn").addEventListener("click", deleteReqItem);
    $("reqPdfBtn").addEventListener("click", () => reqPdf(false).catch(showError));
    $("reqPrintBtn").addEventListener("click", () => reqPdf(true).catch(showError));
    $("trackSearch").addEventListener("input", renderTracking);
    $("trackStatus").addEventListener("change", renderTracking);
    $("trackImportBtn").addEventListener("click", () => importTracking().catch(showError));
    $("trackExportBtn").addEventListener("click", () => exportTracking().catch(showError));
    $("trackRefreshBtn").addEventListener("click", () => load().catch(showError));
    $("trackSaveBtn").addEventListener("click", () => saveTrack().catch(showError));
    $("trackOpenReqBtn").addEventListener("click", openTrackedRequisition);
    $("hoseApplyPeriodBtn").addEventListener("click", applyHosePeriod);
    $("hoseRefreshBtn").addEventListener("click", () => refreshHoses().catch(showError));
    $("hoseFilterEquipment").addEventListener("change", () => refreshHoses().catch(showError));
    $("hoseNewBtn").addEventListener("click", clearHoseRecord);
    $("hoseSaveBtn").addEventListener("click", () => saveHoseRecord().catch(showError));
    $("hoseDeleteBtn").addEventListener("click", () => deleteHoseRecord().catch(showError));
    $("tireTrackNewBtn").addEventListener("click", resetTireTrackForm);
    $("tireTrackSaveBtn").addEventListener("click", () => saveTireTrackEvent().catch(showError));
    $("tireTrackRefreshBtn").addEventListener("click", () => load().catch(showError));
    $("tireTrackKpiBtn").addEventListener("click", () => {
      $("kpiGroup").value = "KPI Llantas";
      document.querySelector('[data-tab="dashboard"]').click();
      renderDashboard();
    });
    $("tireTrackCode").addEventListener("change", () => {
      const row = tireRows().find(item => String(item.tire_code || "") === String($("tireTrackCode").value || "").trim().toUpperCase());
      if(row) fillTireTrackForm(row);
    });
    $("dieselApplyPeriodBtn").addEventListener("click", applyDieselPeriod);
    $("dieselRefreshBtn").addEventListener("click", () => refreshDiesel().catch(showError));
    $("dieselFilterEquipment").addEventListener("change", renderDiesel);
    $("dieselNewBtn").addEventListener("click", clearDieselRecord);
    $("dieselSaveBtn").addEventListener("click", () => saveDieselRecord().catch(showError));
    $("dieselDeleteBtn").addEventListener("click", () => deleteDieselRecord().catch(showError));
    $("dieselDaySaveBtn").addEventListener("click", () => saveDieselDay().catch(showError));
    $("dieselDayDate").addEventListener("change", () => updateDieselTankCalc(true, 0));
    $("dieselReceived").addEventListener("input", () => updateDieselTankCalc(true));
    $("dieselInitial").addEventListener("input", () => updateDieselTankCalc(false));
    $("dieselPrintBtn").addEventListener("click", () => window.print());
    $("dieselExcelBtn").addEventListener("click", () => downloadDieselExcel().catch(showError));
    ["equipmentSelect","serviceSelect","statusSelect","filterSearch"].forEach(id => {
      const eventName = id.endsWith("Select") ? "change" : "input";
      $(id).addEventListener(eventName, () => { if(id==="equipmentSelect") renderServiceOptions(); renderFilters(); });
    });
    $("inventorySearch").addEventListener("input", renderInventory);
    $("refreshBtn").addEventListener("click", () => load().catch(showError));
    $("exportBtn").addEventListener("click", async () => {
      const r = await fetch("/api/filter-inventory/export", {headers: headers()});
      if(!r.ok) return alert(await apiError(r));
      const blob = await r.blob(); const a = document.createElement("a");
      a.href = URL.createObjectURL(blob); a.download = "Inventario_Filtros_MGA.xlsx"; a.click();
    });
    $("movementBtn").addEventListener("click", async () => {
      if(!hasApiKey(true)) return;
      const payload = { part_number:$("movPart").value, description:$("movDesc").value, movement_type:$("movType").value, quantity:$("movQty").value, unit:$("movUnit").value, equipment_code:$("movEquipment").value, service_interval:$("movService").value, reference:$("movRef").value, created_by:$("movUser").value, notes:$("movNotes").value };
      const r = await fetch("/api/filter-inventory/movement", {method:"POST", headers:headers(true), body:JSON.stringify(payload)});
      if(!r.ok) return alert(await apiError(r));
      ["movPart","movDesc","movRef","movNotes"].forEach(id => $(id).value = "");
      await load();
    });
    $("importBtn").addEventListener("click", async () => {
      const file = $("importFile").files[0]; if(!file) return alert("Selecciona un Excel.");
      if(!hasApiKey(true)) return;
      const dataUrl = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload=()=>res(fr.result); fr.onerror=rej; fr.readAsDataURL(file); });
      const r = await fetch("/api/filter-inventory/import", {method:"POST", headers:headers(true), body:JSON.stringify({file_name:file.name, data:String(dataUrl), replace:true})});
      const payload = await r.json().catch(() => ({}));
      $("importResult").textContent = JSON.stringify(payload, null, 2);
      if(r.ok) await load();
    });
    load().catch(showError);
    setInterval(() => {
      if(document.visibilityState !== "visible") return;
      load().catch(error => console.warn("No se pudo actualizar automatico", error));
    }, AUTO_REFRESH_MS);
  </script>
</body>
</html>"""


@app.get("/almacen-filtros", response_class=HTMLResponse)
def filter_warehouse_page(response: Response) -> str:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return WAREHOUSE_HTML


@app.get("/api/epp")
def get_epp(response: Response, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        return epp_payload(session)


@app.get("/api/epp/snapshot")
def get_epp_snapshot(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        return epp_payload(session)


@app.post("/api/epp/snapshot")
async def replace_epp_snapshot(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    rows = payload.get("items") if isinstance(payload, dict) else []
    if not isinstance(rows, list):
        raise HTTPException(status_code=400, detail="Inventario EPP invalido.")
    movements = payload.get("movements") if isinstance(payload, dict) and isinstance(payload.get("movements"), list) else []
    deliveries = payload.get("deliveries") if isinstance(payload, dict) and isinstance(payload.get("deliveries"), list) else []
    workers = payload.get("workers") if isinstance(payload, dict) and isinstance(payload.get("workers"), list) else None
    with SessionLocal() as session:
        session.query(EppMovement).delete()
        session.query(EppDelivery).delete()
        session.query(EppItem).delete()
        if workers is not None:
            session.query(EppWorker).delete()
        imported = 0
        for row in rows:
            if not isinstance(row, dict) or not normalize_part_key(row.get("code")):
                continue
            upsert_epp_item(
                session,
                code=str(row.get("code") or ""),
                description=str(row.get("description") or ""),
                category=str(row.get("category") or ""),
                size=str(row.get("size") or ""),
                unit=str(row.get("unit") or "PZA"),
                quantity=parse_float(row.get("quantity"), 0),
                min_stock=parse_float(row.get("min_stock"), 0),
                useful_life_days=int(parse_float(row.get("useful_life_days"), 0)),
                risk_area=str(row.get("risk_area") or ""),
                location=str(row.get("location") or ""),
                training_required=int(parse_float(row.get("training_required"), 0)),
                maintenance_notes=str(row.get("maintenance_notes") or ""),
                source_file=str(row.get("source_file") or "desktop-sync"),
            )
            imported += 1
        session.flush()
        item_by_code = {item.code_key: item for item in session.scalars(select(EppItem)).all()}
        for row in movements:
            if not isinstance(row, dict):
                continue
            key = normalize_part_key(row.get("item_code") or row.get("code"))
            item = item_by_code.get(key)
            if item is None:
                continue
            session.add(
                EppMovement(
                    item_id=item.id,
                    movement_date=str(row.get("movement_date") or utc_now().date().isoformat()),
                    movement_type=normalize_text(row.get("movement_type") or "ENTRADA")[:20],
                    quantity=parse_float(row.get("quantity"), 0),
                    balance_after=parse_float(row.get("balance_after"), 0),
                    worker_name=normalize_text(row.get("worker_name"))[:180],
                    employee_id=normalize_text(row.get("employee_id"))[:80],
                    area=normalize_text(row.get("area"))[:180],
                    reference=str(row.get("reference") or "")[:180].upper(),
                    notes=str(row.get("notes") or "").upper(),
                    created_at=utc_now(),
                )
            )
        for row in deliveries:
            if not isinstance(row, dict):
                continue
            key = normalize_part_key(row.get("item_code") or row.get("code"))
            item = item_by_code.get(key)
            if item is None:
                continue
            session.add(
                EppDelivery(
                    item_id=item.id,
                    delivery_date=str(row.get("delivery_date") or utc_now().date().isoformat()),
                    worker_name=normalize_text(row.get("worker_name"))[:180],
                    employee_id=normalize_text(row.get("employee_id"))[:80],
                    area=normalize_text(row.get("area"))[:180],
                    quantity=parse_float(row.get("quantity"), 0),
                    useful_life_days=max(int(parse_float(row.get("useful_life_days"), 0)), 0),
                    due_date=str(row.get("due_date") or "")[:20],
                    received_by=normalize_text(row.get("received_by"))[:180],
                    signature=normalize_text(row.get("signature"))[:180],
                    training_done=1 if row.get("training_done") else 0,
                    condition_status=normalize_text(row.get("condition_status") or "ENTREGADO")[:80],
                    notes=str(row.get("notes") or "").upper(),
                    created_at=utc_now(),
                )
            )
        if workers is not None:
            for row in workers:
                if not isinstance(row, dict):
                    continue
                if not str(row.get("worker_name") or row.get("employee_id") or "").strip():
                    continue
                upsert_epp_worker(
                    session,
                    worker_name=str(row.get("worker_name") or row.get("employee_id") or ""),
                    employee_id=str(row.get("employee_id") or ""),
                    area=str(row.get("area") or ""),
                    position=str(row.get("position") or ""),
                    active=1 if row.get("active", 1) else 0,
                )
        session.commit()
        return {"ok": True, "imported": imported}


@app.post("/api/epp/items")
async def save_epp_item(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="EPP invalido.")
    with SessionLocal() as session:
        item = upsert_epp_item(
            session,
            code=str(payload.get("code") or ""),
            description=str(payload.get("description") or ""),
            category=str(payload.get("category") or ""),
            size=str(payload.get("size") or ""),
            unit=str(payload.get("unit") or "PZA"),
            quantity=parse_float(payload.get("quantity"), 0),
            min_stock=parse_float(payload.get("min_stock"), 0),
            useful_life_days=int(parse_float(payload.get("useful_life_days"), 0)),
            risk_area=str(payload.get("risk_area") or ""),
            location=str(payload.get("location") or ""),
            training_required=1 if payload.get("training_required") else 0,
            maintenance_notes=str(payload.get("maintenance_notes") or ""),
            source_file="web",
        )
        session.commit()
        session.refresh(item)
        return {"ok": True, "item": epp_item_payload(item)}


@app.post("/api/epp/items/delete")
async def delete_epp_item(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    code = str(payload.get("code") if isinstance(payload, dict) else "").strip()
    key = normalize_part_key(code)
    if not key:
        raise HTTPException(status_code=400, detail="Codigo EPP requerido.")
    with SessionLocal() as session:
        item = session.scalar(select(EppItem).where(EppItem.code_key == key))
        if item is None:
            raise HTTPException(status_code=404, detail="EPP no encontrado.")
        session.delete(item)
        session.commit()
        return {"ok": True, "deleted": code}


@app.post("/api/epp/delete-all")
async def delete_all_epp(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict) or payload.get("confirm") != "ELIMINAR":
        raise HTTPException(status_code=400, detail="Confirmacion requerida.")
    with SessionLocal() as session:
        count = session.query(EppItem).count()
        session.query(EppMovement).delete()
        session.query(EppDelivery).delete()
        session.query(EppItem).delete()
        session.commit()
        return {"ok": True, "deleted": count}


@app.post("/api/epp/movements")
async def save_epp_movement(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Movimiento EPP invalido.")
    code = normalize_text(payload.get("code"))
    key = normalize_part_key(code)
    if not key:
        raise HTTPException(status_code=400, detail="Codigo EPP requerido.")
    movement_type = normalize_text(payload.get("movement_type") or "ENTRADA")
    if movement_type not in {"ENTRADA", "SALIDA", "AJUSTE"}:
        raise HTTPException(status_code=400, detail="Tipo de movimiento invalido.")
    qty = parse_float(payload.get("quantity"), 0)
    if qty < 0:
        raise HTTPException(status_code=400, detail="La cantidad no puede ser negativa.")
    with SessionLocal() as session:
        item = session.scalar(select(EppItem).where(EppItem.code_key == key))
        if item is None:
            if movement_type != "ENTRADA":
                raise HTTPException(status_code=400, detail="Primero registra ese EPP en inventario.")
            item = EppItem(code_key=key, code=code, description=normalize_text(payload.get("description")), quantity=0, unit="PZA", updated_at=utc_now())
            session.add(item)
            session.flush()
        if movement_type == "ENTRADA":
            item.quantity = max(item.quantity + qty, 0)
        elif movement_type == "SALIDA":
            item.quantity = max(item.quantity - qty, 0)
        else:
            item.quantity = max(qty, 0)
        item.updated_at = utc_now()
        movement = EppMovement(
            item_id=item.id,
            movement_date=str(payload.get("movement_date") or utc_now().date().isoformat()),
            movement_type=movement_type,
            quantity=qty,
            balance_after=item.quantity,
            worker_name=normalize_text(payload.get("worker_name"))[:180],
            employee_id=normalize_text(payload.get("employee_id"))[:80],
            area=normalize_text(payload.get("area"))[:180],
            reference=str(payload.get("reference") or "")[:180].upper(),
            notes=str(payload.get("notes") or "").upper(),
            created_at=utc_now(),
        )
        session.add(movement)
        ensure_epp_worker_from_values(session, movement.worker_name, movement.employee_id, movement.area)
        session.commit()
        return {"ok": True, "item": epp_item_payload(item), "movement_id": movement.id}


@app.post("/api/epp/deliveries")
async def save_epp_delivery(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Entrega EPP invalida.")
    code = normalize_text(payload.get("code"))
    key = normalize_part_key(code)
    if not key:
        raise HTTPException(status_code=400, detail="Codigo EPP requerido.")
    qty = parse_float(payload.get("quantity"), 0)
    if qty <= 0:
        raise HTTPException(status_code=400, detail="La cantidad debe ser mayor a 0.")
    delivery_date = str(payload.get("delivery_date") or utc_now().date().isoformat())
    with SessionLocal() as session:
        item = session.scalar(select(EppItem).where(EppItem.code_key == key))
        if item is None:
            raise HTTPException(status_code=404, detail="EPP no encontrado.")
        life_days = max(int(parse_float(item.useful_life_days, 0)), 0)
        due_date = ""
        if life_days > 0:
            try:
                due_date = (datetime.fromisoformat(delivery_date).date() + timedelta(days=life_days)).isoformat()
            except Exception:
                due_date = ""
        item.quantity = max(item.quantity - qty, 0)
        item.updated_at = utc_now()
        delivery = EppDelivery(
            item_id=item.id,
            delivery_date=delivery_date,
            worker_name=normalize_text(payload.get("worker_name"))[:180],
            employee_id=normalize_text(payload.get("employee_id"))[:80],
            area=normalize_text(payload.get("area"))[:180],
            quantity=qty,
            useful_life_days=life_days,
            due_date=due_date,
            received_by=normalize_text(payload.get("received_by"))[:180],
            signature=normalize_text(payload.get("signature"))[:180],
            training_done=1 if payload.get("training_done") else 0,
            condition_status=normalize_text(payload.get("condition_status") or "ENTREGADO")[:80],
            notes=str(payload.get("notes") or "").upper(),
            created_at=utc_now(),
        )
        movement = EppMovement(
            item_id=item.id,
            movement_date=delivery_date,
            movement_type="SALIDA",
            quantity=qty,
            balance_after=item.quantity,
            worker_name=delivery.worker_name,
            employee_id=delivery.employee_id,
            area=delivery.area,
            reference="ENTREGA EPP",
            notes=delivery.notes,
            created_at=utc_now(),
        )
        session.add(delivery)
        session.add(movement)
        ensure_epp_worker_from_values(session, delivery.worker_name, delivery.employee_id, delivery.area)
        session.commit()
        return {"ok": True, "item": epp_item_payload(item), "delivery_id": delivery.id, "due_date": due_date}


@app.post("/api/epp/workers")
async def save_epp_worker(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Trabajador invalido.")
    with SessionLocal() as session:
        row = upsert_epp_worker(
            session,
            worker_name=str(payload.get("worker_name") or ""),
            employee_id=str(payload.get("employee_id") or ""),
            area=str(payload.get("area") or ""),
            position=str(payload.get("position") or ""),
            active=1,
        )
        session.commit()
        session.refresh(row)
        return {"ok": True, "worker": epp_worker_payload(row)}


@app.post("/api/epp/workers/delete")
async def delete_epp_worker(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    worker_id = int(parse_float(payload.get("id") if isinstance(payload, dict) else 0, 0))
    if worker_id <= 0:
        raise HTTPException(status_code=400, detail="Selecciona trabajador.")
    with SessionLocal() as session:
        row = session.get(EppWorker, worker_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Trabajador no encontrado.")
        session.delete(row)
        session.commit()
        return {"ok": True, "deleted": worker_id}


@app.get("/api/epp/items/{code}/qr.pdf")
def epp_item_qr_pdf(code: str) -> Response:
    key = normalize_part_key(code)
    with SessionLocal() as session:
        item = session.scalar(select(EppItem).where(EppItem.code_key == key))
        if item is None:
            raise HTTPException(status_code=404, detail="EPP no encontrado.")
        data = epp_qr_pdf_bytes(item)
        filename = f"QR_EPP_{item.code}.pdf".replace(" ", "_")
        return Response(
            data,
            media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="{filename}"'},
        )


@app.get("/api/epp/deliveries/{delivery_id}/pdf")
def epp_delivery_pdf(delivery_id: int) -> Response:
    with SessionLocal() as session:
        delivery = session.get(EppDelivery, delivery_id)
        if delivery is None:
            raise HTTPException(status_code=404, detail="Entrega EPP no encontrada.")
        item = session.get(EppItem, delivery.item_id)
        if item is None:
            raise HTTPException(status_code=404, detail="EPP no encontrado.")
        data = epp_delivery_pdf_bytes(delivery, item)
        filename = f"Entrega_EPP_{item.code}_{delivery.worker_name}_{delivery.delivery_date}.pdf".replace(" ", "_")
        return Response(
            data,
            media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="{filename}"'},
        )


@app.post("/api/epp/import")
async def import_epp(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Carga invalida.")
    data = str(payload.get("data") or "")
    if "," in data and data.startswith("data:"):
        data = data.split(",", 1)[1]
    try:
        raw = base64.b64decode(data)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"No se pudo leer el archivo: {exc}")
    file_name = str(payload.get("file_name") or "epp.xlsx")
    try:
        wb = load_workbook(BytesIO(raw), read_only=True, data_only=True)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Excel invalido: {exc}")
    ws = wb[wb.sheetnames[0]]
    header_row = None
    fields: dict[int, str] = {}
    for idx, values in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 25), values_only=True), 1):
        mapped = {col_idx: epp_field_for(value) for col_idx, value in enumerate(values)}
        mapped = {col_idx: field for col_idx, field in mapped.items() if field}
        if "code" in mapped.values() and "quantity" in mapped.values():
            header_row = idx
            fields = mapped
            break
    if header_row is None:
        raise HTTPException(status_code=400, detail="No encontre encabezados de Codigo y Cantidad.")
    rows: list[dict[str, Any]] = []
    skipped = 0
    for values in ws.iter_rows(min_row=header_row + 1, values_only=True):
        item: dict[str, Any] = {}
        for col_idx, field in fields.items():
            item[field] = values[col_idx] if col_idx < len(values) else None
        if not normalize_part_key(item.get("code")):
            skipped += 1
            continue
        rows.append(item)
    with SessionLocal() as session:
        if bool(payload.get("replace", True)):
            session.query(EppMovement).delete()
            session.query(EppDelivery).delete()
            session.query(EppItem).delete()
        imported = 0
        for row in rows:
            upsert_epp_item(
                session,
                code=str(row.get("code") or ""),
                description=str(row.get("description") or ""),
                category=str(row.get("category") or ""),
                size=str(row.get("size") or ""),
                unit=str(row.get("unit") or "PZA"),
                quantity=parse_float(row.get("quantity"), 0),
                min_stock=parse_float(row.get("min_stock"), 0),
                useful_life_days=int(parse_float(row.get("useful_life_days"), 0)),
                risk_area=str(row.get("risk_area") or ""),
                location=str(row.get("location") or ""),
                maintenance_notes=str(row.get("maintenance_notes") or ""),
                source_file=file_name,
            )
            imported += 1
        session.commit()
        return {"ok": True, "imported": imported, "skipped": skipped, "summary": epp_payload(session)["summary"]}


@app.get("/api/epp/export")
def export_epp(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> StreamingResponse:
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario EPP"
    ws.append(["Codigo", "Descripcion", "Categoria", "Talla", "Cantidad", "Unidad", "Minimo", "Vida util dias", "Prox. reposicion", "Trabajador prox.", "Vencidos", "Proximos", "Area riesgo", "Ubicacion", "Capacitacion", "Notas", "Estado", "QR", "Actualizado"])
    with SessionLocal() as session:
        payload = epp_payload(session)
        for row in payload.get("items") or []:
            ws.append([
                row.get("code") or "",
                row.get("description") or "",
                row.get("category") or "",
                row.get("size") or "",
                row.get("quantity") or 0,
                row.get("unit") or "PZA",
                row.get("min_stock") or 0,
                row.get("useful_life_days") or 0,
                row.get("next_due_date") or "",
                row.get("next_worker") or "",
                row.get("overdue_deliveries") or 0,
                row.get("due_soon_deliveries") or 0,
                row.get("risk_area") or "",
                row.get("location") or "",
                "SI" if row.get("training_required") else "NO",
                row.get("maintenance_notes") or "",
                row.get("status") or "",
                row.get("qr_payload") or "",
                row.get("updated_at") or "",
            ])
        ws_workers = wb.create_sheet("Trabajadores")
        ws_workers.append(["No empleado", "Trabajador", "Area", "Puesto", "Activo", "Actualizado"])
        for row in payload.get("workers") or []:
            ws_workers.append([
                row.get("employee_id") or "", row.get("worker_name") or "", row.get("area") or "",
                row.get("position") or "", "SI" if row.get("active", 1) else "NO", row.get("updated_at") or "",
            ])
        ws2 = wb.create_sheet("Entregas")
        ws2.append(["Fecha", "Codigo", "Trabajador", "No empleado", "Area", "Cantidad", "Vida util dias", "Fecha reposicion", "Recibio", "Firma", "Capacitado", "Estado", "Notas"])
        for row in payload.get("deliveries") or []:
            ws2.append([
                row.get("delivery_date") or "", row.get("item_code") or "", row.get("worker_name") or "", row.get("employee_id") or "",
                row.get("area") or "", row.get("quantity") or 0, row.get("useful_life_days") or 0, row.get("due_date") or "",
                row.get("received_by") or "", row.get("signature") or "", "SI" if row.get("training_done") else "NO",
                row.get("condition_status") or "", row.get("notes") or "",
            ])
        ws3 = wb.create_sheet("Movimientos")
        ws3.append(["Fecha", "Tipo", "Codigo", "Cantidad", "Saldo", "Trabajador", "No empleado", "Area", "Referencia", "Notas"])
        for row in payload.get("movements") or []:
            ws3.append([
                row.get("movement_date") or "", row.get("movement_type") or "", row.get("item_code") or "",
                row.get("quantity") or 0, row.get("balance_after") or 0, row.get("worker_name") or "",
                row.get("employee_id") or "", row.get("area") or "", row.get("reference") or "", row.get("notes") or "",
            ])
    for sheet in wb.worksheets:
        for column_cells in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 11), 44)
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Inventario_EPP_MGA.xlsx"'},
    )


@app.get("/api/filter-inventory")
def get_filter_inventory(response: Response, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        payload = catalog_with_inventory(session)
        movements = session.scalars(
            select(FilterInventoryMovement).order_by(FilterInventoryMovement.id.desc()).limit(80)
        ).all()
        part_by_id = {item.id: item for item in session.scalars(select(FilterInventoryItem)).all()}
        payload["movements"] = [
            {
                "id": row.id,
                "part_number": part_by_id.get(row.item_id).part_number if part_by_id.get(row.item_id) else "",
                "movement_date": row.movement_date,
                "movement_type": row.movement_type,
                "quantity": row.quantity,
                "balance_after": row.balance_after,
                "reference": row.reference,
                "equipment_code": row.equipment_code,
                "service_interval": row.service_interval,
                "notes": row.notes,
                "created_by": row.created_by,
                "created_at": row.created_at.isoformat(timespec="seconds") if row.created_at else "",
            }
            for row in movements
        ]
        return payload


@app.get("/api/filter-inventory/snapshot")
def get_filter_inventory_snapshot(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        return {
            "ok": True,
            "generated_at": utc_now().isoformat(timespec="seconds"),
            "inventory": inventory_rows_payload(session),
        }


@app.post("/api/filter-inventory/snapshot")
async def replace_filter_inventory_snapshot(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    rows = payload.get("inventory") if isinstance(payload, dict) else []
    if not isinstance(rows, list):
        raise HTTPException(status_code=400, detail="Inventario invalido.")
    with SessionLocal() as session:
        session.query(FilterInventoryMovement).delete()
        session.query(FilterInventoryItem).delete()
        imported = 0
        for row in rows:
            if not isinstance(row, dict):
                continue
            part_number = row.get("part_number")
            if not normalize_part_key(part_number):
                continue
            upsert_inventory_item(
                session,
                part_number=str(part_number),
                description=str(row.get("description") or ""),
                quantity=parse_float(row.get("quantity"), 0),
                unit=str(row.get("unit") or "PZA"),
                min_stock=parse_float(row.get("min_stock"), 0),
                location=str(row.get("location") or ""),
                source_file=str(row.get("source_file") or "desktop-sync"),
            )
            imported += 1
        session.commit()
        return {"ok": True, "imported": imported}


@app.post("/api/filter-inventory/import")
async def import_filter_inventory(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Carga invalida.")
    data = str(payload.get("data") or "")
    if "," in data and data.startswith("data:"):
        data = data.split(",", 1)[1]
    try:
        raw = base64.b64decode(data)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"No se pudo leer el archivo: {exc}")
    file_name = str(payload.get("file_name") or "inventario.xlsx")
    try:
        wb = load_workbook(BytesIO(raw), read_only=True, data_only=True)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Excel invalido: {exc}")
    ws = wb[wb.sheetnames[0]]
    header_row = None
    fields: dict[int, str] = {}
    for idx, values in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True), 1):
        mapped = {col_idx: inventory_field_for(value) for col_idx, value in enumerate(values)}
        mapped = {col_idx: field for col_idx, field in mapped.items() if field}
        if "part_number" in mapped.values() and "quantity" in mapped.values():
            header_row = idx
            fields = mapped
            break
    if header_row is None:
        raise HTTPException(status_code=400, detail="No encontre encabezados de No. parte y cantidad.")
    rows: list[dict[str, Any]] = []
    skipped = 0
    for values in ws.iter_rows(min_row=header_row + 1, values_only=True):
        item: dict[str, Any] = {}
        for col_idx, field in fields.items():
            item[field] = values[col_idx] if col_idx < len(values) else None
        if not normalize_part_key(item.get("part_number")):
            skipped += 1
            continue
        rows.append(item)
    with SessionLocal() as session:
        if bool(payload.get("replace", True)):
            session.query(FilterInventoryMovement).delete()
            session.query(FilterInventoryItem).delete()
        imported = 0
        for row in rows:
            upsert_inventory_item(
                session,
                part_number=str(row.get("part_number") or ""),
                description=str(row.get("description") or ""),
                quantity=parse_float(row.get("quantity"), 0),
                unit=str(row.get("unit") or "PZA"),
                min_stock=parse_float(row.get("min_stock"), 0),
                location=str(row.get("location") or ""),
                source_file=file_name,
            )
            imported += 1
        session.commit()
        summary = catalog_with_inventory(session)["summary"]
        return {"ok": True, "imported": imported, "skipped": skipped, "summary": summary}


@app.get("/api/filter-inventory/export")
def export_filter_inventory(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> StreamingResponse:
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario filtros"
    ws.append(["Equipo", "Tipo", "No. parte", "Descripcion", "Cantidad", "Unidad", "Minimo", "Ubicacion", "Actualizado"])
    with SessionLocal() as session:
        for item in catalog_with_inventory(session)["inventory"]:
            ws.append(
                [
                    item.get("equipment") or item.get("equipment_codes") or "",
                    item.get("item_type") or "",
                    item.get("part_number") or "",
                    item.get("description") or "",
                    item.get("quantity") or 0,
                    item.get("unit") or "PZA",
                    item.get("min_stock") or 0,
                    item.get("location") or "",
                    item.get("updated_at") or "",
                ]
            )
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Inventario_Filtros_MGA.xlsx"'},
    )


@app.post("/api/filter-inventory/movement")
async def save_filter_inventory_movement(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Movimiento invalido.")
    part_number = normalize_text(payload.get("part_number"))
    if not normalize_part_key(part_number):
        raise HTTPException(status_code=400, detail="Numero de parte requerido.")
    qty = parse_float(payload.get("quantity"), 0)
    if qty < 0:
        raise HTTPException(status_code=400, detail="La cantidad no puede ser negativa.")
    movement_type = normalize_text(payload.get("movement_type") or "ENTRADA")
    if movement_type not in {"ENTRADA", "SALIDA", "AJUSTE"}:
        raise HTTPException(status_code=400, detail="Tipo de movimiento invalido.")
    with SessionLocal() as session:
        item = session.scalar(select(FilterInventoryItem).where(FilterInventoryItem.part_key == normalize_part_key(part_number)))
        if item is None:
            item = FilterInventoryItem(part_key=normalize_part_key(part_number), part_number=part_number, quantity=0)
            session.add(item)
            session.flush()
        if payload.get("description"):
            item.description = normalize_text(payload.get("description"))
        item.unit = normalize_text(payload.get("unit") or item.unit or "PZA")
        if movement_type == "ENTRADA":
            item.quantity = max(item.quantity + qty, 0)
        elif movement_type == "SALIDA":
            item.quantity = max(item.quantity - qty, 0)
        else:
            item.quantity = max(qty, 0)
        item.updated_at = utc_now()
        movement = FilterInventoryMovement(
            item_id=item.id,
            movement_date=str(payload.get("movement_date") or utc_now().date().isoformat()),
            movement_type=movement_type,
            quantity=qty,
            balance_after=item.quantity,
            reference=str(payload.get("reference") or "")[:180],
            equipment_code=normalize_text(payload.get("equipment_code"))[:120],
            service_interval=normalize_text(payload.get("service_interval"))[:80],
            notes=str(payload.get("notes") or ""),
            created_by=str(payload.get("created_by") or "")[:160],
            created_at=utc_now(),
        )
        session.add(movement)
        session.commit()
        session.refresh(item)
        return {"ok": True, "item": inventory_item_payload(item), "movement_id": movement.id}


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


@app.post("/api/portal/snapshot")
async def publish_portal_snapshot(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Portal invalido.")
    equipment = payload.get("equipment") or []
    captures = payload.get("captures") or []
    preventives = payload.get("preventives") or []
    availability = payload.get("availability") or []
    if not isinstance(equipment, list) or not isinstance(captures, list) or not isinstance(preventives, list):
        raise HTTPException(status_code=400, detail="El portal debe contener listas validas.")
    payload["ok"] = True
    payload["source"] = "cloud-portal"
    payload["updated_at"] = utc_now().isoformat(timespec="seconds")
    with SessionLocal() as session:
        snapshot = session.scalar(select(PortalSnapshot).where(PortalSnapshot.name == "default"))
        previous_payload: dict[str, Any] = {}
        if snapshot is not None:
            previous_raw = json_loads(snapshot.payload_json)
            if isinstance(previous_raw, dict):
                previous_payload = previous_raw
        if "diesel" not in payload and isinstance(previous_payload.get("diesel"), dict):
            payload["diesel"] = previous_payload["diesel"]
            settings = payload.setdefault("settings", {})
            previous_settings = previous_payload.get("settings") if isinstance(previous_payload.get("settings"), dict) else {}
            if isinstance(settings, dict) and "meta_diesel_lh" not in settings and isinstance(previous_settings, dict):
                settings["meta_diesel_lh"] = previous_settings.get("meta_diesel_lh", 25)
        if "service_history" not in payload and isinstance(previous_payload.get("service_history"), list):
            payload["service_history"] = previous_payload["service_history"]
        if "preventive_execution" not in payload and isinstance(previous_payload.get("preventive_execution"), dict):
            payload["preventive_execution"] = previous_payload["preventive_execution"]
        if "work_orders" not in payload and isinstance(previous_payload.get("work_orders"), dict):
            payload["work_orders"] = previous_payload["work_orders"]
        if "parts_manuals" not in payload and isinstance(previous_payload.get("parts_manuals"), dict):
            payload["parts_manuals"] = previous_payload["parts_manuals"]
        if "audit_log" not in payload and isinstance(previous_payload.get("audit_log"), list):
            payload["audit_log"] = previous_payload["audit_log"]
        if "backlog" not in payload and isinstance(previous_payload.get("backlog"), dict):
            payload["backlog"] = previous_payload["backlog"]
        if "tire_tracking" not in payload and isinstance(previous_payload.get("tire_tracking"), dict):
            payload["tire_tracking"] = previous_payload["tire_tracking"]
        payload = enrich_tire_tracking(merge_work_orders_into_backlog(merge_preventive_execution_into_portal(payload)))
        if snapshot is None:
            snapshot = PortalSnapshot(name="default")
            session.add(snapshot)
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(payload)
        session.commit()
    return {
        "ok": True,
        "equipment": len(equipment),
        "captures": len(captures),
        "preventives": len(preventives),
        "service_history": len(payload.get("service_history") or []) if isinstance(payload.get("service_history"), list) else 0,
        "preventive_execution": len((payload.get("preventive_execution") or {}).get("records") or []) if isinstance(payload.get("preventive_execution"), dict) else 0,
        "work_orders": len((payload.get("work_orders") or {}).get("records") or []) if isinstance(payload.get("work_orders"), dict) else 0,
        "parts_manuals": len((payload.get("parts_manuals") or {}).get("rows") or []) if isinstance(payload.get("parts_manuals"), dict) else 0,
        "audit_log": len(payload.get("audit_log") or []) if isinstance(payload.get("audit_log"), list) else 0,
        "availability": len(availability) if isinstance(availability, list) else 0,
    }


def next_tire_event_id(events: list[dict[str, Any]]) -> int:
    current = 0
    for event in events:
        current = max(current, int(parse_float(event.get("id"), 0) or 0))
    return current + 1


def portal_snapshot_for_preventive_execution_update(session: Session) -> tuple[PortalSnapshot, dict[str, Any]]:
    snapshot = session.scalar(select(PortalSnapshot).where(PortalSnapshot.name == "default"))
    if snapshot is None:
        snapshot = PortalSnapshot(name="default")
        session.add(snapshot)
        payload = portal_fallback_payload(session)
    else:
        raw = json_loads(snapshot.payload_json)
        payload = raw if isinstance(raw, dict) else portal_fallback_payload(session)
    payload.setdefault("ok", True)
    payload.setdefault("source", "cloud-portal")
    payload.setdefault("service_history", [])
    payload.setdefault("preventive_execution", {"records": []})
    return snapshot, payload


def portal_snapshot_for_work_order_update(session: Session) -> tuple[PortalSnapshot, dict[str, Any]]:
    snapshot = session.scalar(select(PortalSnapshot).where(PortalSnapshot.name == "default"))
    if snapshot is None:
        snapshot = PortalSnapshot(name="default")
        session.add(snapshot)
        payload = portal_fallback_payload(session)
    else:
        raw = json_loads(snapshot.payload_json)
        payload = raw if isinstance(raw, dict) else portal_fallback_payload(session)
    payload.setdefault("ok", True)
    payload.setdefault("source", "cloud-portal")
    payload.setdefault("work_orders", {"records": []})
    return snapshot, payload


@app.get("/api/work-orders")
def get_work_orders() -> dict[str, Any]:
    with SessionLocal() as session:
        portal = latest_portal_payload(session)
        orders = portal.get("work_orders") if isinstance(portal.get("work_orders"), dict) else {"records": []}
        return {"ok": True, "work_orders": orders}


@app.post("/api/work-orders/records")
async def save_work_order_record(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Orden de trabajo invalida.")
    with SessionLocal() as session:
        snapshot, portal = portal_snapshot_for_work_order_update(session)
        records = [normalize_work_order_record(row) for row in work_order_records(portal) if isinstance(row, dict)]
        requested_id = str(payload.get("id") or payload.get("folio") or "").strip()
        index = next((idx for idx, row in enumerate(records) if requested_id and requested_id in {str(row.get("id") or ""), str(row.get("folio") or "")}), None)
        existing = records[index] if index is not None else None
        folio = existing.get("folio") if existing else next_work_order_folio(records)
        record = normalize_work_order_record(payload, existing=existing, folio=folio)
        if not record.get("equipment_code"):
            raise HTTPException(status_code=400, detail="Selecciona un equipo.")
        if not record.get("description"):
            raise HTTPException(status_code=400, detail="Describe el trabajo.")
        if index is None:
            records.append(record)
        else:
            records[index] = record
        portal["work_orders"] = {"records": records}
        portal["updated_at"] = utc_now().isoformat(timespec="seconds")
        portal = enrich_tire_tracking(merge_work_orders_into_backlog(merge_preventive_execution_into_portal(portal)))
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(portal)
        session.commit()
        return {"ok": True, "record": record, "portal": latest_portal_payload(session)}


@app.post("/api/work-orders/records/delete")
async def delete_work_order_record(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Solicitud invalida.")
    record_id = str(payload.get("id") or payload.get("folio") or "").strip()
    if not record_id:
        raise HTTPException(status_code=400, detail="Selecciona una OT.")
    with SessionLocal() as session:
        snapshot, portal = portal_snapshot_for_work_order_update(session)
        records = [
            row for row in work_order_records(portal)
            if isinstance(row, dict) and record_id not in {str(row.get("id") or ""), str(row.get("folio") or "")}
        ]
        portal["work_orders"] = {"records": records}
        portal["updated_at"] = utc_now().isoformat(timespec="seconds")
        portal = enrich_tire_tracking(merge_work_orders_into_backlog(merge_preventive_execution_into_portal(portal)))
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(portal)
        session.commit()
        return {"ok": True, "portal": latest_portal_payload(session)}


def apply_preventive_oil_inventory_delta(
    session: Session,
    *,
    previous: dict[str, Any] | None,
    current: dict[str, Any] | None,
    service_id: str,
) -> None:
    previous_closed = bool(previous and normalize_text(previous.get("status")) in PREVENTIVE_CLOSED_STATUSES)
    current_closed = bool(current and normalize_text(current.get("status")) in PREVENTIVE_CLOSED_STATUSES)
    service_date = str((current or previous or {}).get("close_date") or (current or previous or {}).get("service_date") or utc_now().date().isoformat())
    equipment_code = str((current or previous or {}).get("equipment_code") or "")
    service_type = str((current or previous or {}).get("service_type") or "")
    service_hours = PREVENTIVE_SERVICE_HOURS.get(normalize_text(service_type), parse_float((current or previous or {}).get("service_hours"), 0))
    service_interval = f"{service_type} / {int(service_hours)}H" if service_hours else service_type
    reference = f"SERVICIO-PREV-{service_id}"
    for key, part_number, description in OIL_STOCK_FIELDS:
        previous_qty = max(parse_float(previous.get(key), 0), 0) if previous_closed and previous else 0
        current_qty = max(parse_float(current.get(key), 0), 0) if current_closed and current else 0
        delta = round(current_qty - previous_qty, 4)
        if abs(delta) < 0.0001:
            continue
        item = session.scalar(select(FilterInventoryItem).where(FilterInventoryItem.part_key == normalize_part_key(part_number)))
        if item is None:
            item = FilterInventoryItem(part_key=normalize_part_key(part_number), part_number=part_number, description=description, quantity=0, unit="L")
            session.add(item)
            session.flush()
        if not item.description:
            item.description = description
        item.unit = "L"
        movement_type = "SALIDA" if delta > 0 else "ENTRADA"
        qty = abs(delta)
        if movement_type == "SALIDA":
            item.quantity = max(parse_float(item.quantity, 0) - qty, 0)
        else:
            item.quantity = max(parse_float(item.quantity, 0) + qty, 0)
        item.updated_at = utc_now()
        movement = FilterInventoryMovement(
            item_id=item.id,
            movement_date=service_date[:10],
            movement_type=movement_type,
            quantity=qty,
            balance_after=item.quantity,
            reference=reference,
            equipment_code=equipment_code[:120],
            service_interval=service_interval[:80],
            notes="Movimiento automatico por servicio preventivo web",
            created_by="Portal web",
            created_at=utc_now(),
        )
        session.add(movement)


@app.get("/api/preventive-execution")
def get_preventive_execution() -> dict[str, Any]:
    with SessionLocal() as session:
        portal = latest_portal_payload(session)
        execution = portal.get("preventive_execution") if isinstance(portal.get("preventive_execution"), dict) else {"records": []}
        return {"ok": True, "preventive_execution": execution, "service_history": portal.get("service_history") or []}


@app.post("/api/preventive-execution/records")
async def save_preventive_execution_record(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Solicitud invalida.")
    record = normalize_preventive_execution_record(payload)
    if not record.get("equipment_code"):
        raise HTTPException(status_code=400, detail="Selecciona un equipo.")
    if record.get("service_type") not in PREVENTIVE_SERVICE_HOURS:
        raise HTTPException(status_code=400, detail="Selecciona PM1, PM2, PM3 o PM4.")
    with SessionLocal() as session:
        snapshot, portal = portal_snapshot_for_preventive_execution_update(session)
        records = [normalize_preventive_execution_record(row) for row in preventive_execution_records(portal) if isinstance(row, dict)]
        index = next((idx for idx, row in enumerate(records) if str(row.get("id") or "") == str(record.get("id") or "")), None)
        previous_record = records[index] if index is not None else None
        if index is None:
            records.append(record)
        else:
            record["created_at"] = records[index].get("created_at") or record["created_at"]
            records[index] = record
        apply_preventive_oil_inventory_delta(session, previous=previous_record, current=record, service_id=str(record.get("id") or ""))
        portal["preventive_execution"] = {"records": records}
        portal["updated_at"] = utc_now().isoformat(timespec="seconds")
        portal = enrich_tire_tracking(merge_preventive_execution_into_portal(portal))
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(portal)
        session.commit()
        return {"ok": True, "record": record, "portal": latest_portal_payload(session)}


@app.post("/api/preventive-execution/records/delete")
async def delete_preventive_execution_record(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Solicitud invalida.")
    record_id = str(payload.get("id") or "").strip()
    if not record_id:
        raise HTTPException(status_code=400, detail="Selecciona un registro.")
    with SessionLocal() as session:
        snapshot, portal = portal_snapshot_for_preventive_execution_update(session)
        previous_record = next(
            (normalize_preventive_execution_record(row) for row in preventive_execution_records(portal) if isinstance(row, dict) and str(row.get("id") or "") == record_id),
            None,
        )
        if previous_record:
            apply_preventive_oil_inventory_delta(session, previous=previous_record, current=None, service_id=record_id)
        records = [row for row in preventive_execution_records(portal) if isinstance(row, dict) and str(row.get("id") or "") != record_id]
        history = [
            row for row in (portal.get("service_history") or [])
            if not (isinstance(row, dict) and str(row.get("web_id") or "") == record_id)
        ]
        portal["preventive_execution"] = {"records": records}
        portal["service_history"] = history
        portal["updated_at"] = utc_now().isoformat(timespec="seconds")
        portal = enrich_tire_tracking(merge_preventive_execution_into_portal(portal))
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(portal)
        session.commit()
        return {"ok": True, "portal": latest_portal_payload(session)}


def portal_snapshot_for_tire_update(session: Session) -> tuple[PortalSnapshot, dict[str, Any]]:
    snapshot = session.scalar(select(PortalSnapshot).where(PortalSnapshot.name == "default"))
    if snapshot is None:
        snapshot = PortalSnapshot(name="default")
        session.add(snapshot)
        payload = portal_fallback_payload(session)
    else:
        raw = json_loads(snapshot.payload_json)
        payload = raw if isinstance(raw, dict) else portal_fallback_payload(session)
    payload.setdefault("ok", True)
    payload.setdefault("source", "cloud-portal")
    payload.setdefault("tire_kpi", {"rows": [], "summary": {}})
    payload.setdefault("tire_tracking", {"events": []})
    return snapshot, payload


@app.get("/api/tire-tracking")
def get_tire_tracking(response: Response) -> dict[str, Any]:
    response.headers["Cache-Control"] = "no-store, max-age=0"
    with SessionLocal() as session:
        portal = latest_portal_payload(session)
        return {"ok": True, "tire_kpi": portal.get("tire_kpi", {}), "tire_tracking": portal.get("tire_tracking", {})}


@app.post("/api/tire-tracking/events")
async def save_tire_tracking_event(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Evento de llanta invalido.")
    tire_code = normalize_text(payload.get("tire_code"))
    if not tire_code:
        raise HTTPException(status_code=400, detail="La serie de llanta es obligatoria.")
    event_type = normalize_text(payload.get("event_type") or "INSPECCION")
    if event_type not in {"MODIFICACION", "INSPECCION", "MOVIMIENTO", "MONTAJE", "ROTACION", "REPARACION", "DESMONTAJE", "BAJA"}:
        raise HTTPException(status_code=400, detail="Tipo de evento invalido.")
    event_date = iso_date(payload.get("event_date"))
    with SessionLocal() as session:
        snapshot, portal = portal_snapshot_for_tire_update(session)
        rows = tire_rows_list(portal)
        index = next((idx for idx, row in enumerate(rows) if normalize_text(row.get("tire_code")) == tire_code), None)
        row = dict(rows[index]) if index is not None else {"tire_code": tire_code}
        for key in (
            "equipment_code",
            "position",
            "brand",
            "model",
            "size",
            "install_date",
            "install_meter",
            "current_meter",
            "target_life_hours",
            "tread_initial",
            "tread_current",
            "pressure_current",
            "status",
            "notes",
        ):
            if key in payload:
                row[key] = payload.get(key)
        if event_type == "BAJA":
            row["status"] = "BAJA"
        if event_type == "REPARACION":
            row["status"] = "REPARACION"
        row = normalize_tire_row(row)
        if index is None:
            rows.append(row)
        else:
            rows[index] = row
        events = tire_event_list(portal)
        event = {
            "id": next_tire_event_id(events),
            "event_date": event_date,
            "event_type": event_type,
            "tire_code": tire_code,
            "equipment_code": row.get("equipment_code", ""),
            "position": row.get("position", ""),
            "meter": row.get("current_meter", 0),
            "tread_mm": row.get("tread_current", 0),
            "pressure_psi": row.get("pressure_current", 0),
            "status": row.get("status", ""),
            "control_status": row.get("control_status", ""),
            "technician": normalize_text(payload.get("technician"))[:180],
            "notes": str(payload.get("event_notes") or payload.get("notes") or "").strip(),
            "created_at": utc_now().isoformat(timespec="seconds"),
            "source": "web",
        }
        events.append(event)
        portal["tire_kpi"] = {"rows": rows, "summary": {}}
        portal["tire_tracking"] = {"events": events}
        portal["updated_at"] = utc_now().isoformat(timespec="seconds")
        portal = enrich_tire_tracking(portal)
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(portal)
        session.commit()
        return {"ok": True, "row": row, "event": event, "portal": latest_portal_payload(session)}


@app.post("/api/tire-tracking/events/delete")
async def delete_tire_tracking_event(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Solicitud invalida.")
    event_id = int(parse_float(payload.get("id"), 0) or 0)
    if not event_id:
        raise HTTPException(status_code=400, detail="Selecciona un evento.")
    with SessionLocal() as session:
        snapshot, portal = portal_snapshot_for_tire_update(session)
        events = [event for event in tire_event_list(portal) if int(parse_float(event.get("id"), 0) or 0) != event_id]
        portal["tire_tracking"] = {"events": events}
        portal["updated_at"] = utc_now().isoformat(timespec="seconds")
        portal = enrich_tire_tracking(portal)
        snapshot.updated_at = utc_now()
        snapshot.payload_json = json_dumps(portal)
        session.commit()
        return {"ok": True, "portal": latest_portal_payload(session)}


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
    updated = 0
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
                photos = record.get("photos") or []
                stored_record = dict(record)
                stored_record["photos"] = []
                stored_record["shift"] = normalize_capture_shift_py(stored_record.get("shift") or stored_record.get("turno"))
                record["shift"] = stored_record["shift"]
                period_error = closed_capture_period_error(stored_record)
                if period_error:
                    raise ValueError(period_error)
                existing = session.scalar(select(MobileCapture).where(MobileCapture.mobile_id == mobile_id))
                if existing is not None:
                    existing_payload = json_loads(existing.payload_json)
                    changed = json.dumps(existing_payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")) != json.dumps(
                        stored_record,
                        ensure_ascii=False,
                        sort_keys=True,
                        separators=(",", ":"),
                    )
                    if isinstance(existing_payload, dict):
                        for hose_payload in mobile_hose_change_rows(existing_payload, existing.source_device, existing.user_name):
                            upsert_hose_change(session, hose_payload)
                    if changed:
                        existing.source_device = source_device
                        existing.user_name = str(record.get("user_name") or user_name)
                        existing.equipment_code = str(record.get("equipment_code") or "")
                        existing.component_name = str(record.get("component_name") or "")
                        existing.work_date = str(record.get("work_date") or "")
                        existing.payload_json = json_dumps(stored_record)
                        existing.received_at = utc_now()
                        existing.desktop_imported_at = None
                        updated += 1
                    else:
                        skipped += 1
                    evidence_count = replace_mobile_photos(session, existing.id, mobile_id, photos)
                    results.append(
                        {
                            "mobile_id": mobile_id,
                            "capture_id": existing.id,
                            "created": False,
                            "updated": changed,
                            "stored": True,
                            "desktop_imported": existing.desktop_imported_at is not None,
                            "evidence": evidence_count,
                        }
                    )
                    continue

                duplicate = mobile_capture_by_merge_key(session, stored_record)
                if duplicate is not None:
                    duplicate_payload = json_loads(duplicate.payload_json)
                    changed = json.dumps(duplicate_payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")) != json.dumps(
                        stored_record,
                        ensure_ascii=False,
                        sort_keys=True,
                        separators=(",", ":"),
                    )
                    if changed:
                        duplicate.mobile_id = mobile_id
                        duplicate.source_device = source_device
                        duplicate.user_name = str(record.get("user_name") or user_name)
                        duplicate.equipment_code = str(record.get("equipment_code") or "")
                        duplicate.component_name = str(record.get("component_name") or "")
                        duplicate.work_date = str(record.get("work_date") or "")
                        duplicate.payload_json = json_dumps(stored_record)
                        duplicate.received_at = utc_now()
                        duplicate.desktop_imported_at = None
                        updated += 1
                    else:
                        skipped += 1
                    evidence_count = replace_mobile_photos(session, duplicate.id, mobile_id, photos)
                    results.append(
                        {
                            "mobile_id": mobile_id,
                            "capture_id": duplicate.id,
                            "created": False,
                            "updated": changed,
                            "stored": True,
                            "duplicate_key": True,
                            "desktop_imported": duplicate.desktop_imported_at is not None,
                            "evidence": evidence_count,
                        }
                    )
                    continue

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
                hose_changes = 0
                for hose_payload in mobile_hose_change_rows(stored_record, source_device, capture.user_name):
                    upsert_hose_change(session, hose_payload)
                    hose_changes += 1
                evidence_count = replace_mobile_photos(session, capture.id, mobile_id, photos)
                created += 1
                results.append(
                    {
                        "mobile_id": mobile_id,
                        "capture_id": capture.id,
                        "created": True,
                        "stored": True,
                        "desktop_imported": False,
                        "evidence": evidence_count,
                        "hose_changes": hose_changes,
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
        "updated": updated,
        "skipped": skipped,
        "errors": errors,
        "stored": created + updated + skipped,
        "captures": counts,
        "results": results,
    }


@app.post("/api/portal/captures/delete")
async def delete_portal_capture(request: Request, _auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
    require_api_key(_auth)
    payload = await request.json()
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="Solicitud invalida.")
    mobile_id = str(payload.get("mobile_id") or "").strip()
    capture_id = int(parse_float(payload.get("id"), 0) or 0)
    delete_key = capture_delete_key(payload)
    if not mobile_id and not capture_id and not any(delete_key):
        raise HTTPException(status_code=400, detail="No se recibio identificador de captura.")
    period_error = closed_capture_period_error({"work_date": payload.get("work_date") or delete_key[0]})
    if period_error:
        raise HTTPException(status_code=409, detail=period_error)

    with SessionLocal() as session:
        mobile_rows: list[MobileCapture] = []
        if mobile_id:
            row = session.scalar(select(MobileCapture).where(MobileCapture.mobile_id == mobile_id))
            if row is not None:
                mobile_rows.append(row)
        if any(delete_key):
            candidates = session.scalars(
                select(MobileCapture)
                .where(MobileCapture.work_date == delete_key[0])
                .where(MobileCapture.equipment_code == delete_key[2])
            ).all()
            existing_ids = {row.id for row in mobile_rows}
            for row in candidates:
                if row.id in existing_ids:
                    continue
                if capture_merge_key(mobile_capture_portal_row(row)) == delete_key:
                    mobile_rows.append(row)
                    existing_ids.add(row.id)
        if capture_id and not mobile_rows:
            row = session.scalar(select(MobileCapture).where(MobileCapture.id == capture_id))
            if row is not None:
                mobile_rows.append(row)
        for row in mobile_rows:
            period_error = closed_capture_period_error(mobile_capture_portal_row(row))
            if period_error:
                raise HTTPException(status_code=409, detail=period_error)

        deleted_mobile = 0
        for row in mobile_rows:
            session.delete(row)
            deleted_mobile += 1

        removed_snapshot = 0
        snapshot = session.scalar(select(PortalSnapshot).where(PortalSnapshot.name == "default"))
        if snapshot is not None:
            snapshot_payload = json_loads(snapshot.payload_json)
            if isinstance(snapshot_payload, dict):
                removed_snapshot = remove_capture_rows_from_portal_payload(snapshot_payload, delete_key, mobile_id, capture_id)
                if removed_snapshot:
                    snapshot.updated_at = utc_now()
                    snapshot.payload_json = json_dumps(snapshot_payload)
        deletion_payload = {
            "id": capture_id,
            "mobile_id": mobile_id,
            "work_date": payload.get("work_date") or delete_key[0],
            "shift": payload.get("shift") or delete_key[1],
            "equipment_code": payload.get("equipment_code") or payload.get("equipment") or delete_key[2],
            "component_name": payload.get("component_name") or payload.get("component") or delete_key[3],
            "component": payload.get("component") or payload.get("component_name") or delete_key[3],
            "source": "portal-web",
            "deleted_at": utc_now().isoformat(timespec="seconds"),
        }
        deletion = CaptureDeletion(
            mobile_id=mobile_id,
            source_device="portal-web",
            user_name=str(payload.get("user_name") or "Portal web"),
            equipment_code=str(deletion_payload["equipment_code"] or ""),
            component_name=str(deletion_payload["component_name"] or ""),
            work_date=str(deletion_payload["work_date"] or ""),
            shift=str(deletion_payload["shift"] or ""),
            payload_json=json_dumps(deletion_payload),
        )
        session.add(deletion)
        session.commit()
        deletion_id = deletion.id
        counts = capture_counts(session)

    return {"ok": True, "deleted_mobile": deleted_mobile, "deleted_portal": removed_snapshot, "deletion_id": deletion_id, "captures": counts}


@app.get("/api/desktop/deletions")
def desktop_deletions(
    limit: int = Query(default=1000, ge=1, le=5000),
    order: str = Query(default="asc"),
    _auth: str | None = Header(default=None, alias="X-MGA-API-Key"),
) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        sort_order = CaptureDeletion.id.desc() if str(order or "").lower().startswith("desc") else CaptureDeletion.id.asc()
        rows = session.scalars(select(CaptureDeletion).order_by(sort_order).limit(limit)).all()
        deletions = [
            {
                "id": row.id,
                "mobile_id": row.mobile_id,
                "source_device": row.source_device,
                "user_name": row.user_name,
                "deleted_at": row.deleted_at.isoformat(timespec="seconds") if row.deleted_at else "",
                "payload": capture_deletion_payload(row),
            }
            for row in rows
        ]
    return {"ok": True, "deletions": deletions, "count": len(deletions)}


@app.post("/api/desktop-sync")
async def desktop_capture_sync(
    request: Request,
    _auth: str | None = Header(default=None, alias="X-MGA-API-Key"),
) -> dict[str, Any]:
    require_api_key(_auth)
    package = await request.json()
    if not isinstance(package, dict):
        raise HTTPException(status_code=400, detail="Paquete de sincronizacion invalido.")
    device_id = str(package.get("device_id") or "").strip()[:180]
    cursor = max(0, int(package.get("cursor") or 0))
    incoming = package.get("records") or []
    if not device_id:
        raise HTTPException(status_code=400, detail="Falta device_id.")
    if not isinstance(incoming, list):
        raise HTTPException(status_code=400, detail="records debe ser una lista.")

    mappings: list[dict[str, str]] = []
    created = 0
    with SessionLocal() as session:
        for raw in incoming[:500]:
            if not isinstance(raw, dict):
                continue
            sent_id = str(raw.get("sync_id") or "").strip()[:80]
            logical_key = str(raw.get("logical_key") or "").strip()[:500]
            if not sent_id or not logical_key:
                continue
            latest = session.scalar(
                select(DesktopCaptureChange)
                .where(
                    (DesktopCaptureChange.sync_id == sent_id)
                    | (DesktopCaptureChange.logical_key == logical_key)
                )
                .order_by(DesktopCaptureChange.id.desc())
                .limit(1)
            )
            canonical_id = latest.sync_id if latest is not None else sent_id
            deleted = 1 if raw.get("deleted") else 0
            payload = raw.get("payload") if isinstance(raw.get("payload"), dict) else {}
            payload_json = json_dumps(payload)
            unchanged = latest is not None and latest.deleted == deleted and latest.payload_json == payload_json
            if not unchanged:
                session.add(
                    DesktopCaptureChange(
                        sync_id=canonical_id,
                        logical_key=logical_key,
                        source_device=device_id,
                        source_updated_at=str(raw.get("updated_at") or "")[:40],
                        deleted=deleted,
                        payload_json=payload_json,
                    )
                )
                session.flush()
                created += 1
            mappings.append({"sent_id": sent_id, "sync_id": canonical_id})
        session.commit()

        changes = session.scalars(
            select(DesktopCaptureChange)
            .where(DesktopCaptureChange.id > cursor)
            .order_by(DesktopCaptureChange.id.asc())
            .limit(1000)
        ).all()
        records = [
            {
                "revision": row.id,
                "sync_id": row.sync_id,
                "logical_key": row.logical_key,
                "source_device": row.source_device,
                "updated_at": row.received_at.isoformat(timespec="seconds"),
                "deleted": bool(row.deleted),
                "payload": json_loads(row.payload_json),
            }
            for row in changes
        ]
        next_cursor = records[-1]["revision"] if records else cursor
    return {
        "ok": True,
        "accepted": len(mappings),
        "created": created,
        "mappings": mappings,
        "records": records,
        "cursor": next_cursor,
        "has_more": len(records) >= 1000,
    }


@app.get("/api/desktop/pending")
def desktop_pending(
    limit: int = Query(default=200, ge=1, le=1000),
    include_imported: bool = Query(default=False),
    order: str = Query(default="asc"),
    _auth: str | None = Header(default=None, alias="X-MGA-API-Key"),
) -> dict[str, Any]:
    require_api_key(_auth)
    with SessionLocal() as session:
        sort_order = MobileCapture.id.desc() if str(order or "").lower().startswith("desc") else MobileCapture.id.asc()
        query = select(MobileCapture).order_by(sort_order).limit(limit)
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
