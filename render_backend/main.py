from __future__ import annotations

import base64
import json
import math
import os
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas as pdf_canvas
from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, create_engine, func, select
from sqlalchemy import Float
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


engine = create_engine(database_url(), pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
Base.metadata.create_all(engine)

app = FastAPI(title="MGA Cloud Sync", version="1.3.1")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)
STATIC_DIR = Path(__file__).resolve().parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
PRODUCT_CATALOG_PATH = STATIC_DIR / "productos_catalog.json"
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
    return json.loads(value)


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


def normalize_part_key(value: Any) -> str:
    text = normalize_text(value)
    return "".join(ch for ch in text if ch.isalnum())


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


def latest_catalog_payload(session: Session) -> dict[str, Any]:
    snapshot = session.scalar(select(CatalogSnapshot).where(CatalogSnapshot.name == "default"))
    if snapshot is None:
        return {"ok": True, "source": "cloud-empty", "equipment": []}
    payload = json_loads(snapshot.payload_json)
    return payload if isinstance(payload, dict) else {"ok": True, "source": "cloud", "equipment": []}


def portal_fallback_payload(session: Session) -> dict[str, Any]:
    catalog = latest_catalog_payload(session)
    equipment = catalog.get("equipment") if isinstance(catalog, dict) else []
    if not isinstance(equipment, list):
        equipment = []
    captures: list[dict[str, Any]] = []
    for row in session.scalars(select(MobileCapture).order_by(MobileCapture.work_date.desc(), MobileCapture.id.desc()).limit(1000)).all():
        payload = json_loads(row.payload_json)
        if not isinstance(payload, dict):
            payload = {}
        captures.append(
            {
                "id": row.id,
                "work_date": row.work_date or str(payload.get("work_date") or ""),
                "shift": str(payload.get("shift") or payload.get("turno") or "General"),
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
                "fault": str(payload.get("fault") or ""),
                "wear": str(payload.get("wear") or ""),
                "status": str(payload.get("status") or "Disponible"),
                "observations": str(payload.get("observations") or payload.get("details") or ""),
                "evidence_count": len(row.photos or []),
            }
        )
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
            "meta_tmef": 8,
            "meta_tmpr": 4,
        },
        "equipment": equipment,
        "preventives": [],
        "captures": captures,
        "availability": [],
        "kpi_groups": ["Todos los equipos", "Equipos de Barrenacion", "Equipos de Rezagado", "KPI Aceites", "KPI Llantas"],
        "kpi_reports": {},
        "oil_kpi": {"rows": [], "totals": {}, "columns": []},
        "tire_kpi": {"rows": [], "summary": {}},
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
    payload["updated_at"] = snapshot.updated_at.isoformat(timespec="seconds") if snapshot.updated_at else ""
    return payload


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


def requisition_pdf_bytes(row: CloudRequisition, items: list[CloudRequisitionItem]) -> bytes:
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
def get_portal() -> dict[str, Any]:
    with SessionLocal() as session:
        return latest_portal_payload(session)


@app.get("/api/products")
def get_products(q: str = Query(default=""), limit: int = Query(default=120, ge=1, le=25000)) -> dict[str, Any]:
    with SessionLocal() as session:
        rows = product_rows(session, q, limit)
        return {"ok": True, "products": rows, "count": len(rows)}


@app.get("/api/requisitions")
def get_requisitions() -> dict[str, Any]:
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


WAREHOUSE_HTML = r"""<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MGA Almacen de filtros</title>
  <style>
    :root { --blue:#071f49; --blue2:#0d3272; --teal:#009c9a; --red:#c81e1e; --muted:#667085; --line:#d8dee8; --bg:#f4f6f9; --panel:#ffffff; --shadow:0 16px 40px rgba(7,31,73,.12); }
    * { box-sizing:border-box; }
    body { margin:0; font-family:Segoe UI, Arial, sans-serif; color:#1f2937; background:linear-gradient(180deg,#eef3fa 0%,#f7f9fc 42%,#eef3fa 100%); }
    .hero { position:relative; overflow:hidden; color:white; padding:18px 28px 24px; display:flex; justify-content:space-between; gap:22px; align-items:center; background:radial-gradient(circle at 15% 0%, rgba(255,255,255,.18), transparent 28%), linear-gradient(135deg,#050b18 0%, var(--blue) 56%, #0b1735 100%); box-shadow:0 18px 44px rgba(7,31,73,.24); }
    .hero::after { content:""; position:absolute; right:-60px; bottom:-90px; width:360px; height:220px; border-radius:999px; background:radial-gradient(circle, rgba(0,156,154,.26), transparent 68%); pointer-events:none; }
    .brand { position:relative; z-index:1; display:flex; align-items:center; gap:18px; min-width:0; }
    .corner-logo { width:118px; height:78px; object-fit:contain; flex:0 0 auto; padding:8px 10px; border-radius:8px; background:white; border:1px solid rgba(255,255,255,.55); box-shadow:0 16px 34px rgba(0,0,0,.24); }
    header h1 { margin:0; font-size:28px; line-height:1.05; letter-spacing:0; }
    header p { margin:7px 0 0; color:#dbeafe; font-size:14px; }
    .key-card { position:relative; z-index:1; min-width:280px; padding:12px; border:1px solid rgba(255,255,255,.16); border-radius:8px; background:rgba(255,255,255,.08); backdrop-filter:blur(10px); }
    .key-card label { color:#dbeafe; }
    header input { min-width:260px; padding:10px 11px; border:1px solid rgba(255,255,255,.28); border-radius:6px; color:white; background:rgba(255,255,255,.1); outline:none; }
    header input::placeholder { color:#cbd5e1; }
    main { width:min(1480px, 100%); margin:0 auto; padding:18px; display:grid; gap:14px; }
    .tabs { display:flex; gap:8px; flex-wrap:wrap; padding:6px; border:1px solid var(--line); border-radius:8px; background:rgba(255,255,255,.78); box-shadow:0 8px 28px rgba(7,31,73,.08); }
    .tabs button, .btn { border:0; background:var(--blue); color:white; padding:10px 14px; border-radius:6px; font-weight:700; cursor:pointer; transition:transform .15s ease, box-shadow .15s ease, background .15s ease; }
    .tabs button:hover, .btn:hover { transform:translateY(-1px); box-shadow:0 10px 20px rgba(7,31,73,.16); }
    .tabs button.active { background:linear-gradient(135deg,var(--teal),#0b7877); }
    .btn.secondary { background:white; color:var(--blue); border:1px solid var(--line); }
    .btn.danger { background:linear-gradient(135deg,#b31212,var(--red)); }
    .panel { position:relative; overflow:hidden; background:rgba(255,255,255,.92); border:1px solid rgba(216,222,232,.9); border-radius:8px; padding:16px; box-shadow:var(--shadow); }
    .toolbar { display:grid; grid-template-columns:repeat(5, minmax(140px, 1fr)); gap:10px; align-items:end; }
    label { display:grid; gap:4px; color:#344054; font-size:12px; font-weight:700; }
    input, select, textarea { width:100%; padding:9px 10px; border:1px solid #cbd5e1; border-radius:6px; font:inherit; background:white; outline:none; transition:border .15s ease, box-shadow .15s ease; }
    input:focus, select:focus, textarea:focus { border-color:var(--teal); box-shadow:0 0 0 3px rgba(0,156,154,.14); }
    .stats { display:grid; grid-template-columns:repeat(5, 1fr); gap:10px; }
    .stat { position:relative; overflow:hidden; background:linear-gradient(180deg,#fff,#f8fbff); border:1px solid var(--line); padding:14px 15px; border-radius:8px; box-shadow:0 10px 26px rgba(7,31,73,.08); }
    .stat::after { content:""; position:absolute; right:-18px; top:-18px; width:74px; height:74px; border-radius:999px; background:rgba(0,156,154,.08); }
    .stat strong { display:block; color:var(--blue); font-size:30px; line-height:1; margin-bottom:5px; }
    .view { display:none; }
    .view.active { display:grid; gap:14px; }
    table { width:100%; border-collapse:separate; border-spacing:0; background:white; }
    th, td { border-bottom:1px solid var(--line); padding:9px 10px; font-size:13px; vertical-align:top; }
    th { background:linear-gradient(180deg,#eef3fb,#e4ebf6); color:#243042; position:sticky; top:0; z-index:1; text-transform:uppercase; font-size:12px; }
    tbody tr:hover { background:#f8fbff; }
    .table-wrap { max-height:620px; overflow:auto; border:1px solid var(--line); border-radius:8px; background:white; }
    .pill { display:inline-block; padding:2px 7px; border-radius:999px; font-weight:700; font-size:12px; }
    .ok { color:#047857; background:#d1fae5; }
    .bad { color:#b91c1c; background:#fee2e2; }
    .warn { color:#92400e; background:#fef3c7; }
    .muted { color:var(--muted); }
    .grid2 { display:grid; grid-template-columns:1.1fr .9fr; gap:14px; align-items:start; }
    .movement-grid { display:grid; grid-template-columns:repeat(4, 1fr); gap:10px; }
    .req-header-grid { display:grid; grid-template-columns:repeat(3, 1fr); gap:10px; }
    .req-item-grid { display:grid; grid-template-columns:110px 150px 1fr 1.6fr; gap:10px; align-items:end; }
    .req-actions { display:flex; gap:8px; flex-wrap:wrap; margin-top:10px; }
    .wide { grid-column:1 / -1; }
    .dashboard-grid { display:grid; grid-template-columns:repeat(4, 1fr); gap:10px; }
    .metric-card { border:1px solid var(--line); border-radius:8px; padding:13px; background:linear-gradient(180deg,#fff,#f8fbff); }
    .metric-card span { display:block; color:var(--muted); font-size:12px; font-weight:800; text-transform:uppercase; }
    .metric-card strong { display:block; color:var(--blue); font-size:30px; margin-top:5px; }
    .metric-card .bar-track { height:8px; border-radius:999px; background:#e5e7eb; margin-top:10px; overflow:hidden; }
    .metric-card .bar-fill { display:block; height:100%; background:var(--teal); }
    .metric-card.bad .bar-fill { background:var(--red); }
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
    .kpi-report-table { margin-top:14px; max-height:420px; }
    .chart { display:flex; align-items:end; gap:12px; min-height:270px; padding:20px 16px 28px; border:1px solid var(--line); border-radius:8px; background:linear-gradient(180deg,#fff,#f8fbff); overflow:auto; }
    .kpi-format-mode .chart { display:block; min-height:330px; padding:12px 14px 18px; }
    .kpi-chart-head { display:flex; align-items:center; gap:10px; margin-bottom:12px; color:#111827; font-size:11px; }
    .kpi-mini-tabs { display:grid; grid-template-columns:repeat(4, minmax(96px, 1fr)); gap:4px; flex:1; }
    .kpi-mini-tabs span, .kpi-mini-tabs button { border:1px solid #111; padding:7px 9px; background:white; color:#111; font:inherit; font-size:12px; text-align:left; cursor:pointer; }
    .kpi-mini-tabs span.active, .kpi-mini-tabs button.active { background:var(--teal); color:#031b1b; }
    .chart-plot { min-height:258px; display:flex; align-items:end; gap:12px; overflow:auto; padding:18px 6px 8px; background:repeating-linear-gradient(to top, transparent 0, transparent 51px, rgba(100,116,139,.25) 52px); }
    .chart-bar { min-width:54px; display:grid; align-content:end; gap:6px; text-align:center; color:#344054; font-size:11px; }
    .chart-bar i { display:block; height:var(--h); min-height:4px; border-radius:6px 6px 0 0; background:linear-gradient(180deg,#12b7b6,#078080); box-shadow:0 9px 18px rgba(0,156,154,.18); }
    .chart-bar.out i { background:linear-gradient(180deg,#e11d48,#b31212); }
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
    .subtle-title { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:10px; }
    .subtle-title h3 { margin:0; color:var(--blue); }
    .print-only { display:none; }
    @media print {
      header, .tabs, #stats, .dashboard-controls, .no-print { display:none !important; }
      main { width:100%; padding:0; }
      body { background:white; }
      .view { display:none !important; }
      #dashboard { display:block !important; }
      .panel { box-shadow:none; border:0; padding:0; }
      .table-wrap { max-height:none; overflow:visible; border:0; }
      th { position:static; }
      .print-only { display:block; }
    }
    @media (max-width: 900px) { .hero, .grid2 { display:block; } .brand { align-items:flex-start; } .corner-logo { width:96px; height:66px; margin-bottom:10px; } .toolbar, .movement-grid, .req-header-grid, .req-item-grid, .stats { grid-template-columns:1fr; } header input { min-width:0; margin-top:10px; } .key-card { margin-top:14px; min-width:0; } }
    @media (max-width: 1050px) { .dashboard-grid, .kpi-format-board, .kpi-special-mode #kpiCards { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <header class="hero">
    <div class="brand">
      <img class="corner-logo" src="/static/mga-corner-logo.jfif" alt="MGA">
      <div><h1>Portal MGA mantenimiento</h1><p>KPI, preventivos, bitacora, disponibilidad e inventario de filtros</p></div>
    </div>
    <div class="key-card"><label>Clave para editar<input id="apiKey" type="password" placeholder="Pegar clave aqui"></label></div>
  </header>
  <main>
    <nav class="tabs">
      <button class="active" data-tab="dashboard">Dashboard KPI</button>
      <button data-tab="preventivos">PR Preventivos</button>
      <button data-tab="bitacora">Bitacora</button>
      <button data-tab="disponibilidad">Disponibilidad</button>
      <button data-tab="requisiciones">Requisiciones</button>
      <button data-tab="equipos">Filtros por equipo</button>
      <button data-tab="inventario">Concentrado / movimientos</button>
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
            <label>Equipo<input id="reqEquipment" value="PARA STOCK"></label>
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
    let portal = { equipment: [], preventives: [], captures: [], availability: [], settings: {}, period: {}, products: [] };
    let products = [];
    let requisitions = [];
    let currentReqId = null;
    let currentReqItemIndex = null;
    let currentReqItems = [];
    let selectedKpiMetric = "availability";
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
    function setDashboardMode(mode){
      const area = $("kpiPrintArea");
      area.classList.toggle("kpi-format-mode", mode === "format");
      area.classList.toggle("kpi-special-mode", mode !== "format");
      $("kpiSideCards").innerHTML = "";
    }
    function metricCardHtml(label, value, note, width, bad=false){
      return `<div class="metric-card ${bad ? "bad" : ""}"><span>${esc(label)}</span><strong>${esc(value)}</strong><small class="muted">${esc(note)}</small><div class="bar-track"><i class="bar-fill" style="width:${Math.max(Math.min(Number(width || 0),100),0)}%"></i></div></div>`;
    }
    async function load(){
      const [r, p, prod, req] = await Promise.all([
        fetch("/api/filter-inventory", {headers: headers()}),
        fetch("/api/portal", {headers: headers()}),
        fetch("/api/products?limit=25000", {headers: headers()}),
        fetch("/api/requisitions", {headers: headers()})
      ]);
      if(!r.ok) throw new Error(await apiError(r));
      if(!p.ok) throw new Error(await apiError(p));
      if(!prod.ok) throw new Error(await apiError(prod));
      if(!req.ok) throw new Error(await apiError(req));
      data = await r.json();
      portal = await p.json();
      products = (await prod.json()).products || [];
      const reqPayload = await req.json();
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
      ].map(([k,v]) => `<div class="stat"><strong>${v}</strong>${k}</div>`).join("");
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
      return text.includes("NO DISPONIBLE") || text.includes("FUERA") || text.includes("NO DISP");
    }
    function groupMatches(eq, group){
      const key = String(group || "Todos").toUpperCase();
      const code = String(eq.code || eq.equipment_code || "").toUpperCase();
      const text = `${code} ${eq.description || ""} ${eq.family || ""}`.toUpperCase();
      if(key.includes("TODOS")) return true;
      if(key.includes("BARRENACION")) return code.startsWith("JL") || code.startsWith("JA") || text.includes("JUMBO") || text.includes("BARREN") || text.includes("ANCLADOR");
      if(key.includes("REZAGADO")) return code.startsWith("ST") || text.includes("SCOOP") || text.includes("CATERPILLAR") || text.includes("EPROC") || text.includes("R1300") || text.includes("R1600") || text.includes("REZAG");
      return true;
    }
    function metric(period, worked, mp, mc, stops){
      const available = Math.max(Number(period || 0) - Number(mp || 0) - Number(mc || 0), 0);
      const availability = period > 0 ? Math.max(Math.min((available / period) * 100, 100), 0) : 0;
      const utilization = available > 0 ? Math.max(Math.min((Number(worked || 0) / available) * 100, 100), 0) : 0;
      const stopCount = Math.max(Number(stops || 0), 0);
      return {
        available,
        availability,
        utilization,
        tmef: stopCount ? (Number(worked || 0) / stopCount) : Number(worked || 0),
        tmpr: stopCount ? (Number(mc || 0) / stopCount) : 0,
      };
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
    function renderPortalSelectors(){
      const period = portal.period || {};
      const today = toIsoDate(new Date());
      if(!$("kpiStart").value) $("kpiStart").value = period.start || today;
      if(!$("kpiEnd").value) $("kpiEnd").value = period.end || today;
      if(!$("prBase").value) $("prBase").value = period.start || today;
      if(!$("bitStart").value) $("bitStart").value = period.start || today;
      if(!$("bitEnd").value) $("bitEnd").value = period.end || today;
      const rawGroups = portal.kpi_groups && portal.kpi_groups.length ? [...portal.kpi_groups] : ["Todos los equipos", "Equipos de Barrenacion", "Equipos de Rezagado"];
      ["KPI Aceites", "KPI Llantas"].forEach(group => { if(!rawGroups.includes(group)) rawGroups.push(group); });
      const groups = rawGroups.map(g => ({value:g, label:g}));
      const previousGroup = $("kpiGroup").value;
      $("kpiGroup").innerHTML = groups.map(g => `<option value="${esc(g.value)}">${esc(g.label)}</option>`).join("");
      $("kpiGroup").value = previousGroup && groups.some(g => g.value === previousGroup) ? previousGroup : groups[0]?.value || "";
      const equipmentOptions = portalEquipment().map(e => ({value:e.code || e.equipment_code, label:`${e.code || e.equipment_code} - ${e.description || e.family || ""}`}));
      setOptions("prEquipment", equipmentOptions, "Todos");
      setOptions("bitEquipment", equipmentOptions, "Todos");
    }
    function calculateKpiRows(){
      const group = $("kpiGroup").value || "Todos los equipos";
      const start = $("kpiStart").value;
      const end = $("kpiEnd").value;
      const settings = portal.settings || {};
      const shiftHours = Number(settings.shift_hours || 9);
      const dailyHours = shiftHours * Number(settings.turns_per_day || 2);
      const days = Math.max(Math.round((parseIsoDate(end) - parseIsoDate(start)) / 86400000) + 1, 1);
      const captures = (portal.captures || []).filter(c => inRange(c.work_date, start, end));
      const grouped = {};
      portalEquipment().filter(eq => groupMatches(eq, group)).forEach(eq => {
        const code = eq.code || eq.equipment_code || "";
        grouped[code] = {code, description:eq.description || "", family:eq.family || "", status:eq.status || "Disponible", period:days * dailyHours, worked:0, mp:0, mc:0, stops:0, unavailableCount:0};
      });
      captures.forEach(c => {
        const code = c.equipment_code || c.code || "";
        if(!grouped[code]) return;
        const row = grouped[code];
        const mp = Number(c.mp_hours || 0);
        let mc = Number(c.mc_hours || 0);
        if(unavailable(c.status)){
          const base = String(c.shift || "").toUpperCase() === "GENERAL" ? dailyHours : shiftHours;
          mc += Math.max(base - mp - mc, 0);
          row.unavailableCount += 1;
          row.status = c.status || "FUERA";
        }
        row.worked += Number(c.worked_hours || 0);
        row.mp += mp;
        row.mc += mc;
        row.stops += Number(c.stops || 0);
      });
      const rows = Object.values(grouped).sort((a,b) => a.code.localeCompare(b.code)).map(row => {
        const out = row.worked <= 0 && (row.unavailableCount > 0 || unavailable(row.status));
        const m = out ? {available:0, availability:0, utilization:0, tmef:0, tmpr:0} : metric(row.period, row.worked, row.mp, row.mc, row.stops);
        return {...row, ...m, out, availabilityText: out ? "FUERA" : pct(m.availability), utilizationText: out ? "FUERA" : pct(m.utilization)};
      });
      const totals = rows.reduce((acc, row) => {
        acc.period += row.period; acc.worked += row.worked; acc.mp += row.mp; acc.mc += row.mc; acc.stops += row.stops; acc.available += row.available;
        return acc;
      }, {period:0, worked:0, mp:0, mc:0, stops:0, available:0});
      totals.availability = totals.period ? (totals.available / totals.period) * 100 : 0;
      totals.utilization = totals.available ? (totals.worked / totals.available) * 100 : 0;
      totals.tmef = totals.stops ? totals.worked / totals.stops : totals.worked;
      totals.tmpr = totals.stops ? totals.mc / totals.stops : 0;
      return {group, start, end, rows, totals};
    }
    const kpiMetricTabs = [
      {key:"availability", label:"% Disponibilidad"},
      {key:"utilization", label:"% Utilizacion"},
      {key:"tmef", label:"TMEF"},
      {key:"tmpr", label:"TMPR"},
    ];
    function kpiMetricValue(row, metric){
      if(metric === "utilization") return Number(row.utilization || 0);
      if(metric === "tmef") return Number(row.tmef || 0);
      if(metric === "tmpr") return Number(row.tmpr || 0);
      return Number(row.availability || 0);
    }
    function kpiMetricText(row, metric){
      if((metric === "availability" || metric === "utilization") && row.out) return "FUERA";
      return one(kpiMetricValue(row, metric));
    }
    function kpiMetricAxisMax(metric, rows, target){
      if(metric === "availability" || metric === "utilization") return 120;
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
        tmef: Number(settings.meta_tmef || 8),
        tmpr: Number(settings.meta_tmpr || 4),
      };
      const axisMax = kpiMetricAxisMax(metric, report.rows, targets[metric]);
      const chartBars = report.rows.map(row => {
        const value = kpiMetricValue(row, metric);
        const h = Math.max(Math.min(value / Math.max(axisMax, 1), 1) * 210, 4);
        const outClass = row.out && (metric === "availability" || metric === "utilization") ? "out" : "";
        const title = `${row.code} ${kpiMetricTabs.find(item => item.key === metric)?.label || ""}: ${kpiMetricText(row, metric)}`;
        return `<div class="chart-bar ${outClass}" title="${esc(title)}"><span>${esc(kpiMetricText(row, metric))}</span><i style="--h:${h}px"></i><b>${esc(row.code)}</b></div>`;
      }).join("") || `<p class="muted">Sin datos KPI para el periodo.</p>`;
      const tabs = kpiMetricTabs.map(item => `<button type="button" data-kpi-metric="${esc(item.key)}" class="${item.key === metric ? "active" : ""}">${esc(item.label)}</button>`).join("");
      return `<div class="kpi-chart-head"><b>KPI</b><div class="kpi-mini-tabs">${tabs}</div></div><div class="chart-plot">${chartBars}</div>`;
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
      const cols = portal.oil_kpi && Array.isArray(portal.oil_kpi.columns) && portal.oil_kpi.columns.length ? portal.oil_kpi.columns : [
        {label:"15W40", key:"oil_motor_15w40"},
        {label:"ISO 68", key:"oil_hco_iso68"},
        {label:"SAE 30", key:"oil_trans_sae30"},
        {label:"SAE 50", key:"oil_sae50"},
        {label:"85W140", key:"oil_85w140"},
        {label:"ALMO", key:"almo_liters"},
        {label:"Refrigerante", key:"coolant_liters"},
      ];
      return cols;
    }
    function oilGroupFor(eq){
      const code = String(eq.code || eq.equipment_code || "").toUpperCase();
      const text = `${code} ${eq.description || ""} ${eq.family || ""}`.toUpperCase();
      if(code.startsWith("JL") || code.startsWith("JA") || text.includes("JUMBO") || text.includes("ANCLADOR")) return "BARRENACION";
      if(code.startsWith("ST") || text.includes("SCOOP") || text.includes("CATERPILLAR") || text.includes("EPROC") || text.includes("R1300") || text.includes("R1600")) return "REZAGADO";
      return "UTILITARIO";
    }
    function oilRowsForPeriod(){
      const start = $("kpiStart").value;
      const end = $("kpiEnd").value;
      const cols = oilColumns();
      const grouped = {};
      portalEquipment().forEach(eq => {
        const code = eq.code || eq.equipment_code || "";
        grouped[code] = {code, description:eq.description || "", group:oilGroupFor(eq), worked_hours:0, total_liters:0};
        cols.forEach(col => grouped[code][col.key] = 0);
      });
      (portal.captures || []).filter(row => inRange(row.work_date, start, end)).forEach(row => {
        const code = row.equipment_code || row.code || "";
        if(!grouped[code]) {
          grouped[code] = {code, description:"", group:"UTILITARIO", worked_hours:0, total_liters:0};
          cols.forEach(col => grouped[code][col.key] = 0);
        }
        grouped[code].worked_hours += Number(row.worked_hours || 0);
        cols.forEach(col => {
          const value = Number(row[col.key] || 0);
          grouped[code][col.key] += value;
          grouped[code].total_liters += value;
        });
      });
      const rows = Object.values(grouped).sort((a,b) => `${a.group} ${a.code}`.localeCompare(`${b.group} ${b.code}`));
      const totals = rows.reduce((acc, row) => {
        acc.worked_hours += row.worked_hours;
        acc.total_liters += row.total_liters;
        cols.forEach(col => acc[col.key] = (acc[col.key] || 0) + Number(row[col.key] || 0));
        return acc;
      }, {worked_hours:0, total_liters:0});
      return {start, end, cols, rows, totals};
    }
    function renderOilDashboard(){
      setDashboardMode("special");
      const report = oilRowsForPeriod();
      $("portalUpdated").textContent = portal.updated_at || portal.generated_at ? `Actualizado ${portal.updated_at || portal.generated_at}` : "Sin sincronizar";
      $("kpiTitle").textContent = `KPI Aceites | ${report.start} a ${report.end}`;
      const litersPerHour = report.totals.worked_hours ? report.totals.total_liters / report.totals.worked_hours : 0;
      const activeRows = report.rows.filter(row => row.worked_hours > 0 || row.total_liters > 0);
      $("kpiCards").innerHTML = [
        ["Equipos", `${activeRows.length}`, "con consumo o trabajo", 100, false],
        ["Litros total", `${one(report.totals.total_liters)} L`, "acumulado periodo", Math.min(report.totals.total_liters / 10, 100), false],
        ["Hrs trabajadas", `${one(report.totals.worked_hours)} h`, "capturas", Math.min(report.totals.worked_hours / 10, 100), false],
        ["L / hora", `${one(litersPerHour)}`, "consumo promedio", Math.min(litersPerHour * 20, 100), litersPerHour > 1.5],
      ].map(([label, value, note, width, bad]) => metricCardHtml(label, value, note, width, bad)).join("");
      const chartRows = [...report.rows].filter(row => row.total_liters > 0).sort((a,b) => b.total_liters - a.total_liters).slice(0,18);
      const maxValue = Math.max(...chartRows.map(row => row.total_liters), 1);
      $("kpiChart").innerHTML = chartRows.map(row => {
        const h = Math.max((row.total_liters / maxValue) * 210, 4);
        return `<div class="chart-bar" title="${esc(row.code)} ${one(row.total_liters)} L"><span>${one(row.total_liters)} L</span><i style="--h:${h}px"></i><b>${esc(row.code)}</b></div>`;
      }).join("") || `<p class="muted">Sin consumos de aceite en el periodo.</p>`;
      const headers = ["Equipo","Grupo","Hrs Trab", ...report.cols.map(col => col.label), "Total L"];
      $("kpiTable").innerHTML = `<thead><tr>${headers.map(h => `<th>${esc(h)}</th>`).join("")}</tr></thead><tbody>` +
        report.rows.map(row => `<tr><td>${esc(row.code)}</td><td>${esc(row.group)}</td><td>${one(row.worked_hours)}</td>${report.cols.map(col => `<td>${one(row[col.key])}</td>`).join("")}<td>${one(row.total_liters)}</td></tr>`).join("") +
        `<tr><td><b>Total</b></td><td></td><td><b>${one(report.totals.worked_hours)}</b></td>${report.cols.map(col => `<td><b>${one(report.totals[col.key])}</b></td>`).join("")}<td><b>${one(report.totals.total_liters)}</b></td></tr></tbody>`;
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
    function renderDashboard(){
      const selectedGroup = $("kpiGroup").value || "";
      if(selectedGroup === "KPI Aceites") {
        renderOilDashboard();
        return;
      }
      if(selectedGroup === "KPI Llantas") {
        renderTireDashboard();
        return;
      }
      setDashboardMode("format");
      const report = calculateKpiRows();
      const settings = portal.settings || {};
      $("portalUpdated").textContent = portal.updated_at || portal.generated_at ? `Actualizado ${portal.updated_at || portal.generated_at}` : "Sin sincronizar";
      $("kpiTitle").textContent = `${report.group} | ${report.start} a ${report.end}`;
      const metaAvailability = Number(settings.meta_availability || 85);
      const metaUtilization = Number(settings.meta_utilization || 75);
      const metaTmef = Number(settings.meta_tmef || 8);
      const metaTmpr = Number(settings.meta_tmpr || 4);
      $("kpiCards").innerHTML = [
        metricCardHtml("% Disponibilidad", pct(report.totals.availability), `Meta ${pct(metaAvailability)}`, report.totals.availability, report.totals.availability < metaAvailability),
        metricCardHtml("Meta", pct(metaAvailability), `${one(report.totals.availability - metaAvailability)}%`, metaAvailability, false),
        metricCardHtml("% Utilizacion", pct(report.totals.utilization), `Meta ${pct(metaUtilization)}`, report.totals.utilization, report.totals.utilization < metaUtilization),
        metricCardHtml("Meta", pct(metaUtilization), `${one(report.totals.utilization - metaUtilization)}%`, metaUtilization, report.totals.utilization < metaUtilization),
      ].join("");
      $("kpiSideCards").innerHTML = [
        metricCardHtml("TMEF", `${one(report.totals.tmef)} h`, `Meta ${one(metaTmef)} h`, Math.min((report.totals.tmef / Math.max(metaTmef, 1)) * 100, 100), report.totals.tmef < metaTmef),
        metricCardHtml("Meta", `${one(metaTmef)} h`, `${one(report.totals.tmef - metaTmef)} h`, 100, false),
        metricCardHtml("TMPR", `${one(report.totals.tmpr)} h`, `Meta ${one(metaTmpr)} h`, Math.min((report.totals.tmpr / Math.max(metaTmpr, 1)) * 100, 100), report.totals.tmpr > metaTmpr),
        metricCardHtml("Meta", `${one(metaTmpr)} h`, `${one(metaTmpr - report.totals.tmpr)} h`, 100, report.totals.tmpr > metaTmpr),
      ].join("");
      $("kpiChart").innerHTML = kpiMetricChartHtml(report, settings);
      bindKpiMetricTabs();
      $("kpiTable").innerHTML = `<thead><tr><th># Eco</th><th>Equipo</th><th>Hrs periodo</th><th>Hrs MP</th><th>Hrs MC</th><th>Hrs trab</th><th># Paradas</th><th>% Disp</th><th>% Util</th><th>TMEF</th><th>TMPR</th><th>Estatus</th></tr></thead><tbody>` +
        report.rows.map(row => `<tr><td>${esc(row.code)}</td><td>${esc(row.description)}</td><td>${one(row.period)}</td><td>${one(row.mp)}</td><td>${one(row.mc)}</td><td>${one(row.worked)}</td><td>${num(row.stops)}</td><td>${esc(row.availabilityText)}</td><td>${esc(row.utilizationText)}</td><td>${one(row.tmef)}</td><td>${one(row.tmpr)}</td><td>${esc(row.out ? "FUERA" : row.status)}</td></tr>`).join("") +
        `<tr><td></td><td><b>Total ${esc(report.group)}</b></td><td><b>${one(report.totals.period)}</b></td><td><b>${one(report.totals.mp)}</b></td><td><b>${one(report.totals.mc)}</b></td><td><b>${one(report.totals.worked)}</b></td><td><b>${num(report.totals.stops)}</b></td><td><b>${pct(report.totals.availability)}</b></td><td><b>${pct(report.totals.utilization)}</b></td><td><b>${one(report.totals.tmef)}</b></td><td><b>${one(report.totals.tmpr)}</b></td><td></td></tr></tbody>`;
    }
    function filteredPreventives(){
      const [start, end] = periodRange($("prPeriod").value, $("prBase").value);
      const selected = $("prEquipment").value;
      const search = ($("prSearch").value || "").toUpperCase();
      const rows = (portal.preventives || []).filter(row => {
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
    function conditionClass(condition){
      const text = String(condition || "").toUpperCase();
      if(text.includes("FUERA") || text.includes("NO DISP")) return "cond-out";
      if(text.includes("OPERATIVA") || text.includes("REPARACION") || text.includes("STAND")) return "cond-warn";
      if(text.includes("DISPONIBLE")) return "cond-ok";
      return "";
    }
    function renderDisponibilidad(){
      const search = ($("dispSearch").value || "").toUpperCase();
      const status = $("dispStatus").value;
      const rows = (portal.availability || []).filter(row => {
        const text = [row.category,row.equipment,row.eco,row.condition,row.observations].join(" ").toUpperCase();
        return (!status || String(row.condition || "").toUpperCase().includes(status)) && (!search || text.includes(search));
      });
      $("dispTable").innerHTML = `<thead><tr><th>Categoria</th><th>Equipo</th><th>No ECO</th><th>Condicion</th><th>Observaciones</th></tr></thead><tbody>` +
        rows.map(row => `<tr><td>${esc(row.category)}</td><td>${esc(row.equipment)}</td><td>${esc(row.eco)}</td><td class="condition-cell ${conditionClass(row.condition)}">${esc(row.condition)}</td><td class="${Number(row.highlight_observation || 0) ? "highlight" : ""}">${esc(row.observations)}</td></tr>`).join("") +
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
      if(!hasApiKey(true)) return;
      if(!confirm(`Se eliminara la requisicion ${$("reqFolio").value}.`)) return;
      const r = await fetch(`/api/requisitions/${currentReqId}`, {method:"DELETE", headers:headers()});
      if(!r.ok) return alert(await apiError(r));
      currentReqId = null;
      await load();
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
      $("reqListTable").innerHTML = `<thead><tr><th>Folio</th><th>Fecha</th><th>Equipo</th><th>Estatus</th><th>Partidas</th></tr></thead><tbody>` +
        rows.map(row => `<tr data-req-id="${row.id}" style="cursor:pointer"><td>${esc(row.folio)}</td><td>${esc(row.request_date)}</td><td>${esc(row.equipment)}</td><td>${esc(row.status)}</td><td>${num(row.items_count)}</td></tr>`).join("") + `</tbody>`;
      document.querySelectorAll("[data-req-id]").forEach(row => row.addEventListener("click", () => loadReq(row.dataset.reqId).catch(showError)));
    }
    function renderRequisiciones(){
      renderReqProducts();
      renderReqList();
      renderReqItems();
    }
    function renderAll(){
      renderStats();
      renderSelectors();
      renderPortalSelectors();
      renderDashboard();
      renderPreventives();
      renderBitacora();
      renderDisponibilidad();
      renderRequisiciones();
      renderFilters();
      renderInventory();
      renderMovements();
    }
    document.querySelectorAll(".tabs button").forEach(btn => btn.addEventListener("click", () => {
      document.querySelectorAll(".tabs button").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
      btn.classList.add("active"); $(btn.dataset.tab).classList.add("active");
    }));
    ["kpiGroup","kpiStart","kpiEnd"].forEach(id => $(id).addEventListener("change", renderDashboard));
    $("renderKpiBtn").addEventListener("click", renderDashboard);
    $("printKpiBtn").addEventListener("click", () => window.print());
    ["prPeriod","prBase","prEquipment"].forEach(id => $(id).addEventListener("change", renderPreventives));
    $("prSearch").addEventListener("input", renderPreventives);
    $("renderPrBtn").addEventListener("click", renderPreventives);
    ["bitEquipment","bitStart","bitEnd"].forEach(id => $(id).addEventListener("change", renderBitacora));
    $("bitSearch").addEventListener("input", renderBitacora);
    $("renderBitBtn").addEventListener("click", renderBitacora);
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
def filter_warehouse_page() -> str:
    return WAREHOUSE_HTML


@app.get("/api/filter-inventory")
def get_filter_inventory(_auth: str | None = Header(default=None, alias="X-MGA-API-Key")) -> dict[str, Any]:
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
        "availability": len(availability) if isinstance(availability, list) else 0,
    }


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
