from __future__ import annotations

import base64
import os
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
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


engine = create_engine(database_url(), pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)
Base.metadata.create_all(engine)

app = FastAPI(title="MGA Cloud Sync", version="1.2.4")
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


def inventory_item_payload(item: FilterInventoryItem) -> dict[str, Any]:
    return {
        "id": item.id,
        "part_key": item.part_key,
        "part_number": item.part_number,
        "description": item.description,
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


def filter_match_keys(item: dict[str, Any]) -> list[str]:
    keys: list[str] = []
    for field in ("part_number", "donaldson_part", "matched_part"):
        key = normalize_part_key(item.get(field))
        if key and key not in keys:
            keys.append(key)
    return keys


def catalog_with_inventory(session: Session) -> dict[str, Any]:
    catalog = latest_catalog_payload(session)
    equipment = catalog.get("equipment") if isinstance(catalog, dict) else []
    if not isinstance(equipment, list):
        equipment = []
    inventory = {item.part_key: item for item in session.scalars(select(FilterInventoryItem)).all()}
    inventory_list = [inventory_item_payload(item) for item in sorted(inventory.values(), key=lambda row: row.part_number)]
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
            filter_item["stock_description"] = stock.description
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
        return catalog_with_inventory(session)


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
    body::before { content:""; position:fixed; inset:0; pointer-events:none; background-image:url('/static/mga-logo.png'); background-size:260px auto; background-repeat:repeat; opacity:.025; transform:rotate(-7deg) scale(1.12); transform-origin:center; z-index:-1; }
    .hero { position:relative; overflow:hidden; color:white; padding:18px 28px 24px; display:flex; justify-content:space-between; gap:22px; align-items:center; background:radial-gradient(circle at 15% 0%, rgba(255,255,255,.18), transparent 28%), linear-gradient(135deg,#050b18 0%, var(--blue) 56%, #0b1735 100%); box-shadow:0 18px 44px rgba(7,31,73,.24); }
    .hero::after { content:""; position:absolute; right:40px; bottom:-52px; width:430px; height:190px; background:url('/static/mga-logo.png') center/contain no-repeat; opacity:.08; pointer-events:none; }
    .brand { position:relative; z-index:1; display:flex; align-items:center; gap:18px; min-width:0; }
    .brand-logo { width:190px; max-width:36vw; height:auto; border-radius:8px; box-shadow:0 14px 34px rgba(0,0,0,.35); border:1px solid rgba(255,255,255,.18); background:#000; }
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
    .logo-ribbon { display:grid; grid-template-columns:repeat(4, minmax(140px, 1fr)); gap:10px; opacity:.9; }
    .logo-ribbon span { min-height:34px; border:1px solid var(--line); border-radius:8px; background:rgba(255,255,255,.72) url('/static/mga-logo.png') center/130px auto no-repeat; box-shadow:0 8px 22px rgba(7,31,73,.06); }
    .panel { position:relative; overflow:hidden; background:rgba(255,255,255,.92); border:1px solid rgba(216,222,232,.9); border-radius:8px; padding:16px; box-shadow:var(--shadow); }
    .panel::after { content:""; position:absolute; right:18px; bottom:14px; width:160px; height:54px; background:url('/static/mga-logo.png') center/contain no-repeat; opacity:.035; pointer-events:none; }
    .toolbar { display:grid; grid-template-columns:repeat(5, minmax(140px, 1fr)); gap:10px; align-items:end; }
    label { display:grid; gap:4px; color:#344054; font-size:12px; font-weight:700; }
    input, select, textarea { width:100%; padding:9px 10px; border:1px solid #cbd5e1; border-radius:6px; font:inherit; background:white; outline:none; transition:border .15s ease, box-shadow .15s ease; }
    input:focus, select:focus, textarea:focus { border-color:var(--teal); box-shadow:0 0 0 3px rgba(0,156,154,.14); }
    .stats { display:grid; grid-template-columns:repeat(5, 1fr); gap:10px; }
    .stat { position:relative; overflow:hidden; background:linear-gradient(180deg,#fff,#f8fbff); border:1px solid var(--line); padding:14px 15px; border-radius:8px; box-shadow:0 10px 26px rgba(7,31,73,.08); }
    .stat::after { content:""; position:absolute; right:10px; top:10px; width:78px; height:26px; background:url('/static/mga-logo.png') center/contain no-repeat; opacity:.07; }
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
    .wide { grid-column:1 / -1; }
    @media (max-width: 900px) { .hero, .grid2 { display:block; } .brand { align-items:flex-start; } .brand-logo { width:145px; margin-bottom:10px; } .toolbar, .movement-grid, .stats, .logo-ribbon { grid-template-columns:1fr; } header input { min-width:0; margin-top:10px; } .key-card { margin-top:14px; min-width:0; } }
  </style>
</head>
<body>
  <header class="hero">
    <div class="brand">
      <img class="brand-logo" src="/static/mga-logo.png" alt="MGA Contratista Minera">
      <div><h1>Inventario de filtros</h1><p>Almacen conectado a FS Filtros servicio</p></div>
    </div>
    <div class="key-card"><label>Clave para editar<input id="apiKey" type="password" placeholder="Pegar clave aqui"></label></div>
  </header>
  <main>
    <nav class="tabs">
      <button class="active" data-tab="equipos">Filtros por equipo</button>
      <button data-tab="inventario">Concentrado / movimientos</button>
      <button data-tab="importar">Importar / exportar</button>
    </nav>
    <div class="logo-ribbon" aria-hidden="true"><span></span><span></span><span></span><span></span></div>
    <section class="stats" id="stats"></section>
    <section id="equipos" class="view active">
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
    function statusClass(s){ return s === "Disponible" ? "ok" : (s === "Faltante" ? "bad" : "warn"); }
    async function load(){
      const r = await fetch("/api/filter-inventory", {headers: headers()});
      if(!r.ok) throw new Error(await apiError(r));
      data = await r.json();
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
      const rows = (data.inventory || []).filter(i => !search || [i.part_number,i.description,i.location].join(" ").toUpperCase().includes(search));
      $("inventoryTable").innerHTML = `<thead><tr><th>No. parte</th><th>Descripcion</th><th>Exist.</th><th>Unidad</th><th>Min.</th><th>Ubicacion</th><th>Actualizado</th></tr></thead><tbody>` +
        rows.map(i => `<tr><td>${esc(i.part_number)}</td><td>${esc(i.description)}</td><td>${num(i.quantity)}</td><td>${esc(i.unit||"PZA")}</td><td>${num(i.min_stock)}</td><td>${esc(i.location)}</td><td>${esc(i.updated_at)}</td></tr>`).join("") +
        `</tbody>`;
    }
    function renderMovements(){
      const rows = data.movements || [];
      $("movementTable").innerHTML = `<thead><tr><th>Fecha</th><th>Parte</th><th>Tipo</th><th>Cant.</th><th>Saldo</th><th>Ref.</th></tr></thead><tbody>` +
        rows.map(m => `<tr><td>${esc(m.movement_date)}</td><td>${esc(m.part_number)}</td><td>${esc(m.movement_type)}</td><td>${num(m.quantity)}</td><td>${num(m.balance_after)}</td><td>${esc(m.reference)}</td></tr>`).join("") +
        `</tbody>`;
    }
    function renderAll(){ renderStats(); renderSelectors(); renderFilters(); renderInventory(); renderMovements(); }
    document.querySelectorAll(".tabs button").forEach(btn => btn.addEventListener("click", () => {
      document.querySelectorAll(".tabs button").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
      btn.classList.add("active"); $(btn.dataset.tab).classList.add("active");
    }));
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
    ws.append(["No. parte", "Descripcion", "Cantidad", "Unidad", "Minimo", "Ubicacion", "Actualizado"])
    with SessionLocal() as session:
        for item in session.scalars(select(FilterInventoryItem).order_by(FilterInventoryItem.part_number.asc())).all():
            ws.append([item.part_number, item.description, item.quantity, item.unit, item.min_stock, item.location, item.updated_at.isoformat(timespec="seconds")])
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
