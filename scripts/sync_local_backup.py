from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description="Restore an MGA desktop SQLite backup into the cloud database.")
    result.add_argument("--source", required=True, help="Desktop SQLite backup path.")
    result.add_argument("--database-url", default="", help="Target PostgreSQL DATABASE_URL.")
    result.add_argument("--api-url", default="", help="Target MGA API base URL.")
    result.add_argument("--api-key", default="", help="MGA API key. Defaults to MGA_API_KEY.")
    result.add_argument("--dry-run", action="store_true", help="Inspect and count source data without writing.")
    result.add_argument("--replace", action="store_true", help="Replace cloud operational data and snapshots.")
    result.add_argument("--portal-only", action="store_true", help="Publish only catalog and portal snapshots.")
    return result


def rows(connection: sqlite3.Connection, query: str, params: tuple[Any, ...] = ()) -> list[dict[str, Any]]:
    return [dict(row) for row in connection.execute(query, params).fetchall()]


def table_exists(connection: sqlite3.Connection, table: str) -> bool:
    row = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (table,)
    ).fetchone()
    return row is not None


def parse_datetime(value: Any) -> datetime:
    text = str(value or "").strip()
    if text:
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
            return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return datetime.now(timezone.utc)


def number(value: Any, default: float = 0) -> float:
    try:
        return float(value if value not in (None, "") else default)
    except (TypeError, ValueError):
        return default


def integer(value: Any, default: int = 0) -> int:
    return int(number(value, default))


def config_value(config: dict[str, str], key: str, default: Any) -> Any:
    value = config.get(key)
    return default if value in (None, "") else value


def source_payload(connection: sqlite3.Connection) -> dict[str, Any]:
    equipment_rows = rows(connection, "SELECT * FROM equipment ORDER BY code")
    component_rows = rows(connection, "SELECT * FROM component ORDER BY equipment_id, id")
    filter_rows = rows(connection, "SELECT * FROM service_filter ORDER BY equipment_id, id")
    evidence_counts = {
        int(row["capture_id"]): int(row["total"])
        for row in rows(connection, "SELECT capture_id, COUNT(*) AS total FROM capture_evidence GROUP BY capture_id")
    }

    equipment_by_id = {int(row["id"]): row for row in equipment_rows}
    components_by_id = {int(row["id"]): row for row in component_rows}
    components_by_equipment: dict[int, list[dict[str, Any]]] = defaultdict(list)
    filters_by_equipment: dict[int, list[dict[str, Any]]] = defaultdict(list)

    for row in component_rows:
        components_by_equipment[int(row["equipment_id"])].append(
            {
                "id": row["id"],
                "name": row["name"],
                "meter_type": row["meter_type"],
                "current_meter": number(row["current_meter"]),
                "service_interval": number(row["service_interval"]),
                "last_service_meter": number(row["last_service_meter"]),
                "daily_average": number(row["daily_average"]),
                "active": integer(row["active"], 1),
            }
        )

    for row in filter_rows:
        filters_by_equipment[int(row["equipment_id"])].append(
            {
                "id": row["id"],
                "service_interval": row["service_interval"],
                "item_type": row["item_type"],
                "part_number": row["part_number"],
                "donaldson_part": row.get("donaldson_part") or "",
                "donaldson_url": row.get("donaldson_url") or "",
                "description": row["description"],
                "quantity": number(row["quantity"], 1),
                "unit": row["unit"],
                "notes": row["notes"],
                "active": integer(row["active"], 1),
            }
        )

    equipment = []
    for row in equipment_rows:
        equipment_id = int(row["id"])
        equipment.append(
            {
                "id": equipment_id,
                "code": row["code"],
                "description": row["description"],
                "family": row["family"],
                "status": row["status"],
                "active": integer(row["active"], 1),
                "components": components_by_equipment[equipment_id],
                "filters": filters_by_equipment[equipment_id],
            }
        )

    captures = []
    capture_rows = rows(connection, "SELECT * FROM capture ORDER BY work_date, id")
    for row in capture_rows:
        equipment_row = equipment_by_id.get(int(row["equipment_id"])) or {}
        component_row = components_by_id.get(integer(row.get("component_id"))) or {}
        captures.append(
            {
                "id": row["id"],
                "captured_at": row["captured_at"],
                "work_date": row["work_date"],
                "shift": row.get("shift") or "General",
                "equipment_code": equipment_row.get("code") or "",
                "equipment_description": equipment_row.get("description") or "",
                "component": component_row.get("name") or "",
                "hi": number(row.get("hi")),
                "hf": number(row.get("hf")),
                "worked_hours": number(row.get("worked_hours")),
                "mp_hours": number(row.get("mp_hours")),
                "mc_hours": number(row.get("mc_hours")),
                "standby_hours": number(row.get("standby_hours")),
                "downtime_hours": number(row.get("downtime_hours")),
                "stops": integer(row.get("stops")),
                "oil_liters": number(row.get("oil_liters")),
                "oil_motor_15w40": number(row.get("oil_motor_15w40")),
                "oil_hco_iso68": number(row.get("oil_hco_iso68")),
                "oil_trans_sae30": number(row.get("oil_trans_sae30")),
                "oil_sae50": number(row.get("oil_sae50")),
                "oil_85w140": number(row.get("oil_85w140")),
                "almo_liters": number(row.get("almo_liters")),
                "coolant_liters": number(row.get("coolant_liters")),
                "fault": row.get("fault") or "",
                "wear": row.get("wear") or "",
                "status": row.get("status") or "Disponible",
                "observations": row.get("observations") or "",
                "evidence_count": evidence_counts.get(int(row["id"]), 0),
            }
        )

    preventive_rows = rows(connection, "SELECT * FROM preventive_service ORDER BY due_date, id")
    preventives = []
    for row in preventive_rows:
        equipment_row = equipment_by_id.get(int(row["equipment_id"])) or {}
        component_row = components_by_id.get(int(row["component_id"])) or {}
        preventives.append(
            {
                **row,
                "equipment_code": equipment_row.get("code") or "",
                "equipment_description": equipment_row.get("description") or "",
                "component": component_row.get("name") or "",
            }
        )

    availability = rows(connection, "SELECT * FROM availability_report_row ORDER BY sort_order, id")
    inventory = rows(connection, "SELECT * FROM filter_inventory ORDER BY part_number")
    inventory_movements = rows(connection, "SELECT * FROM filter_inventory_movement ORDER BY id")
    requisitions = rows(connection, "SELECT * FROM requisition ORDER BY id")
    requisition_items = rows(connection, "SELECT * FROM requisition_item ORDER BY requisition_id, sort_order, id")
    hoses = rows(connection, "SELECT * FROM hose_change ORDER BY id")
    diesel_records = rows(connection, "SELECT * FROM diesel_record ORDER BY work_date, id")
    diesel_days = rows(connection, "SELECT * FROM diesel_day ORDER BY work_date")
    tires = rows(
        connection,
        """
        SELECT t.*, e.code AS equipment_code
        FROM tire t LEFT JOIN equipment e ON e.id = t.equipment_id
        ORDER BY t.tire_code
        """,
    )
    tire_events = rows(connection, "SELECT * FROM tire_event ORDER BY event_date, id")
    settings_rows = rows(connection, "SELECT key, value FROM config")
    settings = {str(row["key"]): str(row["value"]) for row in settings_rows}

    return {
        "equipment": equipment,
        "captures": captures,
        "preventives": preventives,
        "availability": availability,
        "inventory": inventory,
        "inventory_movements": inventory_movements,
        "requisitions": requisitions,
        "requisition_items": requisition_items,
        "hoses": hoses,
        "diesel_records": diesel_records,
        "diesel_days": diesel_days,
        "tires": tires,
        "tire_events": tire_events,
        "settings": settings,
    }


def portal_payload(payload: dict[str, Any]) -> dict[str, Any]:
    now = datetime.now(timezone.utc)
    settings = payload["settings"]
    year = integer(config_value(settings, "ANIO", now.year), now.year)
    month = integer(config_value(settings, "MES_NUMERO", now.month), now.month)
    capture_dates = [str(row.get("work_date") or "") for row in payload["captures"] if row.get("work_date")]
    period_end = max(capture_dates) if capture_dates else now.date().isoformat()
    return {
        "ok": True,
        "source": "cloud-portal",
        "generated_at": now.isoformat(timespec="seconds"),
        "period": {
            "start": f"{year:04d}-{month:02d}-01",
            "end": period_end,
            "year": year,
            "month": month,
        },
        "settings": {
            "shift_hours": number(config_value(settings, "HORAS_TURNO", 9), 9),
            "turns_per_day": number(config_value(settings, "TURNOS_DIA", 2), 2),
            "meta_availability": number(config_value(settings, "META_DISPONIBILIDAD", 85), 85),
            "meta_utilization": number(config_value(settings, "META_UTILIZACION", 75), 75),
            "meta_tmef": number(config_value(settings, "META_TMEF", 8), 8),
            "meta_tmpr": number(config_value(settings, "META_TMPR", 4), 4),
            "meta_diesel_lh": number(config_value(settings, "META_DIESEL_LH", 25), 25),
        },
        "equipment": payload["equipment"],
        "captures": payload["captures"],
        "preventives": payload["preventives"],
        "service_history": payload["preventives"],
        "availability": payload["availability"],
        "kpi_groups": [
            "Todos los equipos",
            "Equipos de Barrenacion",
            "Equipos de Rezagado",
            "KPI Aceites",
            "KPI Llantas",
        ],
        "kpi_reports": {},
        "oil_kpi": {"rows": [], "totals": {}, "columns": []},
        "tire_kpi": {
            "rows": payload["tires"],
            "events": payload["tire_events"],
            "summary": {"tires": len(payload["tires"]), "events": len(payload["tire_events"])},
        },
        "diesel": {
            # Operational diesel endpoints are authoritative after restore.
            "records": [],
            "days": [],
            "rows": [],
            "totals": {},
        },
    }


def api_post(base_url: str, api_key: str, path: str, payload: dict[str, Any]) -> dict[str, Any]:
    request = urllib.request.Request(
        f"{base_url.rstrip('/')}{path}",
        data=json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8"),
        headers={"Content-Type": "application/json", "X-MGA-API-Key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{path} returned HTTP {exc.code}: {detail[:500]}") from exc
    result = json.loads(body) if body else {}
    if not isinstance(result, dict):
        raise RuntimeError(f"{path} returned an invalid response.")
    return result


def restore_via_api(base_url: str, api_key: str, payload: dict[str, Any]) -> dict[str, int]:
    api_post(base_url, api_key, "/api/catalog", {"equipment": payload["equipment"]})
    if payload.get("portal_only"):
        api_post(base_url, api_key, "/api/portal/snapshot", portal_payload(payload))
        return {
            "equipment": len(payload["equipment"]),
            "captures": len(payload["captures"]),
            "preventives": len(payload["preventives"]),
            "availability": len(payload["availability"]),
        }
    api_post(
        base_url,
        api_key,
        "/api/filter-inventory/snapshot",
        {
            "inventory": [
                {
                    "part_number": row.get("part_number") or "",
                    "description": row.get("description") or "",
                    "quantity": number(row.get("quantity")),
                    "unit": "PZA",
                    "min_stock": 0,
                    "location": "",
                    "source_file": row.get("source_file") or "desktop-backup",
                }
                for row in payload["inventory"]
            ]
        },
    )

    requisition_items: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for item in payload["requisition_items"]:
        requisition_items[int(item["requisition_id"])].append(item)
    for index, requisition in enumerate(payload["requisitions"], 1):
        api_post(
            base_url,
            api_key,
            "/api/requisitions",
            {
                **requisition,
                "items": requisition_items[int(requisition["id"])],
            },
        )
        if index % 10 == 0:
            print(f"Requisitions synchronized: {index}/{len(payload['requisitions'])}")

    for hose in payload["hoses"]:
        restored = dict(hose)
        restored.pop("id", None)
        restored["external_id"] = restored.get("external_id") or f"desktop-{hose['id']}"
        api_post(base_url, api_key, "/api/hose-changes", restored)

    for index, record in enumerate(payload["diesel_records"], 1):
        restored = dict(record)
        restored.pop("id", None)
        api_post(base_url, api_key, "/api/diesel/records", restored)
        if index % 50 == 0:
            print(f"Diesel records synchronized: {index}/{len(payload['diesel_records'])}")
    for day in payload["diesel_days"]:
        restored = dict(day)
        restored.pop("id", None)
        api_post(base_url, api_key, "/api/diesel/days", restored)

    api_post(base_url, api_key, "/api/portal/snapshot", portal_payload(payload))
    return {
        "equipment": len(payload["equipment"]),
        "captures": len(payload["captures"]),
        "preventives": len(payload["preventives"]),
        "availability": len(payload["availability"]),
        "inventory": len(payload["inventory"]),
        "requisitions": len(payload["requisitions"]),
        "hoses": len(payload["hoses"]),
        "diesel_records": len(payload["diesel_records"]),
        "diesel_days": len(payload["diesel_days"]),
    }


def main() -> int:
    args = parser().parse_args()
    source = Path(args.source).expanduser().resolve()
    if not source.is_file():
        raise SystemExit(f"Source database not found: {source}")
    if not args.dry_run and not args.replace:
        raise SystemExit("Use --replace to confirm replacement of cloud data, or use --dry-run.")

    with sqlite3.connect(f"file:{source.as_posix()}?mode=ro", uri=True) as connection:
        connection.row_factory = sqlite3.Row
        required = {"equipment", "capture", "config", "filter_inventory", "requisition", "diesel_record"}
        missing = sorted(table for table in required if not table_exists(connection, table))
        if missing:
            raise SystemExit(f"Invalid MGA backup; missing tables: {', '.join(missing)}")
        payload = source_payload(connection)

    counts = {key: len(value) for key, value in payload.items() if isinstance(value, list)}
    print(json.dumps({"source": str(source), "counts": counts}, ensure_ascii=False, indent=2))
    if args.dry_run:
        return 0
    if args.api_url:
        api_key = args.api_key or os.getenv("MGA_API_KEY", "")
        if not api_key:
            raise SystemExit("An --api-key or MGA_API_KEY is required for API synchronization.")
        payload["portal_only"] = args.portal_only
        restored = restore_via_api(args.api_url, api_key, payload)
        print(json.dumps({"ok": True, "restored": restored}, ensure_ascii=False, indent=2))
        return 0
    if args.database_url:
        os.environ["DATABASE_URL"] = args.database_url
    if not os.getenv("DATABASE_URL", "").startswith(("postgres://", "postgresql://")):
        raise SystemExit("A PostgreSQL --database-url is required for the live restore.")

    project_root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(project_root))
    from render_backend.main import (  # noqa: PLC0415
        CaptureDeletion,
        CatalogSnapshot,
        CloudRequisition,
        CloudRequisitionItem,
        DieselDay,
        DieselDeletedDay,
        DieselDeletedRecord,
        DieselRecord,
        FilterInventoryItem,
        FilterInventoryMovement,
        HoseChange,
        MobileCapture,
        MobilePhoto,
        PortalSnapshot,
        SessionLocal,
        diesel_payload,
        json_dumps,
        normalize_part_key,
        utc_now,
    )

    now = utc_now()
    with SessionLocal() as session:
        for model in (
            MobilePhoto,
            CaptureDeletion,
            MobileCapture,
            FilterInventoryMovement,
            FilterInventoryItem,
            CloudRequisitionItem,
            CloudRequisition,
            HoseChange,
            DieselDeletedRecord,
            DieselDeletedDay,
            DieselRecord,
            DieselDay,
            PortalSnapshot,
            CatalogSnapshot,
        ):
            session.query(model).delete(synchronize_session=False)

        for capture in payload["captures"]:
            session.add(
                MobileCapture(
                    mobile_id=f"desktop-restore-{capture['id']}",
                    source_device="desktop-backup",
                    user_name="restauracion",
                    equipment_code=str(capture["equipment_code"]),
                    component_name=str(capture["component"]),
                    work_date=str(capture["work_date"]),
                    created_at=parse_datetime(capture.get("captured_at")),
                    received_at=now,
                    desktop_imported_at=now,
                    payload_json=json_dumps(capture),
                )
            )

        inventory_by_key: dict[str, FilterInventoryItem] = {}
        for item in payload["inventory"]:
            key = normalize_part_key(item.get("part_number"))
            if not key:
                continue
            cloud_item = FilterInventoryItem(
                part_key=key,
                part_number=str(item.get("part_number") or ""),
                description=str(item.get("description") or ""),
                quantity=number(item.get("quantity")),
                unit="PZA",
                min_stock=0,
                location="",
                source_file=str(item.get("source_file") or "desktop-backup"),
                updated_at=parse_datetime(item.get("imported_at")),
            )
            session.add(cloud_item)
            inventory_by_key[key] = cloud_item
        session.flush()
        for movement in payload["inventory_movements"]:
            item = inventory_by_key.get(normalize_part_key(movement.get("part_number")))
            if item is None:
                continue
            session.add(
                FilterInventoryMovement(
                    item_id=item.id,
                    movement_date=str(movement.get("movement_date") or ""),
                    movement_type=str(movement.get("movement_type") or "AJUSTE"),
                    quantity=number(movement.get("quantity")),
                    balance_after=number(movement.get("balance_after")),
                    reference=str(movement.get("reference") or ""),
                    notes=str(movement.get("notes") or movement.get("description") or ""),
                    created_by="desktop-backup",
                    created_at=parse_datetime(movement.get("created_at")),
                )
            )

        items_by_requisition: dict[int, list[dict[str, Any]]] = defaultdict(list)
        for item in payload["requisition_items"]:
            items_by_requisition[int(item["requisition_id"])].append(item)
        for requisition in payload["requisitions"]:
            cloud_requisition = CloudRequisition(
                folio=str(requisition["folio"]),
                created_at=parse_datetime(requisition.get("created_at")),
                updated_at=now,
                request_date=str(requisition.get("request_date") or ""),
                authorization_date=str(requisition.get("authorization_date") or ""),
                equipment=str(requisition.get("equipment") or ""),
                cost_center=str(requisition.get("cost_center") or ""),
                request_area=str(requisition.get("request_area") or ""),
                location=str(requisition.get("location") or ""),
                requesting_unit=str(requisition.get("requesting_unit") or ""),
                operating_unit=str(requisition.get("operating_unit") or ""),
                priority=str(requisition.get("priority") or "URGENTE"),
                recommendation=str(requisition.get("recommendation") or "ORIGINAL"),
                status=str(requisition.get("status") or "Abierta"),
                notes=str(requisition.get("notes") or ""),
            )
            session.add(cloud_requisition)
            session.flush()
            for item in items_by_requisition[int(requisition["id"])]:
                session.add(
                    CloudRequisitionItem(
                        requisition_id=cloud_requisition.id,
                        quantity=number(item.get("quantity"), 1),
                        unit=str(item.get("unit") or "PZA"),
                        part_number=str(item.get("part_number") or ""),
                        description=str(item.get("description") or ""),
                        sort_order=integer(item.get("sort_order")),
                        active=integer(item.get("active"), 1),
                    )
                )

        for hose in payload["hoses"]:
            session.add(
                HoseChange(
                    change_date=str(hose.get("change_date") or ""),
                    equipment=str(hose.get("equipment") or ""),
                    system=str(hose.get("system") or ""),
                    part_type=str(hose.get("part_type") or "MANGUERA"),
                    diameter=str(hose.get("diameter") or ""),
                    length_m=number(hose.get("length_m")),
                    quantity=number(hose.get("quantity"), 1),
                    unit_cost=number(hose.get("unit_cost")),
                    estimated_life_days=number(hose.get("estimated_life_days"), 30),
                    estimated_weekly_qty=number(hose.get("estimated_weekly_qty")),
                    failure_reason=str(hose.get("failure_reason") or ""),
                    technician=str(hose.get("technician") or ""),
                    notes=str(hose.get("notes") or ""),
                    source="desktop-backup",
                    external_id=str(hose.get("external_id") or f"desktop-{hose['id']}"),
                    created_at=parse_datetime(hose.get("created_at")),
                    updated_at=parse_datetime(hose.get("updated_at")),
                )
            )

        for record in payload["diesel_records"]:
            session.add(
                DieselRecord(
                    work_date=str(record.get("work_date") or ""),
                    equipment=str(record.get("equipment") or ""),
                    condition=str(record.get("condition") or "DISPONIBLE"),
                    shift=str(record.get("shift") or "1"),
                    horometer_initial=number(record.get("horometer_initial")),
                    horometer_final=number(record.get("horometer_final")),
                    worked_hours=number(record.get("worked_hours")),
                    diesel_liters=number(record.get("diesel_liters")),
                    operator=str(record.get("operator") or ""),
                    dispatcher=str(record.get("dispatcher") or ""),
                    supervisor=str(record.get("supervisor") or ""),
                    notes=str(record.get("notes") or ""),
                    source="desktop-backup",
                    created_at=parse_datetime(record.get("created_at")),
                    updated_at=parse_datetime(record.get("updated_at")),
                )
            )
        for day in payload["diesel_days"]:
            session.add(
                DieselDay(
                    work_date=str(day.get("work_date") or ""),
                    diesel_received=number(day.get("diesel_received")),
                    initial_stock=number(day.get("initial_stock")),
                    final_stock=number(day.get("final_stock")),
                    prosermin_stock=number(day.get("prosermin_stock")),
                    supplier=str(day.get("supplier") or ""),
                    notes=str(day.get("notes") or ""),
                    source="desktop-backup",
                    created_at=parse_datetime(day.get("created_at")),
                    updated_at=parse_datetime(day.get("updated_at")),
                )
            )
        session.flush()

        settings = payload["settings"]
        year = integer(config_value(settings, "ANIO", now.year), now.year)
        month = integer(config_value(settings, "MES_NUMERO", now.month), now.month)
        period_start = f"{year:04d}-{month:02d}-01"
        capture_dates = [str(row.get("work_date") or "") for row in payload["captures"] if row.get("work_date")]
        period_end = max(capture_dates) if capture_dates else now.date().isoformat()
        diesel = diesel_payload(session, min(capture_dates) if capture_dates else period_start, period_end)
        tire_kpi = {
            "rows": payload["tires"],
            "events": payload["tire_events"],
            "summary": {
                "tires": len(payload["tires"]),
                "events": len(payload["tire_events"]),
            },
        }
        portal = {
            "ok": True,
            "source": "cloud-portal",
            "generated_at": now.isoformat(timespec="seconds"),
            "period": {"start": period_start, "end": period_end, "year": year, "month": month},
            "settings": {
                "shift_hours": number(config_value(settings, "HORAS_TURNO", 9), 9),
                "turns_per_day": number(config_value(settings, "TURNOS_DIA", 2), 2),
                "meta_availability": number(config_value(settings, "META_DISPONIBILIDAD", 85), 85),
                "meta_utilization": number(config_value(settings, "META_UTILIZACION", 75), 75),
                "meta_tmef": number(config_value(settings, "META_TMEF", 8), 8),
                "meta_tmpr": number(config_value(settings, "META_TMPR", 4), 4),
                "meta_diesel_lh": number(config_value(settings, "META_DIESEL_LH", 25), 25),
            },
            "equipment": payload["equipment"],
            "captures": payload["captures"],
            "preventives": payload["preventives"],
            "service_history": payload["preventives"],
            "availability": payload["availability"],
            "kpi_groups": ["Todos los equipos", "Equipos de Barrenacion", "Equipos de Rezagado", "KPI Aceites", "KPI Llantas"],
            "kpi_reports": {},
            "oil_kpi": {"rows": [], "totals": {}, "columns": []},
            "tire_kpi": tire_kpi,
            "diesel": diesel,
        }
        session.add(
            CatalogSnapshot(name="default", updated_at=now, payload_json=json_dumps({"ok": True, "source": "cloud", "equipment": payload["equipment"]}))
        )
        session.add(PortalSnapshot(name="default", updated_at=now, payload_json=json_dumps(portal)))
        session.commit()

    print(json.dumps({"ok": True, "restored": counts}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
