from datetime import datetime
from pathlib import Path
from typing import List

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import Response
from sqlalchemy.orm import Session

from app.api.deps import get_current_user
from app.db.session import get_db
from app.models.consumption import Consumption
from app.models.report_file import ReportFile
from app.services import file_store
from app.models.equipment import Equipment
from app.models.mechanic_report import MechanicReport
from app.models.project import Project
from app.models.supervisor import Supervisor
from app.models.tire_inspection import TireInspection
from app.models.tire_detail import TireDetail
from app.models.user import User
from app.services.excel import build_simple_workbook
from app.services.pdf import (
    build_consumptions_pdf,
    build_equipment_pdf,
    build_mechanic_pdf,
    build_tires_pdf,
)

router = APIRouter(prefix="/reports", tags=["reports"])

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF_MEDIA = "application/pdf"

REPORT_DIR = Path(__file__).resolve().parent.parent.parent / "reports"


def _parse_date(value: str):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(status_code=400, detail="Fecha no válida, usa AAAA-MM-DD")


def _filter_shift(shift: int):
    if shift is None:
        return None
    if shift not in (1, 2):
        raise HTTPException(status_code=400, detail="Turno debe ser 1 o 2")
    return shift


def _pdf(content: bytes, filename: str) -> Response:
    return Response(content=content, media_type=PDF_MEDIA, headers={"Content-Disposition": f'attachment; filename="{filename}"', "Cache-Control": "no-store"})


def _report_modules(module: str):
    pdf_map = {
        "equipment-status": ("equipment_status", _pdf_equipment_rows, build_equipment_pdf),
        "consumptions": ("consumptions", _pdf_consumption_rows, build_consumptions_pdf),
        "mechanic-reports": ("mechanic_reports", _pdf_mechanic_report_rows, build_mechanic_pdf),
        "tire-inspections": ("tire_inspections", _pdf_tire_inspection_rows, build_tires_pdf),
    }
    excel_map = {
        "equipment-status": ("equipment_status", "Equipos", _equipment_rows),
        "consumptions": ("consumptions", "Consumos de lubricantes", _consumption_rows),
        "mechanic-reports": ("mechanic_reports", "Reportes mecánicos", _mechanic_report_rows),
        "tire-inspections": ("tire_inspections", "Inspecciones de llantas", _tire_inspection_rows),
    }
    return pdf_map.get(module), excel_map.get(module)


@router.get("/repositorio")
def list_report_files(db: Session = Depends(get_db)):
    if file_store.is_cloud():
        rows = db.query(ReportFile).order_by(ReportFile.created_at.desc(), ReportFile.id.desc()).all()
        return [
            {"name": r.name, "size": r.size, "modified": r.created_at.strftime("%d/%m/%Y %H:%M")}
            for r in rows
        ]
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    entries = []
    for f in sorted(REPORT_DIR.glob("*"), reverse=True):
        if f.is_file():
            entries.append(
                {
                    "name": f.name,
                    "size": f.stat().st_size,
                    "modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
                }
            )
    return entries


@router.post("/repositorio/generate")
def generate_report_file(module: str, fmt: str, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    pdf_mod, excel_mod = _report_modules(module)
    if fmt not in ("pdf", "xlsx"):
        raise HTTPException(status_code=404, detail="Formato no válido")
    if fmt == "pdf" and pdf_mod is None:
        raise HTTPException(status_code=404, detail="Módulo no encontrado")
    if fmt == "xlsx" and excel_mod is None:
        raise HTTPException(status_code=404, detail="Módulo no encontrado")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if fmt == "pdf":
        basename, rows_builder, builder = pdf_mod
        project, rows = rows_builder(db)
        content = builder(project, rows)
        media = PDF_MEDIA
    else:
        basename, _title, rows_builder = excel_mod
        title, headers, rows = rows_builder(db)
        content = build_simple_workbook(title, headers, rows)
        media = XLSX_MEDIA

    name = f"{basename}_{stamp}.{fmt}"
    if file_store.is_cloud():
        ok = file_store.upload_bytes(f"reports/{name}", content, media)
        if not ok:
            raise HTTPException(status_code=500, detail="Error al guardar el reporte en la nube")
        db.add(ReportFile(name=name, size=len(content)))
        db.commit()
    else:
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        (REPORT_DIR / name).write_bytes(content)
    return {"name": name, "size": len(content)}


@router.delete("/repositorio/{name}")
def delete_report_file(name: str, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    safe = Path(name).name
    if file_store.is_cloud():
        file_store.delete_object(f"reports/{safe}")
        db.query(ReportFile).filter(ReportFile.name == safe).delete()
        db.commit()
        return {"deleted": safe}
    path = (REPORT_DIR / safe).resolve()
    if not str(path).startswith(str(REPORT_DIR.resolve())) or not path.is_file():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    path.unlink()
    return {"deleted": safe}


def _equipment_rows(db: Session) -> tuple[str, List[list]]:
    rows = [
        [item.id, item.project_id, item.category, item.name, item.eco, item.current_status, item.notes or ""]
        for item in db.query(Equipment).order_by(Equipment.project_id.asc(), Equipment.sort_order.asc(), Equipment.category.asc(), Equipment.eco.asc()).all()
    ]
    return (
        "Equipos",
        ["ID", "Proyecto", "Categoría", "Equipo", "ECO", "Estado", "Observaciones"],
        rows,
    )


def _consumption_rows(db, shift=None, supervisor=None, date_from=None, date_to=None):
    q = db.query(Consumption)
    if shift:
        q = q.filter(Consumption.shift == shift)
    if supervisor:
        q = q.filter(Consumption.supervisor_name.ilike(f"%{supervisor.strip()}%"))
    if date_from:
        q = q.filter(Consumption.work_date >= date_from)
    if date_to:
        q = q.filter(Consumption.work_date <= date_to)
    rows = [
        [item.id, item.project_id, item.equipment_id or "", str(item.work_date), item.shift, item.lubricant, float(item.quantity), item.unit, item.supervisor_name or "", item.notes or ""]
        for item in q.order_by(Consumption.work_date.desc(), Consumption.id.desc()).all()
    ]
    return (
        "Consumos de lubricantes",
        ["ID", "Proyecto", "Equipo", "Fecha", "Turno", "Lubricante", "Cantidad", "Unidad", "Supervisor", "Observaciones"],
        rows,
    )


def _mechanic_report_rows(db, shift=None, supervisor=None, date_from=None, date_to=None):
    q = db.query(MechanicReport)
    if shift:
        q = q.filter(MechanicReport.shift == shift)
    if supervisor:
        q = q.filter(MechanicReport.supervisor_name.ilike(f"%{supervisor.strip()}%"))
    if date_from:
        q = q.filter(MechanicReport.work_date >= date_from)
    if date_to:
        q = q.filter(MechanicReport.work_date <= date_to)
    rows = [
        [item.id, item.project_id, item.equipment_id or "", str(item.work_date), item.shift, item.mechanic_name, item.supervisor_name or "", item.hour_meter or "", item.service_type, item.failure, item.work_done, item.parts_used or "", item.final_status]
        for item in q.order_by(MechanicReport.work_date.desc(), MechanicReport.id.desc()).all()
    ]
    return (
        "Reportes mecánicos",
        ["ID", "Proyecto", "Equipo", "Fecha", "Turno", "Mecánico", "Supervisor", "Horómetro", "Servicio", "Falla", "Trabajo", "Refacciones", "Estado final"],
        rows,
    )


def _tire_inspection_rows(db, supervisor=None, date_from=None, date_to=None):
    q = db.query(TireInspection)
    if supervisor:
        q = q.filter(TireInspection.supervisor_name.ilike(f"%{supervisor.strip()}%"))
    if date_from:
        q = q.filter(TireInspection.work_date >= date_from)
    if date_to:
        q = q.filter(TireInspection.work_date <= date_to)
    rows = [
        [item.id, item.project_id, item.equipment_id or "", str(item.work_date), item.hour_meter or "", item.location or "", item.inspector_name or "", item.supervisor_name or "", item.reviewed_by or "", item.approved_by or "", item.general_notes or ""]
        for item in q.order_by(TireInspection.work_date.desc(), TireInspection.id.desc()).all()
    ]
    return (
        "Inspecciones de llantas",
        ["ID", "Proyecto", "Equipo", "Fecha", "Horómetro", "Ubicación", "Inspector", "Supervisor", "Revisó", "Aprobó", "Observaciones"],
        rows,
    )


def _project_label(db: Session) -> str:
    names = [p.name for p in db.query(Project.name).filter(Project.active.is_(True)).all()]
    if len(names) == 1:
        return names[0]
    return "TODOS LOS PROYECTOS"


def _equipment_map(db: Session) -> dict:
    return {e.id: (e.name, e.eco) for e in db.query(Equipment.id, Equipment.name, Equipment.eco).all()}


def _pdf_equipment_rows(db: Session) -> tuple[str, List[list]]:
    rows = [
        [item.category, item.name, item.eco, item.current_status, item.notes or "", item.hour_meter or ""]
        for item in db.query(Equipment)
        .order_by(Equipment.project_id.asc(), Equipment.sort_order.asc(), Equipment.category.asc(), Equipment.eco.asc())
        .all()
    ]
    return _project_label(db), rows


def _pdf_consumption_rows(db, shift=None, supervisor=None, date_from=None, date_to=None):
    emap = _equipment_map(db)
    q = db.query(Consumption)
    if shift:
        q = q.filter(Consumption.shift == shift)
    if supervisor:
        q = q.filter(Consumption.supervisor_name.ilike(f"%{supervisor.strip()}%"))
    if date_from:
        q = q.filter(Consumption.work_date >= date_from)
    if date_to:
        q = q.filter(Consumption.work_date <= date_to)
    rows = [
        [str(item.work_date), emap.get(item.equipment_id or -1, ("", ""))[1], emap.get(item.equipment_id or -1, ("", ""))[0], item.lubricant, float(item.quantity), item.unit, item.notes or "", item.supervisor_name or "", item.shift]
        for item in q.order_by(Consumption.work_date.desc(), Consumption.id.desc()).all()
    ]
    return _project_label(db), rows


def _pdf_mechanic_report_rows(db, shift=None, supervisor=None, date_from=None, date_to=None):
    emap = _equipment_map(db)
    q = db.query(MechanicReport)
    if shift:
        q = q.filter(MechanicReport.shift == shift)
    if supervisor:
        q = q.filter(MechanicReport.supervisor_name.ilike(f"%{supervisor.strip()}%"))
    if date_from:
        q = q.filter(MechanicReport.work_date >= date_from)
    if date_to:
        q = q.filter(MechanicReport.work_date <= date_to)
    rows = [
        [emap.get(item.equipment_id or -1, ("", ""))[1], emap.get(item.equipment_id or -1, ("", ""))[0], str(item.work_date), item.mechanic_name, item.hour_meter or "", item.service_type, item.failure, item.work_done, item.parts_used or "", item.final_status, item.evidence_url or "", item.supervisor_name or ""]
        for item in q.order_by(MechanicReport.work_date.desc(), MechanicReport.id.desc()).all()
    ]
    return _project_label(db), rows


def _pdf_tire_inspection_rows(db, supervisor=None, date_from=None, date_to=None):
    emap = _equipment_map(db)
    q = db.query(TireInspection)
    if supervisor:
        q = q.filter(TireInspection.supervisor_name.ilike(f"%{supervisor.strip()}%"))
    if date_from:
        q = q.filter(TireInspection.work_date >= date_from)
    if date_to:
        q = q.filter(TireInspection.work_date <= date_to)
    rows = []
    inspections = q.order_by(TireInspection.work_date.desc(), TireInspection.id.desc()).all()
    for item in inspections:
        details = [
            [d.position, d.status, d.vida_util or "", d.notes or ""]
            for d in db.query(TireDetail)
            .filter(TireDetail.inspection_id == item.id)
            .order_by(TireDetail.id.asc())
            .all()
        ]
        rows.append(
            [emap.get(item.equipment_id or -1, ("", ""))[1], emap.get(item.equipment_id or -1, ("", ""))[0], str(item.work_date), item.hour_meter or "", item.location or "", item.inspector_name or "", item.general_notes or "", item.supervisor_name or "", details]
        )
    return _project_label(db), rows


def _xl(module: str, filename: str) -> Response:
    return Response(content=module, media_type=XLSX_MEDIA, headers={"Content-Disposition": f'attachment; filename="{filename}"', "Cache-Control": "no-store"})


def _pdf(content: bytes, filename: str) -> Response:
    return Response(content=content, media_type=PDF_MEDIA, headers={"Content-Disposition": f'attachment; filename="{filename}"', "Cache-Control": "no-store"})


@router.get("/supervisors")
def list_supervisors(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    catalog = {name for (name,) in db.query(Supervisor.name).filter(Supervisor.active.is_(True)).all()}
    used = set()
    for (name,) in db.query(Consumption.supervisor_name).all():
        if name:
            used.add(name)
    for (name,) in db.query(MechanicReport.supervisor_name).all():
        if name:
            used.add(name)
    for (name,) in db.query(TireInspection.supervisor_name).all():
        if name:
            used.add(name)
    orphans = sorted(name for name in used if name not in catalog)
    return sorted(catalog) + orphans


@router.get("/excel/equipment-status")
def export_equipment_status_excel(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    title, headers, rows = _equipment_rows(db)
    return _xl(build_simple_workbook(title, headers, rows), "equipment_status.xlsx")


@router.get("/excel/consumptions")
def export_consumptions_excel(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    shift: int = Query(None, ge=1, le=2),
    supervisor: str = None,
    date_from: str = None,
    date_to: str = None,
):
    title, headers, rows = _consumption_rows(db, _filter_shift(shift), supervisor, _parse_date(date_from), _parse_date(date_to))
    return _xl(build_simple_workbook(title, headers, rows), "consumptions.xlsx")


@router.get("/excel/mechanic-reports")
def export_mechanic_reports_excel(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    shift: int = Query(None, ge=1, le=2),
    supervisor: str = None,
    date_from: str = None,
    date_to: str = None,
):
    title, headers, rows = _mechanic_report_rows(db, _filter_shift(shift), supervisor, _parse_date(date_from), _parse_date(date_to))
    return _xl(build_simple_workbook(title, headers, rows), "mechanic_reports.xlsx")


@router.get("/excel/tire-inspections")
def export_tire_inspections_excel(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    supervisor: str = None,
    date_from: str = None,
    date_to: str = None,
):
    title, headers, rows = _tire_inspection_rows(db, supervisor, _parse_date(date_from), _parse_date(date_to))
    return _xl(build_simple_workbook(title, headers, rows), "tire_inspections.xlsx")


@router.get("/pdf/equipment-status")
def export_equipment_status_pdf(db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    project, rows = _pdf_equipment_rows(db)
    return _pdf(build_equipment_pdf(project, rows), "equipment_status.pdf")


@router.get("/pdf/consumptions")
def export_consumptions_pdf(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    shift: int = Query(None, ge=1, le=2),
    supervisor: str = None,
    date_from: str = None,
    date_to: str = None,
):
    project, rows = _pdf_consumption_rows(db, _filter_shift(shift), supervisor, _parse_date(date_from), _parse_date(date_to))
    return _pdf(build_consumptions_pdf(project, rows), "consumptions.pdf")


@router.get("/pdf/mechanic-reports")
def export_mechanic_reports_pdf(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    shift: int = Query(None, ge=1, le=2),
    supervisor: str = None,
    date_from: str = None,
    date_to: str = None,
):
    project, rows = _pdf_mechanic_report_rows(db, _filter_shift(shift), supervisor, _parse_date(date_from), _parse_date(date_to))
    return _pdf(build_mechanic_pdf(project, rows), "mechanic_reports.pdf")


@router.get("/pdf/tire-inspections")
def export_tire_inspections_pdf(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
    supervisor: str = None,
    date_from: str = None,
    date_to: str = None,
):
    project, rows = _pdf_tire_inspection_rows(db, supervisor, _parse_date(date_from), _parse_date(date_to))
    return _pdf(build_tires_pdf(project, rows), "tire_inspections.pdf")