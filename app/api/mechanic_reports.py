from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session

from app.api.deps import get_current_user, require_roles
from app.db.session import get_db
from app.models.equipment import Equipment
from app.models.mechanic_report import MechanicReport
from app.models.project import Project
from app.models.user import User
from app.schemas.mechanic_report import (
    MechanicReportCreate,
    MechanicReportResponse,
    MechanicReportUpdate,
)
from app.services.text_clean import caps, clean_name, clean_status, clean_supervisor, clean_text

router = APIRouter(prefix="/mechanic-reports", tags=["mechanic-reports"])


def sync_equipment_status(equipment: Equipment, final_status: str) -> None:
    normalized = final_status.strip().lower()
    if normalized == "disponible":
        equipment.current_status = "DISPONIBLE"
    elif normalized == "operativa":
        equipment.current_status = "OPERATIVA"
    elif normalized == "fuera de servicio":
        equipment.current_status = "FUERA DE SERVICIO"


@router.get("", response_model=list[MechanicReportResponse])
def list_mechanic_reports(
    project_id: int | None = Query(default=None),
    equipment_id: int | None = Query(default=None),
    shift: int | None = Query(default=None),
    final_status: str | None = Query(default=None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    query = db.query(MechanicReport)
    if project_id is not None:
        query = query.filter(MechanicReport.project_id == project_id)
    if equipment_id is not None:
        query = query.filter(MechanicReport.equipment_id == equipment_id)
    if shift is not None:
        query = query.filter(MechanicReport.shift == shift)
    if final_status:
        query = query.filter(MechanicReport.final_status == final_status)
    return query.order_by(MechanicReport.work_date.desc(), MechanicReport.id.desc()).all()


@router.post("", response_model=MechanicReportResponse, status_code=status.HTTP_201_CREATED)
def create_mechanic_report(
    payload: MechanicReportCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    project = db.query(Project).filter(Project.id == payload.project_id).first()
    equipment = db.query(Equipment).filter(Equipment.id == payload.equipment_id).first()
    if not project:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado")
    if not equipment:
        raise HTTPException(status_code=404, detail="Equipo no encontrado")

    item = MechanicReport(
        project_id=payload.project_id,
        equipment_id=payload.equipment_id,
        work_date=payload.work_date,
        shift=payload.shift,
        mechanic_name=clean_name(payload.mechanic_name),
        supervisor_name=clean_supervisor(payload.supervisor_name),
        hour_meter=payload.hour_meter,
        service_type=clean_text(payload.service_type),
        failure=caps(payload.failure),
        work_done=caps(payload.work_done),
        parts_used=caps(payload.parts_used),
        final_status=clean_status(payload.final_status),
        evidence_url=payload.evidence_url,
        created_by=current_user.id,
        updated_by=current_user.id,
    )
    sync_equipment_status(equipment, item.final_status)
    equipment.updated_by = current_user.id

    db.add(item)
    db.commit()
    db.refresh(item)
    return item


@router.get("/{item_id}", response_model=MechanicReportResponse)
def get_mechanic_report(
    item_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    item = db.query(MechanicReport).filter(MechanicReport.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Reporte mecánico no encontrado")
    return item


@router.put("/{item_id}", response_model=MechanicReportResponse)
def update_mechanic_report(
    item_id: int,
    payload: MechanicReportUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    item = db.query(MechanicReport).filter(MechanicReport.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Reporte mecánico no encontrado")

    data = payload.model_dump(exclude_unset=True)
    for field in ["service_type"]:
        if field in data and data[field] is not None:
            data[field] = clean_text(data[field])
    for field in ["failure", "work_done", "parts_used"]:
        if field in data and data[field] is not None:
            data[field] = caps(data[field])
    if "mechanic_name" in data and data["mechanic_name"] is not None:
        data["mechanic_name"] = clean_name(data["mechanic_name"])
    if "final_status" in data and data["final_status"] is not None:
        data["final_status"] = clean_status(data["final_status"])
    if "supervisor_name" in data and data["supervisor_name"] is not None:
        data["supervisor_name"] = clean_supervisor(data["supervisor_name"])

    for key, value in data.items():
        setattr(item, key, value)
    item.updated_by = current_user.id

    equipment = db.query(Equipment).filter(Equipment.id == item.equipment_id).first()
    if equipment:
        sync_equipment_status(equipment, item.final_status)
        equipment.updated_by = current_user.id

    db.commit()
    db.refresh(item)
    return item


@router.delete("/{item_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_mechanic_report(
    item_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin")
    item = db.query(MechanicReport).filter(MechanicReport.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Reporte mecánico no encontrado")
    db.delete(item)
    db.commit()
    return None
