from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import func
from sqlalchemy.orm import Session

from app.api.deps import get_current_user, require_roles
from app.db.session import get_db
from app.models.consumption import Consumption
from app.models.mechanic_report import MechanicReport
from app.models.project import Project
from app.models.supervisor import Supervisor
from app.models.tire_inspection import TireInspection
from app.models.user import User
from app.schemas.supervisor import SupervisorCreate, SupervisorResponse, SupervisorUpdate
from app.services.text_clean import clean_supervisor

router = APIRouter(prefix="/supervisors", tags=["supervisors"])


def _normalize_name(name: str) -> str:
    return clean_supervisor(name) or ""


def _record_counts(db: Session, supervisor: Supervisor) -> int:
    name = supervisor.name
    return (
        db.query(func.count(Consumption.id))
        .filter(func.upper(Consumption.supervisor_name) == name)
        .scalar()
        or 0
    ) + (
        db.query(func.count(MechanicReport.id))
        .filter(func.upper(MechanicReport.supervisor_name) == name)
        .scalar()
        or 0
    ) + (
        db.query(func.count(TireInspection.id))
        .filter(func.upper(TireInspection.supervisor_name) == name)
        .scalar()
        or 0
    )


def _response(db: Session, item: Supervisor) -> dict:
    return {
        "id": item.id,
        "project_id": item.project_id,
        "name": item.name,
        "active": item.active,
        "sort_order": item.sort_order,
        "record_count": _record_counts(db, item),
        "created_at": item.created_at,
        "updated_at": item.updated_at,
    }


def _find(db: Session, supervisor_id: int) -> Supervisor:
    item = db.query(Supervisor).filter(Supervisor.id == supervisor_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Supervisor no encontrado")
    return item


@router.get("", response_model=list[SupervisorResponse])
def list_supervisors(
    project_id: int | None = None,
    include_inactive: bool = False,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    q = db.query(Supervisor)
    if project_id is not None:
        q = q.filter(Supervisor.project_id == project_id)
    if not include_inactive:
        q = q.filter(Supervisor.active.is_(True))
    rows = q.order_by(Supervisor.sort_order.asc(), func.upper(Supervisor.name).asc()).all()
    return [_response(db, row) for row in rows]


@router.post("", response_model=SupervisorResponse, status_code=status.HTTP_201_CREATED)
def create_supervisor(
    payload: SupervisorCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    project = db.query(Project).filter(Project.id == payload.project_id).first()
    if not project:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado")
    name = _normalize_name(payload.name)
    if not name:
        raise HTTPException(status_code=400, detail="El nombre no puede ir vacío")
    exists = (
        db.query(Supervisor)
        .filter(
            Supervisor.project_id == payload.project_id,
            func.upper(Supervisor.name) == name,
        )
        .first()
    )
    if exists:
        raise HTTPException(status_code=409, detail="Ya existe ese supervisor en el proyecto")

    last_sort = (
        db.query(func.max(Supervisor.sort_order))
        .filter(Supervisor.project_id == payload.project_id)
        .scalar()
        or 0
    )
    item = Supervisor(
        project_id=payload.project_id,
        name=name,
        sort_order=last_sort + 1,
        created_by=current_user.id,
        updated_by=current_user.id,
    )
    db.add(item)
    db.commit()
    db.refresh(item)
    return _response(db, item)


@router.put("/{supervisor_id}", response_model=SupervisorResponse)
def update_supervisor(
    supervisor_id: int,
    payload: SupervisorUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    item = _find(db, supervisor_id)
    data = payload.model_dump(exclude_unset=True)

    if "name" in data and data["name"] is not None:
        new_name = _normalize_name(data["name"])
        if not new_name:
            raise HTTPException(status_code=400, detail="El nombre no puede ir vacío")
        if new_name != item.name:
            conflict = (
                db.query(Supervisor)
                .filter(
                    Supervisor.id != supervisor_id,
                    Supervisor.project_id == item.project_id,
                    func.upper(Supervisor.name) == new_name,
                )
                .first()
            )
            if conflict:
                raise HTTPException(status_code=409, detail="Ya existe otro supervisor con ese nombre en el proyecto")
            old_name = func.upper(item.name)
            db.query(Consumption).filter(func.upper(Consumption.supervisor_name) == old_name).update(
                {"supervisor_name": new_name, "updated_by": current_user.id}, synchronize_session=False
            )
            db.query(MechanicReport).filter(func.upper(MechanicReport.supervisor_name) == old_name).update(
                {"supervisor_name": new_name, "updated_by": current_user.id}, synchronize_session=False
            )
            db.query(TireInspection).filter(func.upper(TireInspection.supervisor_name) == old_name).update(
                {"supervisor_name": new_name, "updated_by": current_user.id}, synchronize_session=False
            )
            item.name = new_name
    if "active" in data and data["active"] is not None:
        item.active = data["active"]
    if "sort_order" in data and data["sort_order"] is not None:
        item.sort_order = data["sort_order"]
    item.updated_by = current_user.id
    db.commit()
    db.refresh(item)
    return _response(db, item)


@router.delete("/{supervisor_id}", response_model=SupervisorResponse)
def delete_supervisor(
    supervisor_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin")
    item = _find(db, supervisor_id)
    used = _record_counts(db, item)
    if used:
        raise HTTPException(
            status_code=400,
            detail="El supervisor tiene registros asignados. Desactívalo o reasigna el nombre para continuar.",
        )
    db.delete(item)
    db.commit()
    return _response(db, item)