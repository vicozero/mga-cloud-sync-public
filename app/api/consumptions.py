from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session

from app.api.deps import get_current_user, require_roles
from app.db.session import get_db
from app.models.consumption import Consumption
from app.models.equipment import Equipment
from app.models.project import Project
from app.models.user import User
from app.schemas.consumption import ConsumptionCreate, ConsumptionResponse, ConsumptionUpdate
from app.services.text_clean import caps, clean_supervisor, clean_text, clean_unit

router = APIRouter(prefix="/consumptions", tags=["consumptions"])


@router.get("", response_model=list[ConsumptionResponse])
def list_consumptions(
    project_id: int | None = Query(default=None),
    equipment_id: int | None = Query(default=None),
    shift: int | None = Query(default=None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    query = db.query(Consumption)
    if project_id is not None:
        query = query.filter(Consumption.project_id == project_id)
    if equipment_id is not None:
        query = query.filter(Consumption.equipment_id == equipment_id)
    if shift is not None:
        query = query.filter(Consumption.shift == shift)
    return query.order_by(Consumption.work_date.desc(), Consumption.id.desc()).all()


@router.post("", response_model=ConsumptionResponse, status_code=status.HTTP_201_CREATED)
def create_consumption(
    payload: ConsumptionCreate,
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

    item = Consumption(
        project_id=payload.project_id,
        equipment_id=payload.equipment_id,
        work_date=payload.work_date,
        shift=payload.shift,
        lubricant=caps(payload.lubricant),
        quantity=payload.quantity,
        unit=clean_unit(payload.unit),
        notes=caps(payload.notes),
        supervisor_name=clean_supervisor(payload.supervisor_name),
        created_by=current_user.id,
        updated_by=current_user.id,
    )
    db.add(item)
    db.commit()
    db.refresh(item)
    return item


@router.get("/{item_id}", response_model=ConsumptionResponse)
def get_consumption(
    item_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    item = db.query(Consumption).filter(Consumption.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Consumo no encontrado")
    return item


@router.put("/{item_id}", response_model=ConsumptionResponse)
def update_consumption(
    item_id: int,
    payload: ConsumptionUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    item = db.query(Consumption).filter(Consumption.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Consumo no encontrado")

    data = payload.model_dump(exclude_unset=True)
    if "lubricant" in data and data["lubricant"] is not None:
        data["lubricant"] = caps(data["lubricant"])
    if "unit" in data and data["unit"] is not None:
        data["unit"] = clean_unit(data["unit"])
    if "notes" in data and data["notes"] is not None:
        data["notes"] = caps(data["notes"])
    if "supervisor_name" in data and data["supervisor_name"] is not None:
        data["supervisor_name"] = clean_supervisor(data["supervisor_name"])
    for key, value in data.items():
        setattr(item, key, value)
    item.updated_by = current_user.id

    db.commit()
    db.refresh(item)
    return item


@router.delete("/{item_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_consumption(
    item_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin")
    item = db.query(Consumption).filter(Consumption.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Consumo no encontrado")
    db.delete(item)
    db.commit()
    return None
