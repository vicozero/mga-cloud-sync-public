from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy import func
from sqlalchemy.orm import Session

from app.api.deps import get_current_user, require_roles
from app.db.session import get_db
from app.models.equipment import Equipment
from app.models.project import Project
from app.models.user import User
from app.schemas.equipment import EquipmentCreate, EquipmentOrder, EquipmentResponse, EquipmentUpdate
from app.services.text_clean import caps, clean_category, clean_eco, clean_name, clean_status, clean_text

router = APIRouter(prefix="/equipment", tags=["equipment"])


@router.get("", response_model=list[EquipmentResponse])
def list_equipment(
    project_id: int | None = Query(default=None),
    category: str | None = Query(default=None),
    status_filter: str | None = Query(default=None, alias="status"),
    active: bool | None = Query(default=None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    query = db.query(Equipment)
    if project_id is not None:
        query = query.filter(Equipment.project_id == project_id)
    if category:
        query = query.filter(Equipment.category == category)
    if status_filter:
        query = query.filter(Equipment.current_status == status_filter)
    if active is not None:
        query = query.filter(Equipment.active == active)
    return query.order_by(Equipment.sort_order.asc(), Equipment.category.asc(), Equipment.eco.asc()).all()


@router.post("", response_model=EquipmentResponse, status_code=status.HTTP_201_CREATED)
def create_equipment(
    payload: EquipmentCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    project = db.query(Project).filter(Project.id == payload.project_id).first()
    if not project:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado")

    exists = db.query(Equipment).filter(
        Equipment.project_id == payload.project_id,
        Equipment.eco == clean_eco(payload.eco),
    ).first()
    if exists:
        raise HTTPException(status_code=409, detail="Ya existe un equipo con ese ECO en el proyecto")

    last_sort = db.query(func.max(Equipment.sort_order)).filter(Equipment.project_id == payload.project_id).scalar() or 0

    equipment = Equipment(
        project_id=payload.project_id,
        category=clean_category(payload.category),
        name=clean_name(payload.name),
        eco=clean_eco(payload.eco),
        current_status=clean_status(payload.current_status),
        notes=caps(payload.notes),
        hour_meter=clean_text(payload.hour_meter),
        active=payload.active,
        sort_order=payload.sort_order if payload.sort_order else last_sort + 1,
        created_by=current_user.id,
        updated_by=current_user.id,
    )
    db.add(equipment)
    db.commit()
    db.refresh(equipment)
    return equipment


@router.put("/order", response_model=list[EquipmentResponse])
def reorder_equipment(
    payload: EquipmentOrder,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    project = db.query(Project).filter(Project.id == payload.project_id).first()
    if not project:
        raise HTTPException(status_code=404, detail="Proyecto no encontrado")

    equipments = db.query(Equipment).filter(Equipment.project_id == payload.project_id).all()
    by_id = {equipment.id: equipment for equipment in equipments}
    if len(payload.ids) != len(set(payload.ids)):
        raise HTTPException(status_code=400, detail="La lista contiene IDs duplicados")
    for equipment_id in payload.ids:
        if equipment_id not in by_id:
            raise HTTPException(status_code=400, detail=f"El equipo {equipment_id} no pertenece al proyecto")

    active = [equipment for equipment in equipments if equipment.active]
    if len(payload.ids) != len(active):
        raise HTTPException(
            status_code=400,
            detail="La lista debe incluir exactamente los equipos activos del proyecto",
        )

    for index, equipment_id in enumerate(payload.ids):
        item = by_id[equipment_id]
        item.sort_order = index + 1
        item.updated_by = current_user.id
    db.commit()

    return (
        db.query(Equipment)
        .filter(Equipment.project_id == payload.project_id)
        .order_by(Equipment.sort_order.asc(), Equipment.category.asc(), Equipment.eco.asc())
        .all()
    )


@router.get("/{equipment_id}", response_model=EquipmentResponse)
def get_equipment(
    equipment_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    equipment = db.query(Equipment).filter(Equipment.id == equipment_id).first()
    if not equipment:
        raise HTTPException(status_code=404, detail="Equipo no encontrado")
    return equipment


@router.put("/{equipment_id}", response_model=EquipmentResponse)
def update_equipment(
    equipment_id: int,
    payload: EquipmentUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    equipment = db.query(Equipment).filter(Equipment.id == equipment_id).first()
    if not equipment:
        raise HTTPException(status_code=404, detail="Equipo no encontrado")

    data = payload.model_dump(exclude_unset=True)
    if "category" in data and data["category"] is not None:
        data["category"] = clean_category(data["category"])
    if "eco" in data and data["eco"] is not None:
        data["eco"] = clean_eco(data["eco"])
    if "current_status" in data and data["current_status"] is not None:
        data["current_status"] = clean_status(data["current_status"])
    if "name" in data and data["name"] is not None:
        data["name"] = clean_name(data["name"])
    if "hour_meter" in data and data["hour_meter"] is not None:
        data["hour_meter"] = clean_text(data["hour_meter"])
    if "notes" in data and data["notes"] is not None:
        data["notes"] = caps(data["notes"])

    for key, value in data.items():
        setattr(equipment, key, value)
    equipment.updated_by = current_user.id

    db.commit()
    db.refresh(equipment)
    return equipment


@router.delete("/{equipment_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_equipment(
    equipment_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin")
    equipment = db.query(Equipment).filter(Equipment.id == equipment_id).first()
    if not equipment:
        raise HTTPException(status_code=404, detail="Equipo no encontrado")
    equipment.active = False
    equipment.updated_by = current_user.id
    db.commit()
    return None
