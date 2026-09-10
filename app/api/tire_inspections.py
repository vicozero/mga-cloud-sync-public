from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session, selectinload

from app.api.deps import get_current_user, require_roles
from app.db.session import get_db
from app.models.equipment import Equipment
from app.models.project import Project
from app.models.tire_detail import TireDetail
from app.models.tire_inspection import TireInspection
from app.models.user import User
from app.schemas.tire_inspection import (
    TireDetailResponse,
    TireInspectionCreate,
    TireInspectionResponse,
    TireInspectionUpdate,
)
from app.services.text_clean import caps, clean_name, clean_supervisor, clean_text

router = APIRouter(prefix="/tire-inspections", tags=["tire-inspections"])


def to_response(item: TireInspection) -> TireInspectionResponse:
    return TireInspectionResponse(
        id=item.id,
        project_id=item.project_id,
        equipment_id=item.equipment_id,
        work_date=item.work_date,
        hour_meter=item.hour_meter,
        location=item.location,
        inspector_name=item.inspector_name,
        supervisor_name=item.supervisor_name,
        reviewed_by=item.reviewed_by,
        approved_by=item.approved_by,
        general_notes=item.general_notes,
        details=[TireDetailResponse.model_validate(detail) for detail in item.details],
        created_by=item.created_by,
        updated_by=item.updated_by,
        created_at=item.created_at,
        updated_at=item.updated_at,
    )


@router.get("", response_model=list[TireInspectionResponse])
def list_tire_inspections(
    project_id: int | None = Query(default=None),
    equipment_id: int | None = Query(default=None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    query = db.query(TireInspection).options(selectinload(TireInspection.details))
    if project_id is not None:
        query = query.filter(TireInspection.project_id == project_id)
    if equipment_id is not None:
        query = query.filter(TireInspection.equipment_id == equipment_id)
    items = query.order_by(TireInspection.work_date.desc(), TireInspection.id.desc()).all()
    return [to_response(item) for item in items]


@router.post("", response_model=TireInspectionResponse, status_code=status.HTTP_201_CREATED)
def create_tire_inspection(
    payload: TireInspectionCreate,
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

    item = TireInspection(
        project_id=payload.project_id,
        equipment_id=payload.equipment_id,
        work_date=payload.work_date,
        hour_meter=payload.hour_meter,
        location=caps(payload.location),
        inspector_name=clean_name(payload.inspector_name),
        supervisor_name=clean_supervisor(payload.supervisor_name),
        reviewed_by=clean_text(payload.reviewed_by),
        approved_by=clean_text(payload.approved_by),
        general_notes=caps(payload.general_notes),
        created_by=current_user.id,
        updated_by=current_user.id,
    )
    for detail in payload.details:
        item.details.append(
            TireDetail(
                position=clean_text(detail.position),
                status=clean_text(detail.status),
                vida_util=detail.vida_util,
                notes=caps(detail.notes),
                photo_url=detail.photo_url,
            )
        )

    db.add(item)
    db.commit()
    db.refresh(item)
    item = db.query(TireInspection).options(selectinload(TireInspection.details)).filter(TireInspection.id == item.id).first()
    return to_response(item)


@router.get("/{item_id}", response_model=TireInspectionResponse)
def get_tire_inspection(
    item_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    item = db.query(TireInspection).options(selectinload(TireInspection.details)).filter(TireInspection.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Inspección de llantas no encontrada")
    return to_response(item)


@router.put("/{item_id}", response_model=TireInspectionResponse)
def update_tire_inspection(
    item_id: int,
    payload: TireInspectionUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    item = db.query(TireInspection).options(selectinload(TireInspection.details)).filter(TireInspection.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Inspección de llantas no encontrada")

    data = payload.model_dump(exclude_unset=True)
    details = data.pop("details", None)
    for key in ["reviewed_by", "approved_by"]:
        if key in data and data[key] is not None:
            data[key] = clean_text(data[key])
    for key in ["location", "general_notes"]:
        if key in data and data[key] is not None:
            data[key] = caps(data[key])
    if "inspector_name" in data and data["inspector_name"] is not None:
        data["inspector_name"] = clean_name(data["inspector_name"])
    if "supervisor_name" in data and data["supervisor_name"] is not None:
        data["supervisor_name"] = clean_supervisor(data["supervisor_name"])
    for key, value in data.items():
        setattr(item, key, value)
    item.updated_by = current_user.id

    if details is not None:
        item.details.clear()
        for detail in details:
            item.details.append(
                TireDetail(
                    position=clean_text(detail["position"]),
                    status=clean_text(detail.get("status", "Bueno")),
                    vida_util=detail.get("vida_util"),
                    notes=caps(detail.get("notes")),
                    photo_url=detail.get("photo_url"),
                )
            )

    db.commit()
    db.refresh(item)
    item = db.query(TireInspection).options(selectinload(TireInspection.details)).filter(TireInspection.id == item.id).first()
    return to_response(item)


@router.delete("/{item_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_tire_inspection(
    item_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin")
    item = db.query(TireInspection).options(selectinload(TireInspection.details)).filter(TireInspection.id == item_id).first()
    if not item:
        raise HTTPException(status_code=404, detail="Inspección de llantas no encontrada")
    db.delete(item)
    db.commit()
    return None
