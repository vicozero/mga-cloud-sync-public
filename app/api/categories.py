from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import func
from sqlalchemy.orm import Session

from app.api.deps import get_current_user, require_roles
from app.db.session import get_db
from app.models.category import Category
from app.models.equipment import Equipment
from app.models.user import User
from app.schemas.category import CategoryCreate, CategoryMerge, CategorySummary, CategoryUpdate, UnlistedCategory
from app.services.text_clean import clean_category

router = APIRouter(prefix="/categories", tags=["categories"])


def _summary(db: Session) -> CategorySummary:
    rows = (
        db.query(
            Category.id,
            Category.name,
            Category.sort_order,
            Category.created_at,
            Category.updated_at,
            func.count(Equipment.id),
        )
        .outerjoin(Equipment, func.upper(Equipment.category) == func.upper(Category.name))
        .group_by(Category.id, Category.name, Category.sort_order, Category.created_at, Category.updated_at)
        .order_by(Category.sort_order.asc(), func.upper(Category.name).asc())
        .all()
    )
    categories = [
        {
            "id": row.id,
            "name": row.name,
            "sort_order": row.sort_order,
            "created_at": row.created_at,
            "updated_at": row.updated_at,
            "equipment_count": row[5],
        }
        for row in rows
    ]
    known = {row.name.upper() for row in rows}
    used = (
        db.query(Equipment.category, func.count(Equipment.id))
        .filter(Equipment.active.is_(True))
        .group_by(Equipment.category)
        .all()
    )
    unlisted = [
        UnlistedCategory(name=name, equipment_count=count_)
        for name, count_ in used
        if name and name.strip().upper() not in known
    ]
    unlisted.sort(key=lambda item: item.name.upper())
    return CategorySummary(categories=categories, unlisted=unlisted)


@router.get("", response_model=CategorySummary)
def list_categories(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    return _summary(db)


@router.post("", response_model=CategorySummary, status_code=status.HTTP_201_CREATED)
def create_category(
    payload: CategoryCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    name = payload.name.strip().upper()
    if not name:
        raise HTTPException(status_code=400, detail="El nombre no puede ir vacío")
    exists = db.query(Category).filter(func.upper(Category.name) == name).first()
    if exists:
        raise HTTPException(status_code=409, detail="Ya existe la categoría")

    last_sort = db.query(func.max(Category.sort_order)).scalar() or 0
    db.add(Category(name=name, sort_order=payload.sort_order or last_sort + 1))
    db.commit()
    return _summary(db)


@router.put("/{category_id}", response_model=CategorySummary)
def update_category(
    category_id: int,
    payload: CategoryUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    category = db.query(Category).filter(Category.id == category_id).first()
    if not category:
        raise HTTPException(status_code=404, detail="Categoría no encontrada")

    data = payload.model_dump(exclude_unset=True)
    if "name" in data and data["name"] is not None:
        new_name = clean_category(data["name"])
        if not new_name:
            raise HTTPException(status_code=400, detail="El nombre no puede ir vacío")
        conflict = (
            db.query(Category)
            .filter(Category.id != category_id, func.upper(Category.name) == new_name)
            .first()
        )
        if conflict:
            raise HTTPException(
                status_code=409,
                detail="Ya existe otra categoría con ese nombre. Usa 'Unificar' para fusionarlas.",
            )
        old_name = category.name
        if new_name != old_name:
            db.query(Equipment).filter(func.upper(Equipment.category) == func.upper(old_name)).update(
                {"category": new_name, "updated_by": current_user.id}, synchronize_session=False
            )
        category.name = new_name
    if "sort_order" in data and data["sort_order"] is not None:
        category.sort_order = data["sort_order"]
    db.commit()
    return _summary(db)


@router.post("/merge", response_model=CategorySummary)
def merge_categories(
    payload: CategoryMerge,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    if payload.from_id == payload.into_id:
        raise HTTPException(status_code=400, detail="La categoría de origen y destino son la misma")
    from_category = db.query(Category).filter(Category.id == payload.from_id).first()
    into_category = db.query(Category).filter(Category.id == payload.into_id).first()
    if not from_category or not into_category:
        raise HTTPException(status_code=404, detail="Categoría no encontrada")

    db.query(Equipment).filter(func.upper(Equipment.category) == func.upper(from_category.name)).update(
        {"category": into_category.name, "updated_by": current_user.id}, synchronize_session=False
    )
    db.delete(from_category)
    db.commit()
    return _summary(db)


@router.delete("/{category_id}", response_model=CategorySummary)
def delete_category(
    category_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin")
    category = db.query(Category).filter(Category.id == category_id).first()
    if not category:
        raise HTTPException(status_code=404, detail="Categoría no encontrada")

    used = (
        db.query(func.count(Equipment.id))
        .filter(func.upper(Equipment.category) == func.upper(category.name))
        .scalar()
        or 0
    )
    if used:
        raise HTTPException(
            status_code=400,
            detail="La categoría tiene equipos asignados. Usa 'Unificar' para moverlos a otra categoría.",
        )
    db.delete(category)
    db.commit()
    return _summary(db)