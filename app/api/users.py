from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from app.api.deps import get_current_user
from app.core.security import get_password_hash, verify_password
from app.db.session import get_db
from app.models.user import User
from app.schemas.user import (
    AdminPasswordResetRequest,
    PasswordChangeRequest,
    UserCreate,
    UserResponse,
    UserUpdate,
)


DEFAULT_BOOTSTRAP_ADMIN_USERNAME = "admin"

router = APIRouter(prefix="/users", tags=["users"])


def require_admin(current_user: User) -> None:
    if current_user.role.lower() != "admin":
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Solo un administrador puede realizar esta acción")


def count_active_admins(db: Session) -> int:
    return db.query(User).filter(User.role == "admin", User.active.is_(True)).count()


def validate_admin_safety(db: Session, user: User, next_role: str | None, next_active: bool | None) -> None:
    is_admin = user.role == "admin"
    will_be_admin = is_admin if next_role is None else next_role == "admin"
    will_be_active = user.active if next_active is None else next_active
    if is_admin and (not will_be_admin or not will_be_active) and count_active_admins(db) <= 1:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No puedes desactivar o cambiar el único administrador activo",
        )


@router.get("", response_model=list[UserResponse])
def list_users(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_admin(current_user)
    return db.query(User).order_by(User.full_name.asc()).all()


@router.post("/bootstrap-admin", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
def bootstrap_admin(
    payload: UserCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_admin(current_user)
    if current_user.username != DEFAULT_BOOTSTRAP_ADMIN_USERNAME or count_active_admins(db) > 1:
        raise HTTPException(status_code=400, detail="Bootstrap admin solo disponible durante configuración inicial")
    username = payload.username.strip().lower()
    if db.query(User).filter(User.username == username).first():
        raise HTTPException(status_code=409, detail="Ya existe un usuario con ese nombre")
    user = User(
        username=username,
        full_name=payload.full_name.strip(),
        password_hash=get_password_hash(payload.password),
        role="admin",
        active=True,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    return user


@router.post("", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
def create_user(
    payload: UserCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_admin(current_user)
    username = payload.username.strip().lower()
    if db.query(User).filter(User.username == username).first():
        raise HTTPException(status_code=409, detail="Ya existe un usuario con ese nombre")

    user = User(
        username=username,
        full_name=payload.full_name.strip(),
        password_hash=get_password_hash(payload.password),
        role=payload.role.strip().lower(),
        active=payload.active,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    return user


@router.put("/{user_id}", response_model=UserResponse)
def update_user(
    user_id: int,
    payload: UserUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_admin(current_user)
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")

    data = payload.model_dump(exclude_unset=True)
    if "username" in data and data["username"] is not None:
        normalized = data["username"].strip().lower()
        exists = db.query(User).filter(User.username == normalized, User.id != user_id).first()
        if exists:
            raise HTTPException(status_code=409, detail="Ya existe un usuario con ese nombre")
        data["username"] = normalized
    if "full_name" in data and data["full_name"] is not None:
        data["full_name"] = data["full_name"].strip()
    if "role" in data and data["role"] is not None:
        data["role"] = data["role"].strip().lower()

    validate_admin_safety(db, user, data.get("role"), data.get("active"))

    for key, value in data.items():
        setattr(user, key, value)
    db.commit()
    db.refresh(user)
    return user


@router.post("/me/change-password", status_code=status.HTTP_204_NO_CONTENT)
def change_my_password(
    payload: PasswordChangeRequest,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if not verify_password(payload.current_password, current_user.password_hash):
        raise HTTPException(status_code=400, detail="La contraseña actual no es correcta")
    current_user.password_hash = get_password_hash(payload.new_password)
    db.commit()
    return None


@router.post("/{user_id}/reset-password", status_code=status.HTTP_204_NO_CONTENT)
def reset_user_password(
    user_id: int,
    payload: AdminPasswordResetRequest,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    require_admin(current_user)
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="Usuario no encontrado")
    user.password_hash = get_password_hash(payload.new_password)
    db.commit()
    return None
