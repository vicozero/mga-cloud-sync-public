from sqlalchemy.orm import Session

from app.core.security import get_password_hash
from app.models.user import User


DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "admin123"
DEFAULT_ADMIN_NAME = "Administrador"


def create_default_admin(db: Session) -> User:
    user = db.query(User).filter(User.username == DEFAULT_ADMIN_USERNAME).first()
    if user:
        return user

    user = User(
        username=DEFAULT_ADMIN_USERNAME,
        full_name=DEFAULT_ADMIN_NAME,
        password_hash=get_password_hash(DEFAULT_ADMIN_PASSWORD),
        role="admin",
        active=True,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    return user
