from app.db.base import Base
from app.db.session import SessionLocal, engine
from app.models import Category, Consumption, Equipment, MechanicReport, Project, Supervisor, TireDetail, TireInspection, User  # noqa: F401
from app.services.bootstrap import create_default_admin

DEFAULT_CATEGORIES = [
    "ACARREO",
    "BARRENACIÓN",
    "CONVENCIONAL",
    "PERSONAL",
    "REZAGADO",
    "SERVICIOS",
    "VEHÍCULO UTILITARIO",
    "OTROS",
]


def seed_default_categories(db) -> None:
    if db.query(Category).count() > 0:
        return
    for index, name in enumerate(DEFAULT_CATEGORIES, start=1):
        db.add(Category(name=name, sort_order=index))
    db.commit()


def init_db() -> None:
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    try:
        create_default_admin(db)
        seed_default_categories(db)
    finally:
        db.close()
