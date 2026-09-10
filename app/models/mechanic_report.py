from datetime import date, datetime

from sqlalchemy import Date, DateTime, ForeignKey, SmallInteger, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.db.base import Base


class MechanicReport(Base):
    __tablename__ = "mechanic_reports"

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    project_id: Mapped[int] = mapped_column(ForeignKey("projects.id"), index=True)
    equipment_id: Mapped[int] = mapped_column(ForeignKey("equipment.id"), index=True)
    work_date: Mapped[date] = mapped_column(Date, index=True)
    shift: Mapped[int] = mapped_column(SmallInteger, index=True)
    mechanic_name: Mapped[str] = mapped_column(String(120))
    supervisor_name: Mapped[str | None] = mapped_column(String(120), nullable=True)
    hour_meter: Mapped[str | None] = mapped_column(String(40), nullable=True)
    service_type: Mapped[str] = mapped_column(String(30), default="Correctivo")
    failure: Mapped[str] = mapped_column(Text)
    work_done: Mapped[str] = mapped_column(Text)
    parts_used: Mapped[str | None] = mapped_column(Text, nullable=True)
    final_status: Mapped[str] = mapped_column(String(30), default="Disponible")
    evidence_url: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"), nullable=True)
    updated_by: Mapped[int | None] = mapped_column(ForeignKey("users.id"), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
