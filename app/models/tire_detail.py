from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.db.base import Base


class TireDetail(Base):
    __tablename__ = "tire_details"
    __table_args__ = (UniqueConstraint("inspection_id", "position", name="uq_tire_position"),)

    id: Mapped[int] = mapped_column(primary_key=True, index=True)
    inspection_id: Mapped[int] = mapped_column(ForeignKey("tire_inspections.id"), index=True)
    position: Mapped[str] = mapped_column(String(30))
    status: Mapped[str] = mapped_column(String(20), default="Bueno")
    vida_util: Mapped[str | None] = mapped_column(String(20), nullable=True)
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    photo_url: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    inspection = relationship("TireInspection", back_populates="details")
