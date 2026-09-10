from datetime import datetime

from pydantic import BaseModel


class EquipmentCreate(BaseModel):
    project_id: int
    category: str
    name: str
    eco: str
    current_status: str = "DISPONIBLE"
    notes: str | None = None
    hour_meter: str | None = None
    active: bool = True
    sort_order: int = 0


class EquipmentUpdate(BaseModel):
    category: str | None = None
    name: str | None = None
    eco: str | None = None
    current_status: str | None = None
    notes: str | None = None
    hour_meter: str | None = None
    active: bool | None = None
    sort_order: int | None = None


class EquipmentOrder(BaseModel):
    project_id: int
    ids: list[int]


class EquipmentResponse(BaseModel):
    id: int
    project_id: int
    category: str
    name: str
    eco: str
    current_status: str
    notes: str | None
    hour_meter: str | None
    active: bool
    sort_order: int
    created_by: int | None
    updated_by: int | None
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True
