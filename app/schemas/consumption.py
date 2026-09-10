from datetime import date, datetime

from pydantic import BaseModel


class ConsumptionCreate(BaseModel):
    project_id: int
    equipment_id: int
    work_date: date
    shift: int
    lubricant: str
    quantity: float
    unit: str = "L"
    notes: str | None = None
    supervisor_name: str | None = None


class ConsumptionUpdate(BaseModel):
    work_date: date | None = None
    shift: int | None = None
    lubricant: str | None = None
    quantity: float | None = None
    unit: str | None = None
    notes: str | None = None
    supervisor_name: str | None = None


class ConsumptionResponse(BaseModel):
    id: int
    project_id: int
    equipment_id: int
    work_date: date
    shift: int
    lubricant: str
    quantity: float
    unit: str
    notes: str | None
    supervisor_name: str | None
    created_by: int | None
    updated_by: int | None
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True
