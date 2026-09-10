from datetime import date, datetime

from pydantic import BaseModel


class TireDetailInput(BaseModel):
    position: str
    status: str = "Bueno"
    vida_util: str | None = None
    notes: str | None = None
    photo_url: str | None = None


class TireDetailResponse(TireDetailInput):
    id: int
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True


class TireInspectionCreate(BaseModel):
    project_id: int
    equipment_id: int
    work_date: date
    hour_meter: str | None = None
    location: str | None = None
    inspector_name: str | None = None
    supervisor_name: str | None = None
    reviewed_by: str | None = None
    approved_by: str | None = None
    general_notes: str | None = None
    details: list[TireDetailInput]


class TireInspectionUpdate(BaseModel):
    work_date: date | None = None
    hour_meter: str | None = None
    location: str | None = None
    inspector_name: str | None = None
    supervisor_name: str | None = None
    reviewed_by: str | None = None
    approved_by: str | None = None
    general_notes: str | None = None
    details: list[TireDetailInput] | None = None


class TireInspectionResponse(BaseModel):
    id: int
    project_id: int
    equipment_id: int
    work_date: date
    hour_meter: str | None
    location: str | None
    inspector_name: str | None
    supervisor_name: str | None
    reviewed_by: str | None
    approved_by: str | None
    general_notes: str | None
    details: list[TireDetailResponse]
    created_by: int | None
    updated_by: int | None
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True
