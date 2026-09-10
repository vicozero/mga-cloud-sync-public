from datetime import date, datetime

from pydantic import BaseModel


class MechanicReportCreate(BaseModel):
    project_id: int
    equipment_id: int
    work_date: date
    shift: int
    mechanic_name: str
    supervisor_name: str | None = None
    hour_meter: str | None = None
    service_type: str = "Correctivo"
    failure: str
    work_done: str
    parts_used: str | None = None
    final_status: str = "Disponible"
    evidence_url: str | None = None


class MechanicReportUpdate(BaseModel):
    work_date: date | None = None
    shift: int | None = None
    mechanic_name: str | None = None
    supervisor_name: str | None = None
    hour_meter: str | None = None
    service_type: str | None = None
    failure: str | None = None
    work_done: str | None = None
    parts_used: str | None = None
    final_status: str | None = None
    evidence_url: str | None = None


class MechanicReportResponse(BaseModel):
    id: int
    project_id: int
    equipment_id: int
    work_date: date
    shift: int
    mechanic_name: str
    supervisor_name: str | None
    hour_meter: str | None
    service_type: str
    failure: str
    work_done: str
    parts_used: str | None
    final_status: str
    evidence_url: str | None
    created_by: int | None
    updated_by: int | None
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True
