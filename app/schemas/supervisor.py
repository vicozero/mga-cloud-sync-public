from datetime import datetime

from pydantic import BaseModel


class SupervisorCreate(BaseModel):
    project_id: int
    name: str


class SupervisorUpdate(BaseModel):
    name: str | None = None
    active: bool | None = None
    sort_order: int | None = None


class SupervisorResponse(BaseModel):
    id: int
    project_id: int
    name: str
    active: bool
    sort_order: int
    record_count: int = 0
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True