from datetime import datetime

from pydantic import BaseModel


class ProjectCreate(BaseModel):
    name: str
    description: str | None = None
    active: bool = True


class ProjectUpdate(BaseModel):
    name: str | None = None
    description: str | None = None
    active: bool | None = None


class ProjectResponse(BaseModel):
    id: int
    name: str
    description: str | None
    active: bool
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True
