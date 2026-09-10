from datetime import datetime

from pydantic import BaseModel


class CategoryCreate(BaseModel):
    name: str
    sort_order: int = 0


class CategoryUpdate(BaseModel):
    name: str | None = None
    sort_order: int | None = None


class CategoryMerge(BaseModel):
    from_id: int
    into_id: int


class CategoryResponse(BaseModel):
    id: int
    name: str
    sort_order: int
    equipment_count: int
    created_at: datetime
    updated_at: datetime


class UnlistedCategory(BaseModel):
    name: str
    equipment_count: int


class CategorySummary(BaseModel):
    categories: list[CategoryResponse]
    unlisted: list[UnlistedCategory]