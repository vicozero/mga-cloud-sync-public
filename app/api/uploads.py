from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from app.api.deps import get_current_user, require_roles
from app.models.user import User
from app.services.storage import save_image

router = APIRouter(prefix="/uploads", tags=["uploads"])


@router.post("/mechanic-evidence")
def upload_mechanic_evidence(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    try:
        path = save_image(file, "mechanic-evidence")
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"file_url": path}


@router.post("/tire-photo")
def upload_tire_photo(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    require_roles(current_user, "admin", "supervisor")
    try:
        path = save_image(file, "tire-photos")
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return {"file_url": path}
