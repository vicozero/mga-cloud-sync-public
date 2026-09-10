from pathlib import Path
from uuid import uuid4

from fastapi import UploadFile

from app.core.config import get_settings
from app.services import file_store

ALLOWED_IMAGE_TYPES = {"image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp"}


def uploads_root() -> Path:
    settings = get_settings()
    return Path(settings.upload_dir).resolve()


def ensure_dir(relative_dir: str) -> Path:
    target = uploads_root() / relative_dir
    target.mkdir(parents=True, exist_ok=True)
    return target


def save_image(file: UploadFile, relative_dir: str) -> str:
    settings = get_settings()
    content_type = (file.content_type or "").lower()
    if content_type not in ALLOWED_IMAGE_TYPES:
        raise ValueError("Tipo de archivo no permitido. Usa JPG, PNG o WEBP")

    extension = ALLOWED_IMAGE_TYPES[content_type]
    filename = f"{uuid4().hex}{extension}"
    relative_path = f"/uploads/{relative_dir}/{filename}".replace("\\", "/")
    max_bytes = settings.max_upload_mb * 1024 * 1024

    if file_store.is_cloud():
        data = bytearray()
        while True:
            chunk = file.file.read(1024 * 1024)
            if not chunk:
                break
            data.extend(chunk)
            if len(data) > max_bytes:
                raise ValueError(f"El archivo supera el límite de {settings.max_upload_mb} MB")
        ok = file_store.upload_bytes(relative_path, bytes(data), content_type)
        if not ok:
            raise ValueError("Error al subir la imagen al almacenamiento en la nube")
        return relative_path

    directory = ensure_dir(relative_dir)
    target = directory / filename
    total = 0

    try:
        with target.open("wb") as output:
            while True:
                chunk = file.file.read(1024 * 1024)
                if not chunk:
                    break
                total += len(chunk)
                if total > max_bytes:
                    raise ValueError(f"El archivo supera el límite de {settings.max_upload_mb} MB")
                output.write(chunk)
    except Exception:
        if target.exists():
            target.unlink()
        raise

    return relative_path
