import json
import urllib.error
import urllib.request

from app.core.config import get_settings

TIMEOUT = 30


def is_cloud() -> bool:
    s = get_settings()
    return bool(s.supabase_url.strip() and s.supabase_service_role_key.strip())


def bucket() -> str:
    return get_settings().supabase_bucket.strip() or "disponibilidad"


def _base_url() -> str:
    return get_settings().supabase_url.strip().rstrip("/")


def _headers() -> dict:
    key = get_settings().supabase_service_role_key.strip()
    return {"apikey": key, "Authorization": f"Bearer {key}"}


def _call(method: str, url: str, body: bytes | None = None, content_type: str = "application/json") -> tuple[int, bytes]:
    headers = _headers()
    if content_type:
        headers["Content-Type"] = content_type
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        raw = exc.read()
        return exc.code, raw
    except urllib.error.URLError:
        return 0, b""


def public_url(relative_path: str) -> str:
    return f"{_base_url()}/storage/v1/object/public/{bucket()}/{relative_path.lstrip('/')}"


def ensure_bucket() -> bool:
    if not is_cloud():
        return False
    status, raw = _call("GET", f"{_base_url()}/storage/v1/bucket")
    if status in (200, 201):
        try:
            existing = [item.get("id") for item in json.loads(raw)]
            if bucket() in existing:
                return True
        except Exception:
            pass
    body = json.dumps({"id": bucket(), "name": bucket(), "public": True}).encode()
    status, _raw = _call("POST", f"{_base_url()}/storage/v1/bucket", body)
    return status in (200, 201)


def upload_bytes(relative_path: str, data: bytes, content_type: str) -> bool:
    path = relative_path.lstrip("/")
    status, _raw = _call("POST", f"{_base_url()}/storage/v1/object/{bucket()}/{path}", data, content_type)
    return status in (200, 201)


def get_public_bytes(relative_path: str) -> bytes | None:
    status, raw = _call("GET", f"{_base_url()}/storage/v1/object/{bucket()}/{relative_path.lstrip('/')}")
    return raw if status == 200 else None


def delete_object(relative_path: str) -> bool:
    status, _raw = _call("DELETE", f"{_base_url()}/storage/v1/object/{bucket()}/{relative_path.lstrip('/')}")
    return status in (200, 201)


def list_objects(prefix: str = "") -> list[dict]:
    body = json.dumps({"prefix": prefix.rstrip("/") + "/" if prefix else prefix, "limit": 1000, "offset": 0}).encode()
    status, raw = _call("POST", f"{_base_url()}/storage/v1/object/list/{bucket()}", body)
    if status not in (200, 201):
        return []
    try:
        items = json.loads(raw)
    except Exception:
        return []
    result = []
    for item in items:
        name = item.get("name", "")
        if not name:
            continue
        metadata = item.get("metadata", {}) or {}
        result.append(
            {
                "name": name,
                "size": metadata.get("size", 0),
                "lastModified": metadata.get("lastModified", ""),
            }
        )
    return result