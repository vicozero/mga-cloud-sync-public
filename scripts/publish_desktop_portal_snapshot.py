from __future__ import annotations

import argparse
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


DEFAULT_CONFIG_PATH = Path(r"C:\ProgramData\MGA Mantenimiento\mga_config.json")
DEFAULT_APP_PATH = Path(r"D:\SIIS\New project")
DEFAULT_API_URL = "https://mga-cloud-sync.onrender.com"


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(
        description="Publish the exact MGA desktop portal snapshot, including KPI reports, to Render."
    )
    result.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help="MGA desktop config JSON path.")
    result.add_argument("--app-path", default=str(DEFAULT_APP_PATH), help="Desktop application source folder containing app.py.")
    result.add_argument("--db-path", default="", help="Override desktop SQLite database path.")
    result.add_argument("--api-url", default="", help="Override cloud API base URL.")
    result.add_argument("--api-key", default="", help="Override API key. Defaults to config cloud_api_key.")
    return result


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def post_json(url: str, payload: dict[str, Any], api_key: str) -> dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "X-MGA-API-Key": api_key,
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            response_body = response.read().decode("utf-8", errors="replace")
            return json.loads(response_body) if response_body else {"ok": True}
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {error_body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"No se pudo conectar a Render: {exc}") from exc


def main() -> int:
    args = parser().parse_args()
    config = load_config(Path(args.config))
    app_path = Path(args.app_path)
    db_path = Path(args.db_path or config.get("db_path") or r"C:\ProgramData\MGA Mantenimiento\mga_mantenimiento.db")
    api_url = str(args.api_url or config.get("cloud_url") or DEFAULT_API_URL).rstrip("/")
    api_key = str(args.api_key or config.get("cloud_api_key") or "").strip()

    if not app_path.joinpath("app.py").exists():
        raise SystemExit(f"No se encontro app.py en: {app_path}")
    if not db_path.exists():
        raise SystemExit(f"No se encontro la base de datos: {db_path}")
    if not api_key:
        raise SystemExit("No hay cloud_api_key en la configuracion local.")

    sys.path.insert(0, str(app_path))
    import app  # type: ignore

    repo = app.Repository(db_path)
    try:
        portal = app.build_cloud_portal_snapshot(repo)
    finally:
        repo.conn.close()

    response = post_json(f"{api_url}/api/portal/snapshot", portal, api_key)
    if not response.get("ok"):
        raise SystemExit(f"Render rechazo la sincronizacion: {response}")

    kpi_reports = portal.get("kpi_reports") if isinstance(portal.get("kpi_reports"), dict) else {}
    print("Sincronizacion publicada correctamente.")
    print(f"Web: {api_url}/almacen-filtros")
    print(f"Base: {db_path}")
    print(f"Equipos: {len(portal.get('equipment') or [])}")
    print(f"Capturas: {len(portal.get('captures') or [])}")
    print(f"Reportes KPI: {len(kpi_reports)}")
    print(f"Generado: {portal.get('generated_at')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
