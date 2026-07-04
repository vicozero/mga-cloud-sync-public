from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Reset portal and catalog snapshots without touching operational tables."
    )
    parser.add_argument(
        "--database-url",
        default="",
        help="Optional DATABASE_URL override. If omitted, uses the current environment.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be deleted without changing the database.",
    )
    parser.add_argument(
        "--portal-only",
        action="store_true",
        help="Reset only the portal snapshot.",
    )
    parser.add_argument(
        "--catalog-only",
        action="store_true",
        help="Reset only the catalog snapshot.",
    )
    parser.add_argument(
        "--backup-dir",
        default="",
        help="Optional directory where the current snapshot payloads will be saved before deletion.",
    )
    return parser


def timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def write_backup(backup_dir: Path, name: str, payload: dict[str, object]) -> Path:
    backup_dir.mkdir(parents=True, exist_ok=True)
    file_path = backup_dir / f"{name}_{timestamp()}.json"
    file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return file_path


def main() -> int:
    args = build_parser().parse_args()

    if args.database_url:
        os.environ["DATABASE_URL"] = args.database_url

    project_root = Path(__file__).resolve().parents[1]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    from render_backend.main import CatalogSnapshot, PortalSnapshot, SessionLocal, json_loads  # noqa: WPS433

    reset_portal = not args.catalog_only
    reset_catalog = not args.portal_only
    if args.portal_only and args.catalog_only:
        raise SystemExit("Use only one of --portal-only or --catalog-only, or omit both to reset both.")

    backup_dir = Path(args.backup_dir).expanduser() if args.backup_dir else None

    with SessionLocal() as session:
        actions: list[str] = []

        if reset_portal:
            row = session.query(PortalSnapshot).filter(PortalSnapshot.name == "default").one_or_none()
            if row is None:
                print("Portal snapshot: no row found.")
            else:
                payload = json_loads(row.payload_json)
                if backup_dir is not None:
                    path = write_backup(backup_dir, "portal_snapshot", payload if isinstance(payload, dict) else {})
                    print(f"Portal snapshot backed up to: {path}")
                print(f"Portal snapshot row found. payload_type={type(payload).__name__}")
                if not args.dry_run:
                    session.delete(row)
                    actions.append("portal")

        if reset_catalog:
            row = session.query(CatalogSnapshot).filter(CatalogSnapshot.name == "default").one_or_none()
            if row is None:
                print("Catalog snapshot: no row found.")
            else:
                payload = json_loads(row.payload_json)
                if backup_dir is not None:
                    path = write_backup(backup_dir, "catalog_snapshot", payload if isinstance(payload, dict) else {})
                    print(f"Catalog snapshot backed up to: {path}")
                print(f"Catalog snapshot row found. payload_type={type(payload).__name__}")
                if not args.dry_run:
                    session.delete(row)
                    actions.append("catalog")

        if args.dry_run:
            print("Dry run complete. No changes were committed.")
            return 0

        if actions:
            session.commit()
            print(f"Reset completed for: {', '.join(actions)}")
        else:
            print("No snapshot rows were deleted.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
