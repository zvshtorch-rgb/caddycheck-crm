"""
backup_to_onedrive.py
---------------------
1. Runs pg_dump against Supabase (custom format, compressed).
2. Uploads the resulting .dump file to OneDrive (zvikas@video-inform.com)
   under  Backups/Supabase/  via Microsoft Graph API.
3. Downloads every file from the app's Supabase Storage buckets (sent-invoices,
   order-pdfs, bank-payments, ticket-attachments, purchase-order-pdfs), zips them,
   and uploads that zip to OneDrive under Backups/SupabaseStorage/ (pg_dump alone
   only captures the storage.objects *metadata* rows, not the file bytes).
4. Keeps the last N backups of each kind in OneDrive (older ones are deleted).

Credentials are read from environment variables — set them in GitHub Secrets
or in a local .env file (never commit the .env file).

Required env vars
-----------------
SUPABASE_DB_PASSWORD       Supabase database password
GRAPH_TENANT_ID            Azure AD tenant ID        (ORDER_INTAKE_GRAPH_TENANT_ID)
GRAPH_CLIENT_ID            App registration client ID (ORDER_INTAKE_GRAPH_CLIENT_ID)
GRAPH_CLIENT_SECRET        App registration secret    (ORDER_INTAKE_GRAPH_CLIENT_SECRET)
ONEDRIVE_USER_EMAIL        Target mailbox / OneDrive owner  (zvikas@video-inform.com)

Optional env vars
-----------------
SUPABASE_HOST              default: aws-1-ap-southeast-1.pooler.supabase.com
SUPABASE_PORT              default: 5432
SUPABASE_USER              default: postgres.rdoxihpmghrvroddnkdi
SUPABASE_DB                default: postgres
SUPABASE_URL               Supabase project URL (needed for the storage-bucket backup)
SUPABASE_SERVICE_ROLE_KEY  Supabase service role key (needed for the storage-bucket backup)
ONEDRIVE_FOLDER            default: Backups/Supabase
ONEDRIVE_STORAGE_FOLDER    default: Backups/SupabaseStorage
BACKUP_KEEP_COUNT          default: 30  (how many backups of each kind to keep on OneDrive)
PG_DUMP_PATH               default: pg_dump  (must be on PATH in CI)
"""

import os
import subprocess
import sys
import math
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

import requests

# ── Configuration ─────────────────────────────────────────────────────────────

SUPABASE_HOST = os.environ.get("SUPABASE_HOST", "aws-1-ap-southeast-1.pooler.supabase.com")
SUPABASE_PORT = int(os.environ.get("SUPABASE_PORT", "5432"))
SUPABASE_USER = os.environ.get("SUPABASE_USER", "postgres.rdoxihpmghrvroddnkdi")
SUPABASE_DB   = os.environ.get("SUPABASE_DB",   "postgres")
SUPABASE_PASS = os.environ.get("SUPABASE_DB_PASSWORD", "")

TENANT_ID     = os.environ.get("GRAPH_TENANT_ID") or os.environ.get("ORDER_INTAKE_GRAPH_TENANT_ID", "")
CLIENT_ID     = os.environ.get("GRAPH_CLIENT_ID") or os.environ.get("ORDER_INTAKE_GRAPH_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET") or os.environ.get("ORDER_INTAKE_GRAPH_CLIENT_SECRET", "")
ONEDRIVE_USER = os.environ.get("ONEDRIVE_USER_EMAIL") or "zvikas@video-inform.com"
ONEDRIVE_FOLDER = os.environ.get("ONEDRIVE_FOLDER", "Backups/Supabase").strip("/")
ONEDRIVE_STORAGE_FOLDER = os.environ.get("ONEDRIVE_STORAGE_FOLDER", "Backups/SupabaseStorage").strip("/")
BACKUP_KEEP   = int(os.environ.get("BACKUP_KEEP_COUNT", "30"))
PG_DUMP_PATH  = os.environ.get("PG_DUMP_PATH", "pg_dump")

SUPABASE_URL = (os.environ.get("SUPABASE_URL") or "").strip().rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = (os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or "").strip()

# Every Supabase Storage bucket the app writes files into (see services/supabase_service.py
# and services/order_approval_service.py). pg_dump only captures the metadata *rows*
# referencing these files, not the file bytes themselves.
STORAGE_BUCKETS = [
    "sent-invoices",
    "order-pdfs",
    "bank-payments",
    "ticket-attachments",
    "purchase-order-pdfs",
]

CHUNK_SIZE = 5 * 1024 * 1024  # 5 MB upload chunks


def _require_env() -> None:
    missing = [k for k, v in {
        "SUPABASE_DB_PASSWORD": SUPABASE_PASS,
        "GRAPH_TENANT_ID":      TENANT_ID,
        "GRAPH_CLIENT_ID":      CLIENT_ID,
        "GRAPH_CLIENT_SECRET":  CLIENT_SECRET,
    }.items() if not v]
    if missing:
        sys.exit(f"ERROR: missing required environment variables: {', '.join(missing)}")


# ── Microsoft Graph helpers ────────────────────────────────────────────────────

def _get_token() -> str:
    """Obtain an app-only access token via client-credentials grant."""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
        "grant_type":    "client_credentials",
    }, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _graph_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _upload_file(local_path: Path, token: str, folder: str = ONEDRIVE_FOLDER) -> str:
    """Upload a file to OneDrive using an upload session (handles any size).

    Returns the OneDrive item ID of the uploaded file.
    """
    filename  = local_path.name
    file_size = local_path.stat().st_size
    onedrive_path = f"{folder}/{filename}"

    # 1. Create upload session
    session_url = (
        f"https://graph.microsoft.com/v1.0"
        f"/users/{ONEDRIVE_USER}/drive/root:/{onedrive_path}:/createUploadSession"
    )
    session_resp = requests.post(
        session_url,
        headers=_graph_headers(token),
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    session_resp.raise_for_status()
    upload_url = session_resp.json()["uploadUrl"]

    # 2. Upload in chunks
    num_chunks = math.ceil(file_size / CHUNK_SIZE)
    item_id = None
    with open(local_path, "rb") as fh:
        for chunk_index in range(num_chunks):
            start = chunk_index * CHUNK_SIZE
            chunk = fh.read(CHUNK_SIZE)
            end   = start + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range":  f"bytes {start}-{end}/{file_size}",
            }
            up_resp = requests.put(upload_url, data=chunk, headers=headers, timeout=120)
            up_resp.raise_for_status()
            print(f"  chunk {chunk_index + 1}/{num_chunks} uploaded ({end + 1}/{file_size} bytes)")
            if up_resp.status_code in (200, 201):
                item_id = up_resp.json().get("id")

    print(f"Upload complete: OneDrive/{onedrive_path}  (item id: {item_id})")
    return item_id or ""


def _list_backup_files(token: str, folder: str = ONEDRIVE_FOLDER, extension: str = ".dump") -> list[dict]:
    """Return OneDrive items in the given folder, sorted by name (oldest first)."""
    url = (
        f"https://graph.microsoft.com/v1.0"
        f"/users/{ONEDRIVE_USER}/drive/root:/{folder}:/children"
        f"?$select=id,name,createdDateTime&$orderby=name asc&$top=200"
    )
    resp = requests.get(url, headers=_graph_headers(token), timeout=30)
    if resp.status_code == 404:
        return []  # folder does not exist yet
    resp.raise_for_status()
    items = resp.json().get("value", [])
    return [i for i in items if i.get("name", "").endswith(extension)]


def _delete_item(item_id: str, token: str) -> None:
    url = f"https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER}/drive/items/{item_id}"
    resp = requests.delete(url, headers=_graph_headers(token), timeout=30)
    resp.raise_for_status()


def _rotate_old_backups(token: str, folder: str = ONEDRIVE_FOLDER, extension: str = ".dump") -> None:
    """Delete the oldest backup files when more than BACKUP_KEEP exist."""
    items = _list_backup_files(token, folder, extension)
    if len(items) <= BACKUP_KEEP:
        return
    to_delete = items[:len(items) - BACKUP_KEEP]
    for item in to_delete:
        print(f"  deleting old backup: {item['name']}")
        _delete_item(item["id"], token)
    print(f"Rotation complete: kept {BACKUP_KEEP}, deleted {len(to_delete)}.")


# ── Supabase Storage bucket backup ─────────────────────────────────────────────

def _storage_headers() -> dict:
    return {
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
    }


def _list_storage_objects(bucket: str, prefix: str = "") -> list[str]:
    """Recursively list every file path in a Supabase Storage bucket."""
    paths: list[str] = []
    url = f"{SUPABASE_URL}/storage/v1/object/list/{bucket}"
    body = {
        "prefix": prefix,
        "limit": 1000,
        "offset": 0,
        "sortBy": {"column": "name", "order": "asc"},
    }
    resp = requests.post(url, headers=_storage_headers(), json=body, timeout=60)
    resp.raise_for_status()
    for entry in resp.json():
        name = entry.get("name", "")
        if not name:
            continue
        full_path = f"{prefix}/{name}" if prefix else name
        # Supabase Storage represents folders as entries with no id/metadata.
        if entry.get("id") is None and entry.get("metadata") is None:
            paths.extend(_list_storage_objects(bucket, full_path))
        else:
            paths.append(full_path)
    return paths


def _download_storage_object(bucket: str, path: str) -> bytes:
    url = f"{SUPABASE_URL}/storage/v1/object/{bucket}/{path}"
    resp = requests.get(url, headers=_storage_headers(), timeout=120)
    resp.raise_for_status()
    return resp.content


def _backup_storage_buckets(tmp_dir: Path, timestamp: str) -> Optional[Path]:
    """Download every file from the app's Storage buckets into a single zip.

    Returns the zip path, or None if Supabase Storage credentials aren't configured.
    """
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        print(
            "Skipping storage-bucket backup: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY "
            "not set (only the database dump will be backed up)."
        )
        return None

    zip_path = tmp_dir / f"caddycheck_storage_{timestamp}.zip"
    total_files = 0
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for bucket in STORAGE_BUCKETS:
            try:
                object_paths = _list_storage_objects(bucket)
            except Exception as exc:  # noqa: BLE001
                print(f"  could not list bucket '{bucket}': {exc}")
                continue
            print(f"  bucket '{bucket}': {len(object_paths)} file(s)")
            for obj_path in object_paths:
                try:
                    data = _download_storage_object(bucket, obj_path)
                except Exception as exc:  # noqa: BLE001
                    print(f"    failed to download {bucket}/{obj_path}: {exc}")
                    continue
                zf.writestr(f"{bucket}/{obj_path}", data)
                total_files += 1

    size_mb = zip_path.stat().st_size / 1024 / 1024
    print(f"Storage backup complete: {total_files} file(s) zipped -> {zip_path} ({size_mb:.2f} MB)")
    return zip_path


# ── pg_dump ────────────────────────────────────────────────────────────────────

def _run_pg_dump(output_file: Path) -> None:
    timestamp_str = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    print(f"Running pg_dump → {output_file}")
    env = os.environ.copy()
    env["PGPASSWORD"] = SUPABASE_PASS
    result = subprocess.run(
        [
            PG_DUMP_PATH,
            "--host",     SUPABASE_HOST,
            "--port",     str(SUPABASE_PORT),
            "--username", SUPABASE_USER,
            "--dbname",   SUPABASE_DB,
            "--format",   "custom",
            "--file",     str(output_file),
            "--verbose",
        ],
        env=env,
        capture_output=False,
    )
    if result.returncode != 0:
        sys.exit(f"pg_dump failed with exit code {result.returncode}")
    size_mb = output_file.stat().st_size / 1024 / 1024
    print(f"pg_dump complete: {output_file}  ({size_mb:.2f} MB)")


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    _require_env()

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    filename  = f"caddycheck_supabase_{timestamp}.dump"

    with tempfile.TemporaryDirectory() as tmp_dir:
        dump_path = Path(tmp_dir) / filename

        # 1. Dump
        _run_pg_dump(dump_path)

        # 2. Upload to OneDrive
        print(f"\nUploading to OneDrive/{ONEDRIVE_FOLDER}/{filename} …")
        token = _get_token()
        _upload_file(dump_path, token, ONEDRIVE_FOLDER)

        # 3. Rotate old database backups
        print(f"\nChecking backup rotation (keep={BACKUP_KEEP}) …")
        token = _get_token()  # refresh — upload may have taken a while
        _rotate_old_backups(token, ONEDRIVE_FOLDER, ".dump")

        # 4. Back up Supabase Storage bucket files (not covered by pg_dump)
        print("\nBacking up Supabase Storage buckets …")
        storage_zip_path = _backup_storage_buckets(Path(tmp_dir), timestamp)
        if storage_zip_path is not None:
            print(f"\nUploading to OneDrive/{ONEDRIVE_STORAGE_FOLDER}/{storage_zip_path.name} …")
            token = _get_token()
            _upload_file(storage_zip_path, token, ONEDRIVE_STORAGE_FOLDER)

            print(f"\nChecking storage backup rotation (keep={BACKUP_KEEP}) …")
            token = _get_token()
            _rotate_old_backups(token, ONEDRIVE_STORAGE_FOLDER, ".zip")

    print("\nBackup finished successfully.")


if __name__ == "__main__":
    main()
