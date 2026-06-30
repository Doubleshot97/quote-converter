"""
supabase_store.py — persist each conversion (metadata + both Excel files) to
a *private* Supabase project.

Credentials are read from Streamlit secrets or environment variables and are
never sent to the browser. Storage is deliberately fail-soft: if Supabase
isn't configured, or a write fails, the functions return False and the app
keeps working — logging must never block a download.

Required config (set as Hugging Face Space secrets, or .streamlit/secrets.toml
for local dev):
    SUPABASE_URL                 e.g. https://xxxxxxxx.supabase.co
    SUPABASE_SERVICE_ROLE_KEY    the service_role key (server-side only)

See supabase_schema.sql for the table + bucket setup.
"""

import os
import datetime as _dt

try:                      # available when running inside Streamlit
    import streamlit as st
except Exception:         # importable standalone for tests
    st = None

BUCKET = "quote-files"
TABLE = "quote_conversions"
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _get_secret(name: str):
    """Read a secret from st.secrets (Streamlit Cloud) or env (HF Spaces)."""
    if st is not None:
        try:
            if name in st.secrets:
                return st.secrets[name]
        except Exception:
            pass
    return os.environ.get(name)


def is_configured() -> bool:
    return bool(_get_secret("SUPABASE_URL")
                and _get_secret("SUPABASE_SERVICE_ROLE_KEY"))


def _client():
    url = _get_secret("SUPABASE_URL")
    key = _get_secret("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        return None
    from supabase import create_client
    return create_client(url, key)


def store_conversion(
    *,
    file_hash: str,
    source_filename: str,
    input_type: str,            # "pdf" or "po_xlsx"
    supplier: str | None,
    job_code: str | None,
    activity_code: str | None,
    line_count: int,
    subtotal_ex_gst: float,
    catalogue_row_count: int,
    po_bytes: bytes,
    po_name: str,
    catalogue_bytes: bytes,
    catalogue_name: str,
) -> bool:
    """Upload both xlsx files to the private bucket and insert one metadata
    row. Returns True on success, False if storage isn't configured or any
    step fails (the caller treats False as 'not yet logged' and can retry).
    """
    client = _client()
    if client is None:
        return False

    # Namespace each upload under a UTC timestamp so repeated uploads of the
    # same filename never collide in the bucket.
    stamp = _dt.datetime.utcnow().strftime("%Y/%m/%d/%H%M%S_%f")
    po_path = f"{stamp}/{po_name}"
    cat_path = f"{stamp}/{catalogue_name}"
    opts = {"content-type": _XLSX_MIME, "upsert": "false"}

    try:
        client.storage.from_(BUCKET).upload(po_path, po_bytes, opts)
        client.storage.from_(BUCKET).upload(cat_path, catalogue_bytes, opts)
        client.table(TABLE).insert({
            "file_hash": file_hash,
            "source_filename": source_filename,
            "input_type": input_type,
            "supplier": supplier,
            "job_code": job_code or None,
            "activity_code": activity_code or None,
            "line_count": int(line_count),
            "subtotal_ex_gst": round(float(subtotal_ex_gst), 2),
            "catalogue_row_count": int(catalogue_row_count),
            "po_file_path": po_path,
            "catalogue_file_path": cat_path,
        }).execute()
        return True
    except Exception:
        # Fail-soft: never surface a storage error into the download flow.
        return False
