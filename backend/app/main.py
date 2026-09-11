import hashlib
import hmac
import io
import json
import os
import uuid
from pathlib import Path

from fastapi import Depends, FastAPI, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image
from pydantic import BaseModel

from . import blob_store, renderer

FRONTEND_DIR = Path(__file__).resolve().parent.parent.parent / "frontend"

SIGNATURE_EXT_BY_CONTENT_TYPE = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
}


def _validate_is_real_image(raw: bytes) -> None:
    """The client-supplied Content-Type header is attacker-controlled -- trusting it alone
    (as the extension lookup above does) lets arbitrary file content get saved with an
    image extension. Actually decoding the bytes confirms it's a real, intact image before
    it's allowed anywhere near a company's saved signature or a generated proposal."""
    try:
        img = Image.open(io.BytesIO(raw))
        img.verify()
    except Exception:
        raise HTTPException(400, "File is not a valid image")

AUTH_CONFIG_PATH = Path(__file__).resolve().parent.parent / "data" / "auth_config.json"
SESSION_MESSAGE = b"quickprop-session"


def _load_auth_config() -> dict:
    # On Vercel there's no local auth_config.json (it's a secret, kept out of the git repo
    # that gets deployed) -- the same two values are supplied as env vars instead.
    env_hash = os.environ.get("AUTH_PASSWORD_HASH")
    env_secret = os.environ.get("AUTH_TOKEN_SECRET")
    if env_hash and env_secret:
        return {"password_hash": env_hash, "token_secret": env_secret}
    with open(AUTH_CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _current_session_token() -> str:
    # One shared token derived from the server's secret -- valid for as long as the
    # secret stays the same, since this tool uses a single organization-wide password
    # rather than per-user accounts.
    secret = _load_auth_config()["token_secret"].encode()
    return hmac.new(secret, SESSION_MESSAGE, hashlib.sha256).hexdigest()


async def require_auth(authorization: str | None = Header(default=None)) -> None:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(401, "Not authenticated")
    token = authorization.removeprefix("Bearer ").strip()
    if not hmac.compare_digest(token, _current_session_token()):
        raise HTTPException(401, "Invalid or expired session")


class LoginRequest(BaseModel):
    password: str


app = FastAPI(title="QuickProp")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)


@app.post("/auth/login")
def login(req: LoginRequest):
    cfg = _load_auth_config()
    if hashlib.sha256(req.password.encode()).hexdigest() != cfg["password_hash"]:
        raise HTTPException(401, "Incorrect password")
    return {"token": _current_session_token()}


@app.get("/companies", dependencies=[Depends(require_auth)])
def list_companies():
    companies = renderer.load_companies()
    return [
        {
            "id": cid,
            "name": c["name"],
            "services": c["services"],
            "has_signature": renderer.signature_path(cid) is not None,
        }
        for cid, c in companies.items()
    ]


@app.post("/companies/{company}/signature", dependencies=[Depends(require_auth)])
async def upload_signature(company: str, file: UploadFile):
    companies = renderer.load_companies()
    if company not in companies:
        raise HTTPException(404, "Unknown company")
    # Anything that isn't actually PNG or JPEG got silently saved with a ".jpg" extension
    # it didn't match (e.g. a GIF or WEBP upload) -- python-pptx/PIL would later fail to
    # open it as JPEG when this signature gets embedded into a generated proposal.
    ext = SIGNATURE_EXT_BY_CONTENT_TYPE.get(file.content_type)
    if not ext:
        raise HTTPException(400, "Signature must be a PNG or JPEG image")
    raw = await file.read()
    _validate_is_real_image(raw)
    if blob_store.BLOB_ENABLED:
        for old in ("png", "jpg", "jpeg"):
            blob_store.blob_delete(f"signatures/{company}.{old}")
        blob_store.blob_put(f"signatures/{company}{ext}", raw)
        return {"ok": True, "path": f"signatures/{company}{ext}"}
    for old in ("png", "jpg", "jpeg"):
        old_path = renderer.SIGNATURES_DIR / f"{company}.{old}"
        if old_path.exists():
            old_path.unlink()
    dest = renderer.SIGNATURES_DIR / f"{company}{ext}"
    dest.write_bytes(raw)
    return {"ok": True, "path": str(dest)}


@app.post("/signatures/temp-upload", dependencies=[Depends(require_auth)])
async def upload_temp_signature(file: UploadFile):
    """One-off signature upload for a single generation (rep or client) -- not tied to
    any company, since this tool is shared across multiple people who each sign their own
    proposals. Returns an id to reference from the generate request."""
    ext = SIGNATURE_EXT_BY_CONTENT_TYPE.get(file.content_type)
    if not ext:
        raise HTTPException(400, "Signature must be a PNG or JPEG image")
    raw = await file.read()
    _validate_is_real_image(raw)
    sig_id = uuid.uuid4().hex
    if blob_store.BLOB_ENABLED:
        blob_store.blob_put(f"temp_signatures/{sig_id}{ext}", raw)
        return {"id": sig_id}
    dest = renderer.TEMP_SIGNATURES_DIR / f"{sig_id}{ext}"
    dest.write_bytes(raw)
    return {"id": sig_id}


@app.get("/companies/{company}/services", dependencies=[Depends(require_auth)])
def list_services(company: str):
    companies = renderer.load_companies()
    if company not in companies:
        raise HTTPException(404, "Unknown company")
    return companies[company]["services"]


@app.get("/companies/{company}/services/{service}/templates", dependencies=[Depends(require_auth)])
def list_templates(company: str, service: str):
    companies = renderer.load_companies()
    if company not in companies or service not in companies[company]["services"]:
        raise HTTPException(404, "Service not offered by this company")
    own_style = companies[company].get("style")
    if not own_style:
        raise HTTPException(404, "This brand doesn't have a working template yet")
    schema = renderer.load_schema(own_style)
    has_preview = renderer.style_preview_path(own_style) is not None
    return [{"style": own_style, "label": schema["label"], "has_preview": has_preview}]


@app.get("/styles/{style}/preview")
def get_style_preview(style: str):
    # Left open (no auth) on purpose: the frontend loads this directly as an <img src=...>,
    # which can't attach an Authorization header. It's just a non-sensitive design thumbnail.
    if style not in renderer.list_styles():
        raise HTTPException(404, "No preview for this style")
    path = renderer.style_preview_path(style)
    if not path:
        raise HTTPException(404, "No preview for this style")
    return FileResponse(path)


@app.get("/companies/{company}/services/{service}/templates/{style}/schema", dependencies=[Depends(require_auth)])
def get_schema(company: str, service: str, style: str):
    # `style` feeds straight into a filesystem path in renderer.load_schema/render --
    # without this check, a value like ".." resolves outside the templates directory
    # instead of cleanly 404ing.
    if style not in renderer.list_styles():
        raise HTTPException(404, "Template not found")
    try:
        schema = renderer.load_schema(style)
    except FileNotFoundError:
        raise HTTPException(404, "Template not found")
    companies = renderer.load_companies()
    if company not in companies:
        raise HTTPException(404, "Unknown company")

    all_fields = schema.get("fields", []) + schema.get("find_replace_fields", [])
    tables = list(schema.get("content_flow", {}).get("tables", []))
    if schema.get("standalone_table"):
        tables.append(schema["standalone_table"])

    fields_out = [
        {"key": f["key"], "label": f["label"], "type": f["type"], "removable": f.get("removable", False)}
        for f in all_fields
    ]
    # "Who's actually sending this" varies by who's using the tool -- this is shared
    # across multiple people/companies, so the rep's name is an editable field (defaulting
    # to the company's usual rep) rather than a value locked to the company forever.
    has_rep_name = "rep_name" in schema.get("branding_targets", {}) or "rep_name" in schema.get("branding_find_replace", {})
    if has_rep_name:
        fields_out.append({
            "key": "rep_name", "label": "Sent By (Rep Name)", "type": "text", "removable": False,
            "default": companies[company].get("rep_name", ""),
        })

    return {
        "fields": fields_out,
        "tables": [
            {"key": t["key"], "label": t["label"], "columns": t["columns"]}
            for t in tables
        ],
        "section_group": schema.get("section_group"),
        "has_rep_signature": bool(schema.get("rep_signature_slot") or schema.get("signature_slot")),
        "has_client_signature": bool(schema.get("client_signature_slot")),
    }


class GenerateRequest(BaseModel):
    values: dict[str, str] = {}
    tables: dict[str, list[list[str]]] = {}
    table_columns: dict[str, list[str]] = {}
    sections: dict[str, list[dict[str, str]]] = {}
    format: str = "pptx"  # "pptx" | "pdf" | "docx"
    rep_signature_id: str | None = None
    client_signature_id: str | None = None


MEDIA_TYPES = {
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


@app.post("/companies/{company}/services/{service}/templates/{style}/generate", dependencies=[Depends(require_auth)])
def generate(company: str, service: str, style: str, req: GenerateRequest):
    if style not in renderer.list_styles():
        raise HTTPException(404, "Template not found")
    if req.format not in ("pptx", "pdf", "docx"):
        raise HTTPException(400, "format must be one of: pptx, pdf, docx")

    try:
        if req.format == "docx":
            result_path = renderer.render_docx(
                company, service, style, req.values, req.tables, req.sections, req.table_columns
            )
        elif req.format == "pdf":
            # Built natively with reportlab rather than converting a pptx via LibreOffice --
            # LibreOffice isn't installable on most free/shared hosts (PythonAnywhere, plain
            # shared cPanel), so this pure-Python path is what makes PDF export portable.
            result_path = renderer.render_pdf_native(
                company, service, style, req.values, req.tables, req.sections, req.table_columns,
                req.rep_signature_id, req.client_signature_id,
            )
        else:
            result_path = renderer.render(
                company, service, style, req.values, req.tables, req.sections, req.table_columns,
                req.rep_signature_id, req.client_signature_id,
            )
    except FileNotFoundError:
        raise HTTPException(404, "Template not found")
    except KeyError as e:
        raise HTTPException(400, f"Unknown company/field: {e}")

    return FileResponse(
        result_path,
        filename=result_path.name,
        media_type=MEDIA_TYPES.get(result_path.suffix, "application/octet-stream"),
    )


if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
