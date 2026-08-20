"""
DIAL Core file-storage bridge (workstream 5.4).

Minimal client for the DIAL Core Files API, used by the template-input and
export tools so presentations flow through DIAL file storage instead of the
server's local disk.

Endpoints (DIAL Core OpenAPI, docs/open_api_core.yaml in epam/ai-dial-core):
- GET /v1/bucket                       -> {"bucket": "..."} for the caller
- PUT /v1/files/{bucket}/{path}        -> multipart upload, field name "file";
                                          response FileMetadata includes
                                          "url": "files/{bucket}/{path}"
- GET /v1/files/{bucket}/{path}        -> file bytes

Authentication: "Api-Key: <key>" header, or "Authorization: Bearer <JWT>".

Credential resolution, per tool call:
1. Headers forwarded by the caller on the incoming MCP HTTP request
   (Api-Key or Authorization) — so files land in the *caller's* bucket when
   DIAL Quick Apps is configured to forward auth to this toolset.
2. Otherwise the server's own DIAL_API_KEY env var.

Environment variables (all optional unless DIAL upload/download is used):
- DIAL_CORE_URL       Base URL of DIAL Core, e.g. https://dial.example.com
- DIAL_API_KEY        Fallback API key when the caller forwards no credentials
- DIAL_UPLOAD_FOLDER  Folder inside the bucket for exports (default: pptx-mcp)
- DIAL_PUBLIC_URL     Extra host(s) that DIAL file links may carry besides
                      DIAL_CORE_URL (comma-separated URLs or hostnames), for
                      installations reached in-cluster but linked publicly
"""
import io
import os
import uuid
from urllib.parse import urlsplit

import httpx

from logging_utils import get_logger

logger = get_logger("dial_client")

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


class DialConfigError(RuntimeError):
    pass


def _incoming_request_headers():
    """Headers of the MCP HTTP request currently being served, or {}.

    Uses the pinned (<2.0) SDK's request contextvar; degrades to {} for stdio
    transport or if SDK internals change.
    """
    try:
        from mcp.server.lowlevel.server import request_ctx
        request = request_ctx.get().request
        return dict(request.headers) if request is not None else {}
    except Exception:
        logger.debug("incoming_headers_unavailable transport=stdio_or_sdk_change")
        return {}


def resolve_dial_auth_headers(api_key=None):
    """Resolve DIAL credentials per DIAL_AUTH_MODE (shared by the Files API
    client and the DIAL-routed vision LLM):

    - "auto" (default): credentials on the incoming MCP request first
      (the user's bearer, which DIAL Quick Apps attaches when this server
      is deployed behind the DIAL host — operations then run as the
      caller), falling back to DIAL_API_KEY.
    - "caller": incoming request credentials ONLY; error out rather than
      fall back to a shared identity.
    - "server": always DIAL_API_KEY.
    """
    mode = os.environ.get("DIAL_AUTH_MODE", "auto").lower()
    incoming = _incoming_request_headers()
    if mode != "server":
        for name in ("api-key", "authorization"):
            if name in incoming:
                logger.debug("dial_auth_resolved mode=%s identity=caller header=%s",
                             mode, name)
                return {name: incoming[name]}
        if mode == "caller":
            raise DialConfigError(
                "No caller credentials on this request. DIAL_AUTH_MODE="
                "caller requires the user's Api-Key/Authorization header, "
                "which DIAL Quick Apps forwards only to MCP servers "
                "deployed under the DIAL host — deploy this server behind "
                "DIAL Core routing, or set DIAL_AUTH_MODE=auto/server to "
                "allow the shared DIAL_API_KEY."
            )
    api_key = api_key or os.environ.get("DIAL_API_KEY")
    if api_key:
        if mode == "auto" and incoming:
            logger.warning("dial_auth_fallback mode=auto identity=server "
                           "reason=no_caller_credentials")
        else:
            logger.debug("dial_auth_resolved mode=%s identity=server", mode)
        return {"Api-Key": api_key}
    raise DialConfigError(
        "No DIAL credentials: the caller forwarded no Api-Key/Authorization "
        "header and DIAL_API_KEY is not set on the server."
    )


def _extra_file_hosts():
    """Hosts a DIAL file link may legitimately carry besides DIAL_CORE_URL.

    A deployment commonly reaches DIAL Core by its in-cluster URL while the
    links the orchestrator holds carry the public DIAL Chat hostname.
    DIAL_PUBLIC_URL lists those aliases; the bytes are still always fetched
    from DIAL_CORE_URL — only the path is taken from the link.
    """
    hosts = set()
    for entry in os.environ.get("DIAL_PUBLIC_URL", "").split(","):
        entry = entry.strip()
        if entry:
            hosts.add(urlsplit(entry).netloc or entry)
    return hosts


class DialFileClient:
    def __init__(self, base_url=None, api_key=None, timeout=60.0):
        base_url = base_url or os.environ.get("DIAL_CORE_URL")
        if not base_url:
            raise DialConfigError(
                "DIAL file storage is not configured on this server: "
                "set the DIAL_CORE_URL environment variable."
            )
        self._base_url = base_url.rstrip("/")
        self._api_key = api_key or os.environ.get("DIAL_API_KEY")
        self._timeout = timeout

    def _auth_headers(self):
        return resolve_dial_auth_headers(self._api_key)

    def get_bucket_info(self):
        """GET /v1/bucket for the resolved credentials. Returns the full
        response: always {"bucket": ...}; in an application context DIAL
        Core adds {"appdata": "{user-bucket}/appdata/{deployment}"} — a
        path inside the END USER's bucket that the app may write and the
        user may read."""
        headers = self._auth_headers()
        r = httpx.get(f"{self._base_url}/v1/bucket", headers=headers,
                      timeout=self._timeout)
        r.raise_for_status()
        return r.json()

    def get_bucket(self):
        return self.get_bucket_info()["bucket"]

    def upload(self, data: bytes, filename: str, folder: str = None,
               content_type: str = PPTX_MIME):
        """Upload bytes to user-accessible DIAL storage; returns the
        DIAL-relative file URL.

        Uses the appdata path when DIAL Core provides one (application
        context: the file lands in the end user's bucket, so the user can
        download the attachment — the dall-e-3 pattern); otherwise the
        caller's own bucket (direct user context)."""
        headers = self._auth_headers()
        info = self.get_bucket_info()
        base = info.get("appdata") or info["bucket"]
        folder = folder or os.environ.get("DIAL_UPLOAD_FOLDER", "pptx-mcp")
        # Unique path segment so concurrent exports never collide/overwrite.
        path = f"{base}/{folder}/{uuid.uuid4().hex}/{filename}"
        logger.debug("dial_upload_start filename=%s bytes=%d target=%s",
                     filename, len(data), "appdata" if info.get("appdata")
                     else "bucket")
        r = httpx.put(
            f"{self._base_url}/v1/files/{path}",
            headers=headers,
            files={"file": (filename, io.BytesIO(data), content_type)},
            timeout=self._timeout,
        )
        if r.status_code >= 400:
            logger.error("dial_upload_failed filename=%s bytes=%d status=%d",
                         filename, len(data), r.status_code)
        r.raise_for_status()
        meta = r.json()
        url = meta.get("url", f"files/{path}")
        logger.info("dial_upload_ok filename=%s bytes=%d url=%s",
                    filename, len(data), url)
        return url

    def _file_request_url(self, file_url):
        """Map a DIAL file reference to this server's Files API URL.

        Accepts the relative form the upload returns (files/{bucket}/{path}),
        the Core API path (v1/files/...) and the chat frontend's proxy path
        (api/files/...), absolute or not. Anything else — in particular an
        arbitrary web URL — is refused: this server is not a web fetcher.
        """
        ref = (file_url or "").strip()
        if ref.startswith(("http://", "https://")):
            parts = urlsplit(ref)
            allowed = {urlsplit(self._base_url).netloc} | _extra_file_hosts()
            if parts.netloc not in allowed:
                logger.warning("dial_download_refused host=%s reason="
                               "not_dial_core", parts.netloc)
                raise DialConfigError(
                    f"Refusing to download from {parts.netloc}: this server "
                    "reads files from DIAL file storage only, never arbitrary "
                    "web URLs. Save the file to DIAL file storage first and "
                    "pass the 'files/{bucket}/{path}' URL that upload "
                    f"returned. (If {parts.netloc} is this DIAL installation "
                    "under another name, the operator can list it in "
                    "DIAL_PUBLIC_URL.)"
                )
            ref = parts.path
        ref = ref.lstrip("/")
        # The same object is linked as v1/files/... by the Core API and as
        # api/files/... by the chat frontend's proxy.
        for prefix in ("v1/", "api/"):
            if ref.startswith(prefix):
                ref = ref[len(prefix):]
                break
        if not ref.startswith("files/"):
            raise DialConfigError(
                f"'{file_url}' is not a DIAL file reference. Expected "
                "files/{bucket}/{path} — the URL returned when the file was "
                "uploaded to DIAL file storage."
            )
        return f"{self._base_url}/v1/{ref}"

    def download(self, file_url: str) -> bytes:
        """Download a DIAL file by its relative URL (files/{bucket}/{path}),
        or by an absolute URL on this DIAL installation."""
        # Validate the reference before resolving credentials, so a bad URL
        # is reported as a bad URL rather than as a credentials problem.
        url = self._file_request_url(file_url)
        headers = self._auth_headers()
        r = httpx.get(url, headers=headers, timeout=self._timeout)
        if r.status_code >= 400:
            logger.error("dial_download_failed url=%s status=%d",
                         file_url, r.status_code)
        r.raise_for_status()
        logger.info("dial_download_ok url=%s bytes=%d", file_url, len(r.content))
        return r.content
