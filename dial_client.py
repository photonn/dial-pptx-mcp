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
"""
import io
import os
import uuid

import httpx

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
        return {}


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
        """Prefer credentials forwarded on the incoming request; fall back to
        the server's own API key."""
        incoming = _incoming_request_headers()
        for name in ("api-key", "authorization"):
            if name in incoming:
                return {name: incoming[name]}
        if self._api_key:
            return {"Api-Key": self._api_key}
        raise DialConfigError(
            "No DIAL credentials: the caller forwarded no Api-Key/Authorization "
            "header and DIAL_API_KEY is not set on the server."
        )

    def get_bucket(self):
        headers = self._auth_headers()
        r = httpx.get(f"{self._base_url}/v1/bucket", headers=headers,
                      timeout=self._timeout)
        r.raise_for_status()
        return r.json()["bucket"]

    def upload(self, data: bytes, filename: str, folder: str = None,
               content_type: str = PPTX_MIME):
        """Upload bytes to the caller's bucket; returns the DIAL-relative
        file URL, e.g. files/{bucket}/{folder}/{unique}/{filename}."""
        headers = self._auth_headers()
        bucket = self.get_bucket()
        folder = folder or os.environ.get("DIAL_UPLOAD_FOLDER", "pptx-mcp")
        # Unique path segment so concurrent exports never collide/overwrite.
        path = f"{folder}/{uuid.uuid4().hex}/{filename}"
        r = httpx.put(
            f"{self._base_url}/v1/files/{bucket}/{path}",
            headers=headers,
            files={"file": (filename, io.BytesIO(data), content_type)},
            timeout=self._timeout,
        )
        r.raise_for_status()
        meta = r.json()
        return meta.get("url", f"files/{bucket}/{path}")

    def download(self, file_url: str) -> bytes:
        """Download a DIAL file by its relative URL (files/{bucket}/{path})
        or absolute URL under DIAL_CORE_URL."""
        headers = self._auth_headers()
        if file_url.startswith(("http://", "https://")):
            if not file_url.startswith(self._base_url + "/"):
                raise DialConfigError(
                    "Refusing to download from a host other than DIAL_CORE_URL."
                )
            url = f"{self._base_url}/v1/files/{file_url[len(self._base_url) + 1:]}" \
                if not file_url.startswith(f"{self._base_url}/v1/files/") else file_url
        else:
            url = f"{self._base_url}/v1/{file_url.lstrip('/')}"
        r = httpx.get(url, headers=headers, timeout=self._timeout)
        r.raise_for_status()
        return r.content
