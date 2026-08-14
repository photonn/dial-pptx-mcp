# dial-pptx-mcp

A remote, multi-tenant **PowerPoint-generation MCP server for [EPAM AI DIAL](https://dialx.ai)**.

Generates `.pptx` presentations from corporate templates while preserving the template's theme, layouts, masters, and branding. Designed to be registered as an `MCPToolSet` in [DIAL Quick Apps](https://github.com/epam/ai-dial-quickapps-backend), with file input/output flowing through DIAL Core file storage.

This project extends [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) (MIT) — see [Credits](#credits).

## What this adds over upstream

| Area | Upstream | This project |
|---|---|---|
| Transport | stdio (single local client) | streamable-http / SSE, container-ready (`PPT_MCP_*` env vars) |
| State | process globals, guessable sequential IDs | per-deck UUID handles (unguessable), thread-safe store with TTL + LRU bounds, per-deck locking |
| File I/O | local disk paths | DIAL Files API: template in via Quick Apps `file:data::` references, deck out via `export_presentation` returning a DIAL file URL |
| Deployment | — | Dockerfile (non-root, HTTP defaults) + generic Kubernetes example |

The ~30 upstream content/formatting tools (slides, text, charts, tables, connectors, hyperlinks, masters, transitions) are unchanged — see [docs/UPSTREAM_README.md](docs/UPSTREAM_README.md) for the full tool reference.

## Configuration

All environment-specific settings come from environment variables. Nothing is hardcoded.

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `PPT_MCP_TRANSPORT` | no | `stdio` | `http` (streamable-http, recommended for DIAL), `sse`, or `stdio` |
| `PPT_MCP_HOST` | no | `127.0.0.1` | Bind address; set `0.0.0.0` in containers (the Dockerfile does) |
| `PPT_MCP_PORT` | no | `8000` | Listen port for http/sse |
| `DIAL_CORE_URL` | for DIAL export | — | Base URL of DIAL Core, e.g. `https://dial.example.com`. Unset → DIAL upload/download tools return a clear error; local-path tools still work |
| `DIAL_AUTH_MODE` | no | `auto` | `auto`: credentials from the incoming MCP request first (the end user's bearer, attached by Quick Apps for servers deployed under the DIAL host — exports land in that user's own bucket), falling back to `DIAL_API_KEY`. `caller`: incoming credentials only — fail loudly instead of falling back. `server`: always `DIAL_API_KEY` (single shared bucket) |
| `DIAL_API_KEY` | no | — | Server's own DIAL API key — the fallback identity in `auto` mode, the only identity in `server` mode |
| `DIAL_UPLOAD_FOLDER` | no | `pptx-mcp` | Folder inside the bucket for exported decks |
| `PPT_MCP_STATE_TTL_SECONDS` | no | `3600` | Idle time before an in-memory presentation expires |
| `PPT_MCP_STATE_MAX_PRESENTATIONS` | no | `50` | Max concurrently held presentations (LRU eviction) |
| `PPT_TEMPLATE_PATH` | no | — | Extra local directories searched by the local-path template tools (`:`-separated) |
| `VISION_LLM_MODEL` | for visual QA | — | Vision model: the model name (direct endpoint) or the DIAL deployment name (DIAL provider); must accept image input |
| `VISION_LLM_ENDPOINT` | direct provider | — | OpenAI Responses-API endpoint, e.g. `https://<resource>.openai.azure.com/openai/responses?api-version=2025-04-01-preview`. When unset, the model is called through DIAL Core instead: `{DIAL_CORE_URL}/openai/deployments/{model}/chat/completions` with DIAL credentials (caller headers first, `DIAL_API_KEY` fallback) |
| `VISION_LLM_API_KEY` | direct provider | — | Key for the direct endpoint (sent as `api-key` and `Authorization: Bearer`) |
| `VISION_LLM_PROVIDER` | no | auto | Force the backend: `direct` or `dial` (default: `direct` when `VISION_LLM_ENDPOINT` is set, else `dial`) |
| `VISION_LLM_API_VERSION` | no | — | Optional `?api-version=` for the DIAL-routed call |
| `VISION_LLM_MAX_SLIDES` | no | `15` | Cap on slides sent per inspection |
| `SOFFICE_PATH` | no | `soffice` on PATH | LibreOffice binary used to render slides (the Docker image includes LibreOffice) |

Copy [`.env.example`](.env.example) to `.env` for local runs — the server loads it at startup, and `.env` is gitignored.

## Running

Local (stdio, for MCP-client desktop use):

```bash
pip install -r requirements.txt
python ppt_mcp_server.py
```

Remote (streamable-http):

```bash
docker build -t dial-pptx-mcp .
docker run -p 8000:8000 -e DIAL_CORE_URL=https://dial.example.com -e DIAL_API_KEY=... dial-pptx-mcp
# MCP endpoint: http://<host>:8000/mcp
```

Kubernetes: see [deploy/kubernetes.yaml](deploy/kubernetes.yaml) (generic example — replace placeholders).

## DIAL Quick Apps integration

Register the deployed server as an MCP tool set in your Quick App manifest:

```json
{
  "name": "powerpoint",
  "description": "Generate PowerPoint presentations from corporate templates",
  "type": "mcp",
  "mcp_server_info": {
    "url": "https://YOUR-DEPLOYED-HOST/mcp",
    "protocol": "streamable_http",
    "authorization": null
  },
  "attachment": {
    "supported_types": ["*/*"],
    "propagate_types_to_choice": [
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ]
  }
}
```

- `propagate_types_to_choice` makes the exported `.pptx` attachment visible to the end user in DIAL Chat (tool-call results are hidden by default).
- **Deploy the server under the DIAL host** (behind DIAL Core routing) for per-user storage. Quick Apps attaches the end user's `Authorization: Bearer` only to MCP servers whose URL starts with the DIAL host — with it, the default `auto` mode uploads every export to that user's own bucket. For a server at an external URL, Quick Apps sends no user credentials (it deliberately refuses to forward `api-key`/`authorization` as custom headers), so `auto` falls back to the server's `DIAL_API_KEY` and exports land in the server's single bucket; set `DIAL_AUTH_MODE=caller` if you'd rather exports fail loudly than fall back.
- **Template input**: the orchestrating agent passes the template to `create_presentation_from_template_content` as `file:data::files/{bucket}/{path}` — Quick Apps' file preprocessing resolves that reference to a data: URI before this server receives it (base64 via `file:base64::` also accepted). Note Quick Apps' default 10 MiB file-loading limit (`features.file_loading.size_limit`) if your templates are large.
- **Deck output**: `export_presentation` uploads to DIAL file storage and returns the `files/{bucket}/{path}` URL; the tool description instructs the agent to include it in its final answer.

## Automatic visual QA (internal inspect-and-repair loop)

When a vision LLM is configured (`VISION_LLM_*`), quality assurance is **entirely internal to the server** — the calling agent only ever receives a finished, verified presentation. On every `export_presentation`/`save_presentation` of a deck that was created or edited since it last passed:

1. All slides are rendered (LibreOffice → PDF → PNG) and reviewed by the vision LLM for template/brand fidelity and visible errors (overflowing or clipped text, overlaps, unfilled placeholders, broken charts, illegibility).
2. If issues are found, the server **repairs the deck itself**: the LLM is shown the issues, the affected slides' structure, and their images, and returns a plan of whitelisted operations (move/resize shape, set font size, set text, word wrap, delete shape) that are validated and applied with python-pptx.
3. The deck is re-rendered and re-inspected; the loop repeats up to `VISUAL_QA_MAX_ITERATIONS` (default 10) inspections, stopping early if no repair makes progress.
4. Only a deck that passes is exported. `VISUAL_QA_ON_UNRESOLVED` controls what happens if the loop cannot reach a pass: `report` (default) fails the export with the unresolved issue list — a genuine failure report, not a retry request — while `export_as_is` ships the best-effort deck anyway.

Passed decks aren't re-inspected unless edited again, and exports skip QA entirely when the feature is unconfigured. `VISUAL_QA_ENFORCE=false` disables it; `VISUAL_QA_ON_ERROR=allow` lets exports through when inspection itself cannot run (renderer/LLM outage — default blocks); `VISUAL_QA_EXPOSE_TOOL=true` additionally exposes a standalone `visual_inspect_presentation` tool for debugging. The reviewer model can be reached two ways: a direct OpenAI Responses-API endpoint with image input (Azure OpenAI included), or as a DIAL Core deployment via `{DIAL_CORE_URL}/openai/deployments/{model}/chat/completions` — see the `VISION_LLM_*` variables. Cost note: each loop iteration is one render plus one or two LLM calls, so a worst-case export adds a few minutes and a handful of vision-model requests.

## Multi-tenancy and scaling notes

- Presentation handles are server-generated UUIDs and act as unguessable capabilities; clients cannot enumerate or guess other conversations' decks.
- Presentation state lives in process memory (bounded by TTL + LRU). Run a single replica, or use session affinity if you scale out — a deck created on one replica is not visible on another.
- Calls targeting the same presentation are serialized (python-pptx is not thread-safe); different presentations are handled concurrently.

## Development

```bash
uv venv --python 3.12 .venv && uv pip install -r requirements.txt
.venv/bin/python -m unittest discover -s tests        # unit tests
.venv/bin/python spike/fidelity_spike.py <template>   # template-fidelity check
.venv/bin/python spike/http_client_check.py           # transport smoke test
```

`spike/fidelity_spike.py` creates a deck from a template through the same code paths the MCP tools use and byte-compares the theme, slide-master, layout, and media parts of template vs output. Local test templates belong in `templates-local/` (gitignored).

## Credits

This project is built on **[Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server)** by [GongRzhe](https://github.com/GongRzhe), used under the MIT license, with full git history preserved. The core PowerPoint manipulation tools and utilities are upstream work; this fork adds the remote-transport, multi-tenant state, and DIAL integration layers. The original [LICENSE](LICENSE) and copyright notice are retained; the upstream README is preserved at [docs/UPSTREAM_README.md](docs/UPSTREAM_README.md).

## License

MIT — see [LICENSE](LICENSE).
