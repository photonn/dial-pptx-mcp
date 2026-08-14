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
| `DIAL_API_KEY` | no | — | Fallback DIAL API key when the caller forwards no credentials. If Quick Apps forwards the user's `Api-Key`/`Authorization` header to this toolset, that identity is used instead and exports land in the caller's own bucket |
| `DIAL_UPLOAD_FOLDER` | no | `pptx-mcp` | Folder inside the bucket for exported decks |
| `PPT_MCP_STATE_TTL_SECONDS` | no | `3600` | Idle time before an in-memory presentation expires |
| `PPT_MCP_STATE_MAX_PRESENTATIONS` | no | `50` | Max concurrently held presentations (LRU eviction) |
| `PPT_TEMPLATE_PATH` | no | — | Extra local directories searched by the local-path template tools (`:`-separated) |
| `VISION_LLM_ENDPOINT` | for visual QA | — | OpenAI Responses-API endpoint of a vision model, e.g. `https://<resource>.openai.azure.com/openai/responses?api-version=2025-04-01-preview`. Unset → `visual_inspect_presentation` returns a clear error |
| `VISION_LLM_API_KEY` | for visual QA | — | Key for that endpoint (sent as `api-key` and `Authorization: Bearer`) |
| `VISION_LLM_MODEL` | for visual QA | — | Model / Azure deployment name; must accept image input |
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
- **Template input**: the orchestrating agent passes the template to `create_presentation_from_template_content` as `file:data::files/{bucket}/{path}` — Quick Apps' file preprocessing resolves that reference to a data: URI before this server receives it (base64 via `file:base64::` also accepted). Note Quick Apps' default 10 MiB file-loading limit (`features.file_loading.size_limit`) if your templates are large.
- **Deck output**: `export_presentation` uploads to DIAL file storage and returns the `files/{bucket}/{path}` URL; the tool description instructs the agent to include it in its final answer.

## Automatic visual QA (export gate)

When a vision LLM is configured (`VISION_LLM_*`), visual inspection is **enforced by the server, not offered as a tool**: every `export_presentation`/`save_presentation` call on a deck that was created or edited since its last passed inspection first renders all slides (LibreOffice → PDF → PNG) and has the vision LLM review them for template/brand fidelity and visible errors (overflowing or clipped text, overlaps, unfilled placeholders, broken charts, illegibility). A failing deck is **refused** with `{issues: [{slide, severity, description, suggested_fix}]}` and a retry instruction — so the orchestrating agent is forced into a fix-and-re-export loop until the deck passes. Passed decks aren't re-inspected unless edited again, and exports skip the gate entirely when the feature is unconfigured.

Related settings: `VISUAL_QA_ENFORCE=false` disables the gate; `VISUAL_QA_ON_ERROR=allow` lets exports through when inspection itself cannot run (renderer/LLM outage — default blocks); `VISUAL_QA_EXPOSE_TOOL=true` additionally exposes a standalone `visual_inspect_presentation` tool for mid-build checks. Any endpoint speaking the OpenAI Responses API with image input works (Azure OpenAI included).

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
