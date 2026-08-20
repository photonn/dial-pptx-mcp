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
| Images | local path or base64 | `add_image_from_dial_url` fetches orchestrator-generated images from DIAL storage server-side, with aspect-ratio-aware placement |
| Deployment | — | Dockerfile (non-root, HTTP defaults) + generic Kubernetes example |

The ~30 upstream content/formatting tools (slides, text, charts, tables, connectors, hyperlinks, masters, transitions) are unchanged — see [docs/UPSTREAM_README.md](docs/UPSTREAM_README.md) for the full tool reference.

## Configuration

All environment-specific settings come from environment variables. Nothing is hardcoded.

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `PPT_MCP_TRANSPORT` | no | `stdio` | `http` (streamable-http, recommended for DIAL), `sse`, or `stdio` |
| `PPT_MCP_HOST` | no | `127.0.0.1` | Bind address; set `0.0.0.0` in containers (the Dockerfile does) |
| `PPT_MCP_PORT` | no | `8000` | Listen port for http/sse |
| `PPT_MCP_ALLOWED_HOSTS` | no | `dial-pptx-mcp.dial.svc.cluster.local` | Comma-separated Host-header allowlist for http/sse. Only listed hosts (plus localhost) are accepted; others get `421`. The default matches the standard in-cluster service name — override with your own hostname(s), or set `*` to disable Host checking entirely (e.g. behind ingress under other names). Loopback binds keep the SDK's localhost-only protection |
| `DIAL_CORE_URL` | for DIAL export | — | Base URL of DIAL Core, e.g. `https://dial.example.com`. Unset → DIAL upload/download tools return a clear error; local-path tools still work |
| `DIAL_AUTH_MODE` | no | `auto` | `auto`: credentials from the incoming MCP request first (the end user's bearer, attached by Quick Apps for servers deployed under the DIAL host — exports land in that user's own bucket), falling back to `DIAL_API_KEY`. `caller`: incoming credentials only — fail loudly instead of falling back. `server`: always `DIAL_API_KEY` (single shared bucket) |
| `DIAL_API_KEY` | no | — | Server's own DIAL API key — the fallback identity in `auto` mode, the only identity in `server` mode |
| `DIAL_UPLOAD_FOLDER` | no | `pptx-mcp` | Folder inside the bucket for exported decks |
| `DIAL_IMAGE_MAX_MB` | no | `20` | Largest image `add_image_from_dial_url` will download and embed. An unparsable value falls back to the default |
| `PPT_MCP_STATE_TTL_SECONDS` | no | `3600` | Idle time before an in-memory presentation expires |
| `PPT_MCP_STATE_MAX_PRESENTATIONS` | no | `50` | Max concurrently held presentations (LRU eviction) |
| `PPT_TEMPLATE_PATH` | no | — | Extra local directories searched by the local-path template tools (`:`-separated) |
| `VISION_LLM_MODEL` | for visual QA | — | Vision model: the model name (direct endpoint) or the DIAL deployment name (DIAL provider); must accept image input |
| `VISION_LLM_ENDPOINT` | direct provider | — | OpenAI Responses-API endpoint, e.g. `https://<resource>.openai.azure.com/openai/responses?api-version=2025-04-01-preview`. When unset, the model is called through DIAL Core instead: `{DIAL_CORE_URL}/openai/deployments/{model}/chat/completions` with DIAL credentials (caller headers first, `DIAL_API_KEY` fallback) |
| `VISION_LLM_API_KEY` | direct provider | — | Key for the direct endpoint (sent as `api-key` and `Authorization: Bearer`) |
| `VISION_LLM_PROVIDER` | no | auto | Force the backend: `direct` or `dial` (default: `direct` when `VISION_LLM_ENDPOINT` is set, else `dial`) |
| `VISION_LLM_API_VERSION` | no | `2025-04-01-preview` | `?api-version=` added to the vision call when the endpoint URL doesn't already carry one. Azure OpenAI (and DIAL Core's Azure upstream) reject requests without it — `api-version is a required query parameter`. The default covers both the Responses API and chat completions with image input; an `api-version` already present in `VISION_LLM_ENDPOINT` always wins |
| `VISION_LLM_MAX_SLIDES` | no | `15` | Cap on slides sent per whole-deck inspection (an explicit `slides` list is never capped) |
| `VISUAL_QA_ENFORCE` | no | `true` | `false` unregisters the visual QA tools entirely |
| `VISUAL_QA_MAX_ITERATIONS` | no | `10` | Inspect/repair rounds per `visual_repair_slides` call (overridable per call) |
| `VISUAL_QA_EXPORT_GATE` | no | `false` | `true` also runs a whole-deck inspect-repair loop inside export/save and refuses unverified decks |
| `VISUAL_QA_ON_UNRESOLVED` | no | `report` | Export gate only: `report` fails the export with the issue list, `export_as_is` ships the deck |
| `VISUAL_QA_ON_ERROR` | no | `block` | Export gate only: `allow` exports when inspection itself cannot run |
| `SOFFICE_PATH` | no | `soffice` on PATH | LibreOffice binary used to render slides (the Docker image includes LibreOffice) |
| `LOG_LEVEL` | no | `INFO` | `DEBUG`, `INFO`, `WARNING`, `ERROR` or `CRITICAL`. One log line per event on stderr — see [Logging](#logging). An unrecognized value falls back to `INFO` rather than failing startup |

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

## Images (orchestrator-generated)

The server does not generate images — it embeds them. The split is: **the orchestrator generates, the MCP inserts**, and only a short URL travels between the two.

| Tool | Use it for |
|---|---|
| `add_image_from_dial_url(presentation_id, slide_index, image_url, left?, top?, width?, height?, fit?)` | Anything the orchestrator produced with an image model. It calls the image deployment (a DIAL Core deployment sits next to the vision one), saves the result to DIAL file storage, and passes the `files/{bucket}/{path}` URL here; the server downloads the bytes itself with the caller's own DIAL credentials (same `DIAL_AUTH_MODE` resolution as export) |
| `manage_image(..., source_type="base64")` | Small assets only, and deployments not running under the DIAL host |

Prefer the URL tool. A 1024×1024 PNG is ~1–2 MB, so passing it as base64 pushes ~2 MB of payload through the agent's context on every insertion — enough to wreck the iteration budget on a multi-image deck.

**Placement is aspect-ratio aware.** `slide_index` is 0-based (like the other content tools; visual QA slide numbers are 1-based). Give `width` **and** `height` to define the box the picture should occupy, and `fit` decides how it relates to that box:

- `contain` (default) — largest undistorted size that fits, centred in the box. Safe for photos and illustrations.
- `cover` — fills the box exactly, cropping the overflowing edges symmetrically (python-pptx crop, no re-encoding).
- `stretch` — forces the exact box, distorting the image. Visual QA can move, resize and delete a picture, but it cannot un-distort one, so avoid `stretch` unless you mean it.

Pass only one of `width`/`height` to scale proportionally, or neither to keep the image's natural size — clamped to the slide, so a large generated PNG never hangs off the edge. The response reports the geometry actually applied (`"placed"`), which under `contain` may be smaller than the box you asked for; use it to lay out the text beside the image. A half-and-half slide on a 13.33in deck is text at `left=0.8, width=5.6` and the picture at `left=6.9, top=1.2, width=5.6, height=4.5`.

**Telling the agent to use it.** Image generation is the orchestrator's job, so it belongs in the Quick App's system prompt:

```text
When a slide would be stronger with a visual — a supporting image beside the
text, a cover image, an icon — generate it with the image model, upload it to
DIAL file storage, and pass the returned files/... URL to
add_image_from_dial_url. Never paste image data into the conversation. Give
width and height for the box you want it to fill and leave fit at "contain"
so the image is not distorted; for a text-left/image-right slide use roughly
half the slide width for each.
```

Generated images are in scope for `visual_repair_slides` like any other shape. `DIAL_IMAGE_MAX_MB` (default 20) bounds what the server will download; non-raster input is refused with a message telling the agent to ask its image model for PNG or JPEG rather than SVG.

## Visual QA (agent-driven inspect and repair)

When a vision LLM is configured (`VISION_LLM_*`), the server registers two tools the orchestrating agent calls whenever it wants — typically right after building each slide, not only at the end:

| Tool | What it does |
|---|---|
| `visual_inspect_slides(presentation_id, slides?, focus?, reference_presentation_id?)` | Renders the selected slides (LibreOffice → PDF → PNG) and has the vision LLM review them for template/brand fidelity and text placement problems (see below). Read-only: returns `{"passed", "issues": [{slide, severity, description, suggested_fix}]}` |
| `visual_repair_slides(presentation_id, slides?, focus?, max_iterations?)` | Inspects, then **repairs the deck itself** and re-inspects, looping until the slides pass or the budget runs out. The LLM is shown the issues, the affected slides' structure and their images, and returns a plan of whitelisted operations (move/resize shape, set/fit font size, autofit, set text, word wrap, delete shape, table column width/row height/cell text, chart legend and data labels) that are validated and applied with python-pptx |

`slides` is a list of 1-based slide numbers; omit it to work on the whole deck. Issue slide numbers are always absolute deck positions, even when only a subset was rendered, and a scoped repair call never touches a slide outside `slides`. Because LibreOffice converts the whole deck either way, a narrow selection saves the vision call and the repair round, not the render.

`max_iterations` defaults to `VISUAL_QA_MAX_ITERATIONS` (10) and can be lowered per call for a quick single-slide pass. A `"passed": false` result is a report, not a retry request: the agent should edit the content itself and inspect again, or tell the user what remains.

### Export

`export_presentation` does **not** run QA. It reports what it knows — `"visual_qa": "passed" | "unverified" | "unavailable"` — and adds a note when the deck was never inspected or was edited since its last passing inspection. Only a clean whole-deck inspection marks a deck `passed`; a scoped call clears nothing.

Operators who want the old guarantee that no unverified deck ever leaves the server set `VISUAL_QA_EXPORT_GATE=true`: export/save then run the whole-deck inspect-repair loop for dirty decks and refuse the export if it cannot reach a pass. With the gate on, `VISUAL_QA_ON_UNRESOLVED` chooses `report` (default — fail the export with the unresolved issue list) or `export_as_is`, and `VISUAL_QA_ON_ERROR=allow` lets exports through when inspection itself cannot run (renderer/LLM outage — default blocks). Both variables are inert while the gate is off.

### What the reviewer checks

Text is not only in text boxes, so neither is the review. Besides brand fidelity (colors, fonts, logo placement, layout usage) the reviewer is asked to judge **text placement and overlap wherever text is rendered**:

- **Text boxes and placeholders** — overflowing, clipped, or spilling past the slide edge; text overlapping other text or sitting unreadably on top of shapes and images; unfilled placeholders; text too small or too low-contrast to read.
- **Charts and graphs** — axis tick labels colliding with each other or truncated, data labels overlapping their bars/slices or each other, a legend covering the plot area, an axis title rotated into illegibility.
- **Tables** — cell text wrapping into an unreadable stack or clipped by the row height, columns too narrow for their content, headers misaligned with their columns, a table running past the slide.
- **Diagrams, SmartArt and grouped shapes** — labels wider than the node that holds them, text escaping a connector, node labels overlapping their neighbours.

It also flags text sized badly for the space it occupies — a heading set so small its box is mostly empty, or comparable elements at visibly different sizes — while being told not to ask for bigger text where growing it would eat the slide's white space.

Overlapping or unreadable text is graded at least `major`, so it fails the verdict rather than being noted in passing.

**Fitting text to its box.** The `fit_text` operation sizes text to the space it actually has, in both directions: it shrinks text that overflows and grows text that leaves its box mostly empty. The size is computed server-side from the box geometry (minus the frame's own margins, with a slack factor so text never touches its border) rather than guessed by the model, and the plan can bound it with `min_pt`/`max_pt`. Growth is anchored to the deck's own typography — at most 1.5× the shape's current size, or 44pt when the text inherits its size from the layout — so a two-word box cannot balloon to 96pt and shout over the slide. `set_autofit` sets PowerPoint's own autofit behaviour (`shrink_text`, `grow_shape`, `none`) when that suits the shape better. The size estimate is geometric, not a real text layout; the loop's re-render and re-review is what confirms it.

The repair engine can act on all of it: `describe_slides` hands the planner each table's column widths, row heights and cell text, and each chart's type, categories, series count and label/legend state — so a plan can widen a column, raise a row, retitle a cell, shrink a whole table's or chart's font, hide crowded data labels, or move the legend, instead of only nudging the container. Members of a group are not individually addressable; the group is moved, resized or shrunk as a whole.

### Telling the agent to use it

Nothing forces the orchestrator to inspect: with the export gate off, `export_presentation` reports `"visual_qa": "unverified"` but still succeeds. Put the workflow in the Quick App's system prompt so QA actually happens:

```text
After you finish building each slide, call visual_inspect_slides with that
slide's number. If it reports issues, call visual_repair_slides for the same
slide and continue only once it passes or you have fixed the content yourself.
Before export_presentation, call visual_inspect_slides once with no slides
argument to check the deck as a whole. If the export response says
"visual_qa": "unverified", say so in your answer rather than presenting the
deck as checked.
```

Per-slide checks are the cheap path — one render plus one vision call each, caught while the slide is still fresh in context. Keep the whole-deck pass for the end: it is the only thing that marks the deck `passed`, and it catches cross-slide inconsistencies a single-slide review cannot see.

### Sizing the QA work (orchestrator budget, timeouts, pod resources)

| Concern | Guidance |
|---|---|
| Orchestrator iterations (Quick Apps `max_iterations`, default 15) | Now includes the QA calls the agent makes. Roughly 2 calls per slide plus create/export, plus one inspect or repair per slide: a 20-slide deck needs **~65**, so set `max_iterations` to **80** (100 if slides carry charts/tables/images) |
| Tool timeout (Quick Apps `tool_defaults.timeout_seconds`, default 300s) | A single-slide inspect ≈ 15–30s (render + review); a single-slide repair round adds another LLM call. A whole-deck `visual_repair_slides` on 20 slides is the expensive case at ≈ 40–90s per round — budget `max_iterations × 90s` for it, or keep calls slide-scoped and 300s is plenty |
| Slides actually reviewed | `VISION_LLM_MAX_SLIDES` (default 15) caps whole-deck calls only; an explicit `slides` list is never truncated |
| Pod resources | LibreOffice renders in-pod: budget **1 CPU / 2Gi** with a writable `/tmp`. Small limits (e.g. 192Mi) get the renderer OOM-killed, which fails every QA call |

`VISUAL_QA_ENFORCE=false` registers neither tool and turns the feature off. The reviewer model can be reached two ways: a direct OpenAI Responses-API endpoint with image input (Azure OpenAI included), or as a DIAL Core deployment via `{DIAL_CORE_URL}/openai/deployments/{model}/chat/completions` — see the `VISION_LLM_*` variables. Cost note: each inspect is one render plus one LLM call; each repair round adds a second LLM call.

## Logging

Every record is a **single line on stderr**, so `kubectl logs` shows one event per line and nothing wraps across lines:

```
2026-08-20T09:14:02.517Z INFO  dial_pptx.tools.presentation export_ok presentation_id=9f3c1a2b… filename=deck.pptx slides=12 bytes=1841203
```

`timestamp (UTC) · level · logger · message`, with details as `key=value` pairs. Multi-line content is folded onto the same line with ` | ` separators — that includes tracebacks, so a stack trace stays greppable instead of scrolling the pod terminal. FastMCP's default Rich handler (boxed, multi-line, colored) and uvicorn's separate log format are both replaced, so third-party output matches.

`LOG_LEVEL` sets the verbosity of this server and the libraries under it:

| Level | What you get |
|---|---|
| `ERROR` | Failures only: blocked exports, render/vision-LLM outages, upload failures |
| `WARNING` | The above, plus degradations that don't fail the call: QA rounds that found issues, LRU eviction, auth falling back to the server key, skipped repair operations |
| `INFO` (default) | One line per tool call (`tool_ok`/`tool_error` with `duration_ms`), presentation lifecycle, each QA round's verdict, exports and uploads, server startup |
| `DEBUG` | The above, plus per-call argument summaries, render and vision-LLM timings, individual QA issues and repair operations, template content-type coercion, and the underlying HTTP client's own logs |

Handles are truncated (`9f3c1a2b…`) and argument values are summarized by type and size (`template_content=<str:412880>`) rather than logged verbatim, so logs never carry a usable presentation handle, a base64 template, or slide text. Credentials are never logged.

Per-subsystem tuning is available in code: loggers are nested under `dial_pptx` (`dial_pptx.visual_qa`, `dial_pptx.tool.export_presentation`, `dial_pptx.utils.template`, …), so a single subsystem can be raised or lowered independently of `LOG_LEVEL`.

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
