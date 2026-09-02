# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A remote, multi-tenant PowerPoint-generation MCP server for EPAM AI DIAL. It is a fork of
[GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) (MIT) with full upstream
history preserved. The ~30 upstream content/formatting tools are essentially unchanged; this project adds the
remote-transport, multi-tenant state, DIAL file I/O, and visual-QA layers. `docs/UPSTREAM_README.md` is the upstream tool
reference; `README.md` documents this project's configuration and DIAL integration.

Upstream-inherited code (`tools/`, `utils/`, most of `ppt_mcp_server.py`) follows a different style and quality bar than the
DIAL layers (`state.py`, `dial_client.py`, `visual_qa.py`, `visual_fix.py`). Match the file you are editing.

## Commands

```bash
uv venv --python 3.12 .venv && uv pip install -r requirements.txt

.venv/bin/python -m unittest discover -s tests -v          # all tests
.venv/bin/python -m unittest tests.test_visual_gate -v     # one module
.venv/bin/python -m unittest tests.test_state.TestPresentationStore.test_ttl_expiry -v  # one test

.venv/bin/python spike/fidelity_spike.py <template.pptx> [out.pptx]  # template-fidelity check
.venv/bin/python spike/http_client_check.py                          # transport smoke test

.venv/bin/python ppt_mcp_server.py                    # stdio (local MCP clients)
.venv/bin/python ppt_mcp_server.py -t http --host 0.0.0.0 -p 8000    # /mcp endpoint
docker build -t dial-pptx-mcp . && docker run -p 8000:8000 dial-pptx-mcp
```

There is no linter or formatter configured. CI (`.github/workflows/ci.yml`) runs the unittest suite on Python 3.10, 3.12 and 3.14 (3.14 is what the Docker image ships),
the fidelity spike against the bundled demo deck, and a Docker build. Tests requiring LibreOffice self-skip when `soffice`
is absent, so a green local run may still fail in CI — install LibreOffice or check the CI result.

Config comes exclusively from environment variables (`.env` next to `ppt_mcp_server.py` is loaded at startup, real env vars
win). `.env.example` and the README table are the authority — keep all three in sync when adding a variable.

## Architecture

**Registration.** `ppt_mcp_server.py` builds one `FastMCP` app and calls a `register_*_tools(app, presentations, ...)`
function per module in `tools/`. Validators (`is_positive`, `is_valid_rgb`, …) and `add_shape_direct` are defined in the
server and injected as arguments — that is why tool modules take long parameter lists. `tools/` holds MCP tool definitions
(argument validation, dict responses); `utils/` holds the python-pptx work and is re-exported flat through `utils/__init__.py`
(`import utils as ppt_utils`).

`get_server_info` derives everything it reports: the version from `pyproject.toml` via `_project_version()` (parsed with a
regex, not `tomllib`, because CI still runs 3.10) and the tool count from the live registry, since registration is dynamic.
Don't reintroduce hardcoded versions or counts there — the upstream ones were both wrong by the time anyone noticed.

**Slide structure (`utils/slide_utils.py`, `tools/slide_tools.py`).** python-pptx's only entry point is
`slides.add_slide(layout)`, so delete/move/duplicate/cross-deck-copy are implemented at the OPC level: `<p:sldIdLst>`
edits for delete and move, and for duplication a deepcopy of the slide XML plus a rebuilt relationship set that
**preserves the source's rIds** (the copied XML still carries them, so `rels.get_or_add`'s own numbering would break
every `r:embed`). Image/media/font parts are shared, chart/SmartArt/embedded parts are cloned — otherwise
`update_chart_data` on a copy rewrites the original's chart. `cSld` and `spTree` are emptied and refilled rather than
replaced: python-pptx binds a slide's `shapes` collection to the spTree *element* on first access, and `add_slide`
touches it while cloning layout placeholders, so swapping the element out silently detaches every later edit. The
notesSlide rel is never copied (one notes part per slide); the notes text is transferred instead. **The intended
template flow is duplicate-then-fill**, not `add_slide` — a layout holds placeholders, not the template's artwork.

**Structural validation (`deck_validation.py`, `tools/validation_tools.py`).** The axis visual QA cannot see: a deck
with a dangling rId renders in LibreOffice and opens in python-pptx and still fails in PowerPoint. Severities are
`error` (file may not open) / `warning` (user-visible defect) / `info` (advisory), every problem carries a `fix` naming
the tool that resolves it, and `_structure_summary` folds it into export **without blocking** — refusing to deliver a
finished deck over a warning costs the user more than the warning does. Empty placeholders are deliberately not
reported: PowerPoint draws their prompt text only in edit view, and flagging all 68 of them on the demo fixture buried
the two real findings.

**Fonts (`fonts.py`).** Visual QA renders through LibreOffice, so its text-fit verdicts are only trustworthy for fonts
with metric-compatible substitutes (Liberation Sans/Serif/Mono, Carlito, Caladea). Everything else is substituted by
similarity and the widths differ. `unreliable_fonts_in` drives both an `info` problem in validation and a caveat
appended to `REVIEW_PROMPT`, telling the reviewer to report only substantial overflow for text in those fonts. Do not
"fix" a template's fonts to satisfy QA — brand beats QA precision, which is what the caveat exists to make possible.

**Design guidance (`docs/DESIGN_GUIDANCE.md`, `tools/guidance_tools.py`).** The tools answer "how do I place this";
nothing answered "what should this slide look like". The document is the single source of truth (read from disk,
mtime-cached, split on its own `## N. Title` headings) and is served whole or by section, so it costs nothing until a
caller is building. Its premise is that template mode is the default: inherit the user's design, don't invent one over
it. Keep it consistent with the tool names it references.

**Previews (`previews.py`, `tools/preview_tools.py`).** Two composers with opposite audiences, sharing one renderer.
`render_slide_previews` builds contact sheets for choosing a template slide: the agent cannot look at an image, so the
vision-model description is the part it can act on and the uploaded sheet is for the person, which is why an upload
failure must not lose the descriptions. `render_deck_summary_card` is the delivery-time mirror — the whole deck in
**exactly one** image, no slide cap and no vision call, attached beside the exported file so the user sees the result
without opening PowerPoint. There the stored image *is* the deliverable, so a failed upload is an error, not a note.
Its column count is fitted per deck (`auto_columns`) because a fixed grid turns a long deck into a stripe, and cells
scale to the `CARD_MAX_*` budget except that `CARD_MIN_CELL_WIDTH` wins: an over-tall card scrolls, an illegible one is
useless. Both registered only when `soffice` exists.

**Combo charts (`utils/combo_chart_utils.py`).** A plot area holds a *list* of chart-group elements, each naming its
axis pair, so `add_combo_chart` builds an ordinary single-type chart with every series (that is what writes a correct
embedded workbook and the shared category caches) and then redistributes the `c:ser` elements into per-(type, axis)
groups, adding a hidden `catAx` and a right-hand `valAx` when a secondary axis is asked for. Never fall back to a
rendered image: visual repair can only move or delete a picture. **Address value axes by id** — python-pptx's
`chart.value_axis` returns the *second* `valAx` when a chart has two (it assumes a scatter chart), so titling "the
value axis" puts the left axis' title on the right-hand one.

`add_chart`'s `categories` argument means something different for `scatter`, the one type with no category axis: it
carries the **x values** and they must parse as numbers (`parse_scatter_x_values`), because both axes are numeric and
the series need `c:xVal`. That is also why scatter takes `XyChartData` rather than `CategoryChartData` — the type was
advertised for a long time while every call raised.

**Renderer-divergent defaults.** Visual QA renders through LibreOffice but the deck is delivered to PowerPoint, so any
setting left *implicit* in the XML is a place where the two can disagree and the QA render will not show it. Two known
classes, both fixed at the point of creation rather than in QA:

- *Omitted `c:` booleans.* ECMA-376 reads a missing boolean as **true**; PowerPoint applies that, LibreOffice does not.
  python-pptx's bar/column/pie writers emit no `c:varyColors` (PowerPoint then colours a single-series bar chart one
  colour per category and lists the categories in the legend) and `chart.has_legend = True` inserts a bare `<c:legend/>`
  with no `c:overlay` (PowerPoint lays the legend over the plot) and no `c:legendPos`. `normalize_chart_defaults` in
  `utils/content_utils.py` runs at the end of `add_chart` and writes `varyColors`, `overlay`, `legendPos` and
  `plotVisOnly` out; `format_chart` sets `autoTitleDeleted` for an untitled chart, since PowerPoint would otherwise
  auto-title it from the series name. `add_combo_chart` already builds its own explicit XML and needs none of this.
- *Half-written `a:xfrm`.* Setting one dimension on a placeholder that inherits its geometry (`shape.width = ...` with no
  `a:xfrm` present) makes python-pptx create a transform holding only what was set — no `a:off`, or an `a:ext` with a zero
  extent. LibreOffice falls back to the layout and renders it in place; PowerPoint reads it literally and parks the shape
  in the slide's top-left corner. python-pptx also reports the *inherited* value for the missing half, so the defect is
  invisible through the API too. `pin_inherited_geometry` (`utils/slide_utils.py`) materializes all four values first and
  is called by `visual_fix`'s `move_shape`/`resize_shape`; `deck_validation._check_transform` reports any that slip
  through as `partial_transform`. Any new code that writes shape geometry must pin first.

**Two unrelated meanings of "template".** (1) A corporate `.pptx`/`.potx` file that *becomes* the presentation, preserving
theme/masters/layouts byte-for-byte — `create_presentation_from_template{,_content}` in `tools/presentation_tools.py`, with
`.potx` handled by content-type coercion in `utils/presentation_utils.py`. (2) `slide_layout_templates.json` — 23
hardcoded layout recipes with their own color schemes, driving `tools/template_tools.py` and `utils/template_utils.py`.
The second set paints over corporate branding and is irrelevant to the DIAL use case; don't extend it or route branded work
through it.

**State (`state.py`).** `PresentationStore` is a `MutableMapping` (so upstream call sites work unchanged) keyed by
server-generated UUIDs that act as unguessable capabilities, with TTL + LRU eviction and a per-deck lock. There is no
"current presentation": `get_current_presentation_id()` deliberately returns `None`, so every tool must receive an explicit
`presentation_id`. Upstream's `list_presentations`/`switch_presentation` were removed as multi-tenant leaks — don't
reintroduce them.

`serialize_per_presentation(app, presentations)` runs in `main()` and monkey-patches every registered tool's `fn` to
(a) hold that deck's lock for the call (python-pptx is not thread-safe) and (b) mark the deck dirty on any successful
non-read-only, non-export call. It reaches into `app._tool_manager._tools`, which is why `mcp[cli]` is pinned `<2.0`; it
degrades to a warning if those internals change. **Consequence: a new tool must declare `readOnlyHint=True` in its
`ToolAnnotations` if it doesn't modify the deck**, or every call to it triggers a fresh visual-QA pass.

**Visual QA (`visual_qa.py`, `visual_fix.py`, `tools/visual_tools.py`).** Agent-driven, not an export gate: when a vision
LLM is configured, `register_visual_tools` exposes `visual_inspect_slides` (read-only review) and `visual_repair_slides`
(review → repair → re-review loop), both taking an optional 1-based `slides` list so the orchestrator can check the one
slide it just built, as often as it likes. Pipeline: render via LibreOffice → PDF → PNG, vision-LLM review, then
`visual_fix.plan_repairs` asks the model for a plan of **whitelisted, validated** operations
(move/resize/font-size/fit-text/autofit/set-text/word-wrap/delete, plus table column-width/row-height/cell-text and
chart legend/data-label toggles and axis titles) applied with python-pptx, re-render, repeat up to
`VISUAL_QA_MAX_ITERATIONS`. Keep repairs whitelist-driven — never execute model-supplied code or widen the operation set
without validation.

Text in a deck is not only in text frames, and both halves of the loop must keep covering the rest: `REVIEW_PROMPT`
asks explicitly about chart axis/data labels, table cells and diagram/SmartArt node labels, and `describe_slides`
reports table geometry and chart structure so the planner can address them. A new "text container" needs work in both
places plus an operation in `apply_repairs` — a reported issue that no whitelisted op can fix just burns iterations
until the budget runs out. When a round applies nothing the loop stops immediately and reports `skipped_reasons` plus
a `repair_note`: `bad shape_index` almost always means the target lives on the layout or master (slide numbers,
footers), which `slide.shapes` does not expose and no operation can reach.

`fit_text` computes its own size (`estimate_fit_font_size`) instead of taking one from the model, and grows as well as
shrinks. The geometric estimate (`CHAR_WIDTH_RATIO`, `LINE_HEIGHT_RATIO`, `FIT_SLACK`) is deliberately crude — the
re-render is the real check — but the growth cap (`MAX_GROWTH_FACTOR`, `DEFAULT_GROWTH_CEILING_PT`) is policy, not
approximation: without it any short string "fits" at the 96pt ceiling.

Slide scoping runs through the whole stack and must stay consistent: LibreOffice always converts the entire deck, so a
subset only selects pages to rasterize; `image_slides` maps images back to absolute slide numbers for `plan_repairs`; and
`apply_repairs(..., allowed_slides=)` drops operations aimed at slides outside the scope. Issue slide numbers reported to
the agent are always absolute, which is what the `review_prompt(..., slides=)` mapping note buys.

`convert_with_soffice` is the single LibreOffice entry point (PDF export, legacy `.ppt` import, and the QA renderer all
go through it). It owns the isolated `-env:UserInstallation` profile, which is what stops concurrent conversions
fighting over the shared profile lock — a failure that shows up under load, not in tests.

`_visual_qa_gate` in `tools/presentation_tools.py` is the **optional** legacy behaviour (`VISUAL_QA_EXPORT_GATE=true`,
off by default); with it off, export just reports `visual_qa: passed|unverified|unavailable` from the store's dirty flag.
Only a clean whole-deck inspection calls `clear_dirty`; both QA tools are in `state._NON_EDITING_TOOLS` so the wrapper
does not re-dirty a deck they just certified. The reviewer reaches the model either at a direct OpenAI Responses-API
endpoint or as a DIAL Core deployment (`_resolve_provider`); Azure and DIAL's Azure upstream both reject calls without
`?api-version=`, added by `_with_api_version`.

**Images (`tools/image_tools.py`).** The server never generates images; the orchestrator does, stores the result in DIAL
files, and passes the `files/{bucket}/{path}` URL to `add_image_from_dial_url`, which downloads the bytes through
`DialFileClient` (same `DIAL_AUTH_MODE` resolution as export) so a multi-MB PNG never crosses the agent's context.
Upstream's `manage_image(source_type="base64")` stays as the small-asset fallback. `_place` does the geometry: `contain`
(default) scales into the given box and centres, `cover` fills and crops via `Picture.crop_*`, `stretch` distorts, one
dimension scales proportionally, neither keeps native size clamped to the slide. Default to non-distorting fits — visual
QA can move, resize and delete a picture but has no operation that un-distorts one. `DIAL_IMAGE_MAX_MB` bounds the
download. **`image_url` is `Annotated[str, Field(json_schema_extra={"dial_url": True})]` and must stay that way**: DIAL
Quick Apps scans each tool's input schema and grants the toolset's per-request key access to the file named by any
`dial_url` parameter. That key otherwise reaches only its own bucket and `appdata/{this-deployment}`, so every image
another deployment generated 403s — no `DIAL_AUTH_MODE` value fixes it. Any future parameter taking a DIAL file URL
needs the same annotation (`tests/test_image_tools.py` asserts the flag survives to `tools/list`).

**Icons (`svg_icons.py`, `tools/icon_tools.py`, `docs/ICON_GUIDANCE.md`).** Same split as images — the orchestrator draws the SVG, the server renders it — with a review in between. OOXML has no route to SVG, so an icon reaches a slide as a PNG or not at all; rasterizing is PyMuPDF (already a dependency, so icons work without LibreOffice). **The PNG stays on the server** — `svg_icons.IconStore`, UUID handles with the same TTL/LRU bounds as `PresentationStore`, placed by `add_icon_to_slide`. A round trip through DIAL storage looks tidier and does not work: a file this server writes lands in `{user}/appdata/dial-pptx-mcp/`, and placing it means the *orchestrator* asking Core to grant that file to the toolset key, which it cannot do for a folder it neither owns nor is — Core answers 403 before the tool is entered. Exports and summary cards are unaffected because their URLs go to the end user, who owns the bucket; an image-model PNG is unaffected because it reaches the conversation as an attachment the orchestrator can share. Don't "simplify" icons back onto `add_image_from_dial_url`. The review exists because a hand-written path is valid XML long before it is a pictogram, the agent cannot see what it drew, and the visual-repair whitelist can only move, resize or delete a picture — a bad icon on a slide is unfixable, so it is judged alone (full size *and* downscaled to ~1in, composited onto `slide_background`) before placement. **A failed review uploads nothing**: an `image_url` in the response means the icon is fit to place, which is what makes the retry loop unambiguous. The review is gated on `visual_qa.vision_configured()`, not `enforcement_enabled()`: `VISUAL_QA_ENFORCE` hides the slide inspect/repair tools, which is a different decision from whether a model exists to look at one icon. Two guards run before the model: an empty render is refused with its likely cause, and near-blank/near-solid coverage comes back as a `render_note`. The SVG is untrusted input parsed in-process — DOCTYPE/ENTITY, `<script>`, `<foreignObject>`, `<image>`, event handlers and any external `href`/`url()` are refused (XXE/SSRF), and so is text, since substituted glyphs are the artifact class the review is there to catch. `docs/ICON_GUIDANCE.md` is the single source of truth for style, served whole and mtime-cached like the design guidance; keep its tool names and its two variants consistent with what the tool accepts.

**DIAL file I/O (`dial_client.py`).** `resolve_dial_auth_headers` implements `DIAL_AUTH_MODE`: `auto` prefers credentials
pulled off the incoming MCP request (so exports land in the end user's own bucket — only works when the server is deployed
under the DIAL host) and falls back to `DIAL_API_KEY`; `caller` fails instead of falling back; `server` always uses the
server key. Templates arrive as base64/data: URIs resolved by Quick Apps, not as paths; decks leave as
`files/{bucket}/{path}` URLs. `_file_request_url` is the single gate on the download side: it accepts the relative
form, the Core API's `v1/files/...` and the chat frontend's `api/files/...` proxy path, absolute or not, and always
rebuilds the request against `DIAL_CORE_URL` — so a link's host is only ever an authorization question, never where
the bytes come from. Foreign hosts are refused (`DIAL_PUBLIC_URL` lists this installation's public aliases); keep it
that way, an MCP tool that fetches arbitrary URLs is an SSRF primitive. A 401/403/404 on a read is raised as a
`DialConfigError` carrying `_access_hint`, which names the identity used, the bucket it owns and the appdata
deployment that owns the file — because DIAL storage is per-user and per-deployment, and the usual cause is
identity, not the URL. Never retry a refused read under the server key to "fix" a caller-credential failure: that
turns an authorization error into privilege escalation. `get_dial_storage_info` exposes the same identity to the
agent for diagnosis.

**Logging (`logging_utils.py`).** `configure_logging()` runs in `ppt_mcp_server.py` **before FastMCP is imported**, which is what suppresses FastMCP's RichHandler: its `logging.basicConfig` call is a no-op once the root logger has a handler. `SingleLineFormatter` folds newlines (messages and tracebacks alike) into ` | ` so every record is one line on stderr — stdout is the MCP channel on stdio transport. uvicorn is redirected by **mutating `uvicorn.config.LOGGING_CONFIG` in place**; rebinding the module attribute does nothing because `uvicorn.Config.__init__` binds it as a default argument at import time.

Modules call `get_logger(__name__)`, nesting under `dial_pptx.*`. `LOG_LEVEL` (default INFO, unknown values fall back rather than raising) sets root level. Most per-tool logging is centralized: the `serialize_per_presentation` wrapper in `state.py` emits `tool_ok`/`tool_error`/`tool_raised` with `duration_ms` for every tool, so individual tool modules should log only what the wrapper cannot see (what was created, why an internal step degraded) instead of re-logging returned errors. Never log handles or argument payloads directly — use `state.short_id()` and the `_arg_summary` pattern.

**Transport security.** `_transport_security_for` overrides FastMCP's localhost-only Host allowlist, which otherwise 421s
Kubernetes service DNS names once bound to `0.0.0.0`. `PPT_MCP_ALLOWED_HOSTS=*` disables the check entirely.

## Conventions

- Tools return plain dicts and report failures as `{"error": "..."}` rather than raising — the dirty-marking wrapper and the
  export gate both rely on that shape.
- Error strings are agent-facing instructions; the existing ones tell the model what to do next (e.g. which handle to pass,
  that a QA failure is not a retry request). Write new ones the same way.
- New log lines follow `event_name key=value key=value` with a stable event name first, so lines stay greppable; pass values as `%s` args rather than f-strings so DEBUG-level formatting costs nothing when disabled.
- Local test templates go in `templates-local/` (gitignored). `mcp_all_tools_templates_effects_demo.pptx` is the committed
  fixture CI uses.
- `gh` resolves this repo to the `upstream` remote by default — pass `--repo photonn/dial-pptx-mcp` to every `gh` command.
  Note also that upstream's `v2.0.7` tag predates this fork's own line, which starts at `v1.0.0`.
