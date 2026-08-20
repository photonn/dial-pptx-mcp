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

There is no linter or formatter configured. CI (`.github/workflows/ci.yml`) runs the unittest suite on Python 3.10 and 3.12,
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
(move/resize/font-size/set-text/word-wrap/delete) applied with python-pptx, re-render, repeat up to
`VISUAL_QA_MAX_ITERATIONS`. Keep repairs whitelist-driven — never execute model-supplied code or widen the operation set
without validation.

Slide scoping runs through the whole stack and must stay consistent: LibreOffice always converts the entire deck, so a
subset only selects pages to rasterize; `image_slides` maps images back to absolute slide numbers for `plan_repairs`; and
`apply_repairs(..., allowed_slides=)` drops operations aimed at slides outside the scope. Issue slide numbers reported to
the agent are always absolute, which is what the `review_prompt(..., slides=)` mapping note buys.

`_visual_qa_gate` in `tools/presentation_tools.py` is the **optional** legacy behaviour (`VISUAL_QA_EXPORT_GATE=true`,
off by default); with it off, export just reports `visual_qa: passed|unverified|unavailable` from the store's dirty flag.
Only a clean whole-deck inspection calls `clear_dirty`; both QA tools are in `state._NON_EDITING_TOOLS` so the wrapper
does not re-dirty a deck they just certified. The reviewer reaches the model either at a direct OpenAI Responses-API
endpoint or as a DIAL Core deployment (`_resolve_provider`); Azure and DIAL's Azure upstream both reject calls without
`?api-version=`, added by `_with_api_version`.

**DIAL file I/O (`dial_client.py`).** `resolve_dial_auth_headers` implements `DIAL_AUTH_MODE`: `auto` prefers credentials
pulled off the incoming MCP request (so exports land in the end user's own bucket — only works when the server is deployed
under the DIAL host) and falls back to `DIAL_API_KEY`; `caller` fails instead of falling back; `server` always uses the
server key. Templates arrive as base64/data: URIs resolved by Quick Apps, not as paths; decks leave as
`files/{bucket}/{path}` URLs.

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
