"""
Single-line, level-controlled logging for the MCP server.

Every log record is emitted as exactly ONE line on stderr, so a pod's
terminal (and any line-oriented log collector) shows one event per line:

    2026-08-20T09:14:02.517Z INFO  dial_pptx.export presentation_id=9f3c… slides=12 duration_ms=418

Embedded newlines — in messages, and in tracebacks attached via
``exc_info`` — are folded into " | " separators rather than being allowed
to break the record across lines. That keeps stack traces greppable
without the multi-line, boxed output FastMCP's default RichHandler
produces.

Level comes from the LOG_LEVEL environment variable (DEBUG, INFO, WARNING,
ERROR, CRITICAL; default INFO). It applies to this server's own loggers and
to the libraries underneath it:

- FastMCP calls ``logging.basicConfig`` (installing a RichHandler) when the
  app starts. ``configure_logging`` runs first and leaves a handler on the
  root logger, which makes that call a no-op — the Rich handler never gets
  installed.
- uvicorn otherwise applies its own dictConfig, giving its loggers private
  handlers and a different format. ``configure_logging`` replaces
  ``uvicorn.config.LOGGING_CONFIG`` with a minimal dict so uvicorn's records
  propagate to the root logger and come out in this format too.
- stdout is the MCP channel on the stdio transport, so logs always go to
  stderr.
"""
import logging
import os
import re
import sys
import time

DEFAULT_LEVEL = "INFO"
LEVEL_NAMES = ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL")

# Root logger name for this server's own modules; get_logger() nests under it.
ROOT_LOGGER_NAME = "dial_pptx"

_TIME_FORMAT = "%Y-%m-%dT%H:%M:%S"
_LINE_FORMAT = "%(asctime)s.%(msecs)03dZ %(levelname)-5s %(name)s %(message)s"

_NEWLINES = re.compile(r"\s*[\r\n]+\s*")

_configured = False


def resolve_level(value=None) -> str:
    """Normalize a LOG_LEVEL value to a level name. Unknown values fall back
    to INFO (a server must not fail to start over a typo in a log setting)."""
    raw = (value if value is not None else os.environ.get("LOG_LEVEL", "")).strip()
    if not raw:
        return DEFAULT_LEVEL
    name = raw.upper()
    if name in LEVEL_NAMES:
        return name
    if name == "WARN":
        return "WARNING"
    if name == "FATAL":
        return "CRITICAL"
    if raw.isdigit():
        # Numeric levels (10/20/30/...) map to the nearest standard name.
        mapped = logging.getLevelName(int(raw))
        if mapped in LEVEL_NAMES:
            return mapped
    return DEFAULT_LEVEL


def flatten(text: str) -> str:
    """Fold a multi-line string into one line with ' | ' separators."""
    return _NEWLINES.sub(" | ", text.strip()) if text else text


class SingleLineFormatter(logging.Formatter):
    """Formatter that guarantees one line per record, tracebacks included."""

    def format(self, record):
        return flatten(super().format(record))

    def formatException(self, exc_info):
        return flatten(super().formatException(exc_info))

    def formatStack(self, stack_info):
        return flatten(super().formatStack(stack_info))


def _quiet_third_party(level: str) -> None:
    """Libraries that log a line per HTTP call would double up with this
    server's own request logging; keep them at WARNING unless debugging."""
    noisy = ("httpx", "httpcore", "urllib3", "PIL", "pymupdf", "fontTools")
    third_party_level = logging.DEBUG if level == "DEBUG" else logging.WARNING
    for name in noisy:
        logging.getLogger(name).setLevel(third_party_level)


def _route_uvicorn_through_root() -> None:
    """Make uvicorn use this format instead of its own dictConfig.

    ``uvicorn.Config.__init__`` binds ``LOGGING_CONFIG`` as a default
    argument value at import time, so rebinding the module attribute has no
    effect — the dict has to be emptied and refilled in place. Left with no
    handlers of their own, uvicorn's loggers propagate to the root handler
    installed here.
    """
    try:
        from uvicorn.config import LOGGING_CONFIG
    except Exception:  # uvicorn is only needed for the http/sse transports
        return
    LOGGING_CONFIG.clear()
    LOGGING_CONFIG.update({"version": 1, "disable_existing_loggers": False})
    for name in ("uvicorn", "uvicorn.error", "uvicorn.access"):
        logger = logging.getLogger(name)
        logger.handlers = []
        logger.propagate = True


def configure_logging(level=None, stream=None, force=False) -> str:
    """Install the single-line handler on the root logger. Idempotent unless
    ``force``. Returns the level name in effect."""
    global _configured
    resolved = resolve_level(level)
    if _configured and not force:
        return resolved

    handler = logging.StreamHandler(stream or sys.stderr)
    formatter = SingleLineFormatter(_LINE_FORMAT, datefmt=_TIME_FORMAT)
    formatter.converter = time.gmtime  # timestamps in UTC
    handler.setFormatter(formatter)

    root = logging.getLogger()
    root.handlers = [handler]
    root.setLevel(getattr(logging, resolved))

    _quiet_third_party(resolved)
    _route_uvicorn_through_root()
    # warnings.warn writes multi-line text straight to stderr; as log records
    # they go through the single-line formatter like everything else.
    logging.captureWarnings(True)

    _configured = True
    return resolved


def get_logger(name: str = None) -> logging.Logger:
    """Logger for a server module, nested under the ``dial_pptx`` namespace.

    Pass ``__name__``; the module's own package path is preserved so
    LOG_LEVEL can be raised or lowered per subsystem by configuring
    e.g. ``dial_pptx.visual_qa`` directly.
    """
    if not name or name in ("__main__", ROOT_LOGGER_NAME):
        return logging.getLogger(ROOT_LOGGER_NAME)
    return logging.getLogger(f"{ROOT_LOGGER_NAME}.{name}")
