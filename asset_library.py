"""
Server-side image library: brand icons and other pictures shipped with the
deployment instead of with each agent.

An operator mounts a folder of images into the pod and points
`PPT_ASSET_PATH` at it. The agent lists what is there (`list_assets`) and
places an entry by name (`add_asset_to_slide`). Nothing goes through DIAL
file storage, so no per-user sharing is needed, and the agent never builds a
path: names are checked against the folder's own listing, which keeps the
tool from becoming a way to read arbitrary files off the pod.
"""
import os

from logging_utils import get_logger

logger = get_logger("asset_library")

# Formats python-pptx embeds as pictures. SVG is not one of them (OOXML has
# no route to SVG), so it is not listed even if the folder holds some.
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif")


class AssetError(Exception):
    """The library is not configured or the name does not resolve. The
    message is agent-facing."""


def asset_dir():
    """The configured library folder, or None when unset."""
    raw = os.environ.get("PPT_ASSET_PATH", "").strip()
    return os.path.expanduser(raw) if raw else None


def _require_dir():
    directory = asset_dir()
    if directory is None:
        raise AssetError("No asset library is configured on this server "
                         "(PPT_ASSET_PATH is not set). Draw the icon with "
                         "render_svg_icon instead.")
    if not os.path.isdir(directory):
        logger.warning("asset_dir_missing path=%s", directory)
        raise AssetError("The asset library folder is not available on this "
                         "server right now. Draw the icon with render_svg_icon "
                         "instead.")
    return directory


def list_assets(query=None):
    """Sorted image file names in the library, optionally filtered by a
    case-insensitive substring."""
    directory = _require_dir()
    names = sorted(
        name for name in os.listdir(directory)
        if not name.startswith(".")
        and name.lower().endswith(IMAGE_EXTENSIONS)
        and os.path.isfile(os.path.join(directory, name))
    )
    if query:
        needle = query.strip().lower()
        names = [name for name in names if needle in name.lower()]
    return names


def resolve(name):
    """Absolute path of a library entry. Only exact names from
    list_assets() resolve — no directories, no relative segments."""
    directory = _require_dir()
    if not name or name != os.path.basename(name) or name.startswith("."):
        raise AssetError(f"'{name}' is not an asset name. Pass a name exactly "
                         "as list_assets returned it.")
    if name not in list_assets():
        raise AssetError(f"No asset named '{name}'. Call list_assets (with a "
                         "query to narrow it) and pass one of the returned "
                         "names, or draw the icon with render_svg_icon.")
    return os.path.join(directory, name)
