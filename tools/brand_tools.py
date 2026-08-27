"""
Brand rules as a deck-scoped capability: attach the files, then check against
them.

The rules themselves live in DIAL file storage beside the template and the
icons; this server knows only the file *names* the operator configured
(BRAND_PROFILE_FILE, BRAND_REFERENCE_DECK_FILE) and never where they sit,
because buckets and folders move and a name does not.

That leaves the orchestrator to resolve the name to a file it can actually
read, which it is in the best position to do — it can see what it has access
to — and which is also the only way the read can succeed: DIAL Quick Apps
grants the request's key access to the files named by `dial_url` parameters,
so a file the server fetched from a configured path of its own would come back
403. Hence `attach_brand_profile`: the agent passes the URLs once, the server
downloads and validates them inside that same request, and keeps the result on
the deck.

Both tools are registered only where BRAND_PROFILE_FILE is set. Everything
downstream — the export summary, the reviewer's focus, the reference deck
shown to the vision model — reads the attached context off the deck, so a
build attaches once and the rest is automatic.
"""
from typing import Annotated, Dict

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.brand")

UNKNOWN_ID = (
    "Unknown or expired presentation_id. Pass the presentation_id returned "
    "by create_presentation, create_presentation_from_template, or "
    "open_presentation"
)

# Same flag as add_image_from_dial_url's image_url, and for the same reason:
# it is what makes DIAL Quick Apps grant this request's key access to the
# file. Without it the brand files are unreadable no matter how the server
# authenticates. See tools/image_tools.py.
DialFileUrl = Annotated[str, Field(json_schema_extra={"dial_url": True})]

SEVERITIES = ("error", "warning", "info")


def not_attached_error(profile_name):
    """What to say when a deck has no brand context but the server expects one."""
    return (
        f"No brand profile is attached to this presentation. This server "
        f"checks decks against '{profile_name}': find that file among the "
        f"ones you have access to and call attach_brand_profile with its DIAL "
        f"file URL, then try again. If you cannot find it, tell the user the "
        f"deck cannot be brand-checked rather than continuing silently."
    )


def register_brand_tools(app: FastMCP, presentations):
    """Register the brand tools, if this deployment names a brand profile."""
    import brand_validation

    if not brand_validation.enabled():
        logger.debug("brand_tools_not_registered reason=no_profile_configured")
        return
    profile_name = brand_validation.profile_file_name()
    reference_name = brand_validation.reference_deck_file_name()
    logger.info("brand_tools_registered profile_file=%s reference_file=%s",
                profile_name, reference_name or "none")

    def _download(url, what):
        from dial_client import DialFileClient, DialConfigError

        try:
            return DialFileClient().download(url)
        except DialConfigError as e:
            raise brand_validation.BrandConfigError(
                f"The {what} at {url} could not be read from DIAL file "
                f"storage: {e}"
            )
        except Exception as e:
            raise brand_validation.BrandConfigError(
                f"The {what} at {url} could not be fetched from DIAL file "
                f"storage ({type(e).__name__}: {e})."
            )

    def attach_brand_profile(
        presentation_id: str,
        profile_url: DialFileUrl,
        # Defaulted to "" rather than Optional[...]: an Optional flags the
        # parameter as anyOf[string, null] in the schema, which buries the
        # dial_url flag Quick Apps scans for. Empty means "not supplied".
        reference_deck_url: DialFileUrl = "",
    ) -> Dict:
        """Placeholder; the real description is assigned below, because it
        has to name the file the operator configured."""
        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}

        context = {}
        try:
            brand_validation.check_file_name(profile_url, profile_name,
                                             "brand profile")
            blob = _download(profile_url, "brand profile")
            profile = brand_validation.parse_profile(
                blob, f"at {profile_url}")
        except brand_validation.BrandConfigError as e:
            logger.warning("brand_attach_failed presentation_id=%s part=profile"
                           " error=%s", short_id(presentation_id), e)
            return {"error": str(e)}
        context["profile"] = profile

        if reference_deck_url:
            if not reference_name:
                return {"error": "This server does not use a brand reference "
                                 "deck. Attach the profile alone."}
            try:
                brand_validation.check_file_name(
                    reference_deck_url, reference_name,
                    "brand reference deck")
                blob = _download(reference_deck_url, "brand reference deck")
                context["reference"] = brand_validation.open_reference_deck(
                    blob, f"at {reference_deck_url}")
            except brand_validation.BrandConfigError as e:
                logger.warning("brand_attach_failed presentation_id=%s "
                               "part=reference error=%s",
                               short_id(presentation_id), e)
                return {"error": str(e)}

        presentations.set_brand(presentation_id, context)
        logger.info("brand_attached presentation_id=%s brand=%s reference=%s",
                    short_id(presentation_id),
                    profile.get("name", "unnamed"), "reference" in context)

        result = {
            "message": f"This deck is now held to the "
                       f"{profile.get('name', 'configured')} brand profile. "
                       f"Follow it as you build; validate_brand_profile "
                       f"reports where you stand.",
            "brand": profile.get("name", "unnamed"),
            "reference_deck": "reference" in context,
            "rules": _rule_summary(profile),
        }
        if reference_name and not reference_deck_url:
            result["reference_deck_note"] = (
                f"This server also names a reference deck, "
                f"'{reference_name}'. Passing it as reference_deck_url gives "
                f"the visual reviewer a real branded deck to compare against."
            )
        return result

    # The description has to name the configured file — an agent that is not
    # told what to look for cannot find it — so it is composed here and the
    # tool registered by hand. A docstring cannot be an f-string: that is an
    # expression, not a docstring, and the tool would register with no
    # description at all.
    attach_brand_profile.__doc__ = f"""Give this deck the brand rules it will
        be held to.

        Call it once, right after creating the presentation and before you
        build slides — the rules are cheaper to follow than to retrofit.

        This server expects its rules in a file named '{profile_name}'. Look
        through the files you have access to, find that one, and pass its DIAL
        file URL here; the server reads it during this call and remembers it
        for this deck. Where the file lives is yours to work out — only the
        name is fixed, because buckets and folders move.

        profile_url: DIAL file URL of '{profile_name}'.
        reference_deck_url: {"DIAL file URL of '" + reference_name + "', the brand's reference deck — a real on-brand deck the visual reviewer is shown alongside your slides, so it compares against the brand instead of guessing at it." if reference_name else "not used by this server; leave it out."}

        Once attached: validate_brand_profile checks the deck against the
        rules, the visual QA tools enforce the parts of the profile that need
        judgement rather than measurement, and export reports where the deck
        stands. Without it none of that runs, and the deck is delivered
        unchecked.

        Returns the rules themselves, so you can build to them rather than
        only be judged by them.
        """
    app.tool(
        annotations=ToolAnnotations(
            title="Attach Brand Profile",
            readOnlyHint=True,
        ),
    )(attach_brand_profile)

    @app.tool(
        annotations=ToolAnnotations(
            title="Validate Deck Against Brand Profile",
            readOnlyHint=True,
        ),
    )
    def validate_brand_profile(presentation_id: str,
                               min_severity: str = "warning") -> Dict:
        """Check the deck against the brand rules attached to it.

        Registered only where a brand profile is configured, so if you can see
        this tool the deck is being held to a house style: allowed fonts, a
        minimum text size, an approved palette, a safe area content must stay
        inside, chrome every slide must carry (separator rule, logo, page
        number), and how often the deck may repeat one background family.

        Attach the rules first with attach_brand_profile. Then call this while
        you build, not only at the end — every finding names the slide, the
        shape and the tool that fixes it, so a slide can be corrected while
        you still remember what it was for. It is cheap: no rendering and no
        model call.

        It does not judge appearance or structure — visual_inspect_slides and
        validate_presentation do, and all three are worth running before you
        export. The rules that need judgement rather than measurement (a
        headline must carry a message; a slide must not be plain text on
        white) are enforced by the visual reviewer instead, from the same
        profile.
        """
        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}
        if min_severity not in SEVERITIES:
            return {"error": f"Invalid min_severity: {min_severity}. Must be "
                             f"one of {', '.join(SEVERITIES)}."}

        context = presentations.brand_for(presentation_id) or {}
        profile = context.get("profile")
        if not profile:
            # Checking a deck against nothing and reporting it clean would be
            # worse than saying the rules never arrived.
            return {"error": not_attached_error(profile_name)}

        pres = presentations[presentation_id]
        try:
            report = brand_validation.validate_brand(pres, profile)
        except Exception as e:
            logger.error("brand_validation_failed presentation_id=%s error=%s",
                         short_id(presentation_id), e)
            return {"error": f"Failed to check the presentation against the "
                             f"brand profile: {e}"}

        cutoff = SEVERITIES.index(min_severity)
        report["problems"] = [p for p in report["problems"]
                              if SEVERITIES.index(p["severity"]) <= cutoff]
        report["min_severity"] = min_severity

        counts = report["counts"]
        logger.info("brand_validation_done presentation_id=%s brand=%s "
                    "slides=%d errors=%d warnings=%d",
                    short_id(presentation_id), report["brand"],
                    report["slides"], counts["error"], counts["warning"])

        if not report["problems"]:
            report["message"] = (
                f"No {report['brand']} brand problems at severity "
                f"'{min_severity}' or above."
            )
        else:
            report["message"] = (
                f"{counts['error']} error(s) and {counts['warning']} "
                f"warning(s) against the {report['brand']} brand profile. Each "
                "problem names the slide, the shape and the fix; correcting "
                "them now is cheaper than after the deck is written."
            )
        return report


def _rule_summary(profile):
    """The rules in a form the agent can build to, rather than only be judged
    by — it has just been handed a profile it cannot otherwise read."""
    summary = {}
    fonts = (profile.get("fonts") or {}).get("allowed")
    if fonts:
        summary["fonts"] = fonts
    for key in ("min_font_pt", "palette_rgb", "max_consecutive_same_family"):
        if profile.get(key):
            summary[key] = profile[key]
    families = profile.get("families") or {}
    areas = {name: rules.get("safe_area_in")
             for name, rules in families.items() if rules.get("safe_area_in")}
    if areas:
        summary["safe_area_in"] = areas
    required = {name: rules.get("require")
                for name, rules in families.items() if rules.get("require")}
    if required:
        summary["required_chrome"] = required
    notes = profile.get("review_notes")
    if notes:
        summary["review_notes"] = notes
    return summary
