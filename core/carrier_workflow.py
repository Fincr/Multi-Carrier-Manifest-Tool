"""
Workflow decisions shared by the single-sheet and batch processing paths.

Both paths kept their own copy of "which carriers skip local printing", and
they drifted: Royal Mail was missing from both lists, so its internal carrier
sheet went to the printer alongside the manifest PDF the portal prints. Keeping
the rule in one place means the two paths cannot disagree again.
"""


# Carriers whose manifest is produced — and printed — by their portal. The file
# this tool writes for them is an upload artefact, not paperwork.
#
# Deutsche Post is deliberately absent: its carrier sheet, with the
# (EMB) Manifest tab removed, IS the paperwork, so it prints here as well as
# printing the portal's PDF later.
_PORTAL_PRINTS_MANIFEST = ('spring', 'landmark', 'royal mail')


def prints_output_file(carrier_name) -> bool:
    """
    Whether this carrier's local output file is the thing to send to the printer.

    Matching is loose because the carrier name comes from cell B3 of the
    carrier sheet — free text such as "Spring GDS" — rather than a registry key.

    An unknown or missing name prints, so a newly added carrier's paperwork
    cannot silently vanish.
    """
    name = str(carrier_name or "").lower().strip()
    return not any(marker in name for marker in _PORTAL_PRINTS_MANIFEST)
