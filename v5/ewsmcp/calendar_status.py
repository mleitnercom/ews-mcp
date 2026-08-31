"""Stable public calendar free/busy values.

Exchange exposes legacy ``LegacyFreeBusyStatus`` spellings while the MCP
surface uses predictable snake_case values.  Keep the translation in one
module so live DTOs, mirror rows and writes cannot drift apart.
"""

from __future__ import annotations

from typing import Any, Optional

from .errors import ToolError

SHOW_AS_TO_EWS = {
    "free": "Free",
    "working_elsewhere": "WorkingElsewhere",
    "tentative": "Tentative",
    "busy": "Busy",
    "out_of_office": "OOF",
    "no_data": "NoData",
}
EWS_TO_SHOW_AS = {value.lower(): key for key, value in SHOW_AS_TO_EWS.items()}
# Be liberal when reading values materialised as enum-ish strings.
EWS_TO_SHOW_AS.update(
    {
        "working_elsewhere": "working_elsewhere",
        "outofoffice": "out_of_office",
        "out_of_office": "out_of_office",
        "nodata": "no_data",
        "no_data": "no_data",
    }
)
SHOW_AS_VALUES = tuple(SHOW_AS_TO_EWS)


def validate_show_as(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    normalized = str(value).strip().lower()
    if normalized not in SHOW_AS_TO_EWS:
        allowed = ", ".join(SHOW_AS_VALUES)
        raise ToolError(
            "validation",
            f"show_as must be one of: {allowed}; got {value!r}",
        )
    return normalized


def to_ews_show_as(value: Optional[str]) -> Optional[str]:
    normalized = validate_show_as(value)
    return SHOW_AS_TO_EWS[normalized] if normalized is not None else None


def from_ews_show_as(value: Any) -> Optional[str]:
    if value is None:
        return None
    raw = str(value).strip()
    if not raw:
        return None
    # Some enum wrappers stringify as ``LegacyFreeBusyStatus.WorkingElsewhere``.
    raw = raw.rsplit(".", 1)[-1]
    return EWS_TO_SHOW_AS.get(raw.lower())
