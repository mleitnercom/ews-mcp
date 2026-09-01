"""Shared mail-folder tree helpers.

Folder enumeration is used by the background mirror, live reads and folder
mutations.  Keeping it here prevents those three paths from drifting (the old
v3 implementation had three subtly different folder resolvers).
"""

from __future__ import annotations

import logging
from typing import Any

from .gateway.client import WELL_KNOWN

logger = logging.getLogger(__name__)

MAIL_FOLDER_CLASS = "IPF.Note"


def folder_id(folder: Any) -> str | None:
    raw_id = getattr(folder, "id", None)
    return str(raw_id) if raw_id else None


def same_folder(left: Any, right: Any) -> bool:
    left_id, right_id = folder_id(left), folder_id(right)
    return bool(left_id and right_id and left_id == right_id)


def is_mail_parent(account: Any, folder: Any) -> bool:
    """Whether ``folder`` may contain normal IPF.Note mail folders."""
    if same_folder(folder, getattr(account, "msg_folder_root", None)):
        return True
    container_class = (
        getattr(folder, "folder_class", None)
        or getattr(type(folder), "CONTAINER_CLASS", None)
    )
    # Some lightweight adapters do not expose FolderClass.  Exchange itself
    # validates them, so absence is not grounds for a false rejection.
    return not container_class or str(container_class) == MAIL_FOLDER_CLASS


def child_named(parent: Any, name: str, *, exclude_id: str | None = None) -> Any:
    wanted = name.casefold()
    for child in list(getattr(parent, "children", None) or []):
        raw_id = folder_id(child)
        if exclude_id and raw_id == exclude_id:
            continue
        if str(getattr(child, "name", "") or "").casefold() == wanted:
            return child
    return None


def descendant_ids(folder: Any) -> set[str]:
    out: set[str] = set()

    def walk(node: Any) -> None:
        for child in list(getattr(node, "children", None) or []):
            raw_id = folder_id(child)
            if raw_id and raw_id not in out:
                out.add(raw_id)
                walk(child)

    walk(folder)
    return out


def subtree_stats(folder: Any) -> dict[str, int]:
    items = int(getattr(folder, "total_count", None) or 0)
    subfolders = 0

    def walk(node: Any) -> None:
        nonlocal items, subfolders
        for child in list(getattr(node, "children", None) or []):
            subfolders += 1
            items += int(getattr(child, "total_count", None) or 0)
            walk(child)

    walk(folder)
    return {"items": items, "subfolders": subfolders}


def collect_folder_rows(account: Any, root: Any = None) -> list[dict[str, Any]]:
    """Snapshot a message-folder subtree into mirror-shaped rows.

    Paths are relative to ``root``.  Omitting it snapshots the complete
    message-folder hierarchy and therefore produces canonical mirror paths.
    """
    wk_by_raw: dict[str, str] = {}
    for wk_alias, attr in WELL_KNOWN.items():
        try:
            raw_id = folder_id(getattr(account, attr, None))
            if raw_id:
                wk_by_raw.setdefault(raw_id, wk_alias)
        except Exception as exc:  # noqa: BLE001 — well-known tagging is best effort
            logger.debug("well-known folder lookup failed for %s: %s", attr, exc)

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()

    def walk(folder: Any, prefix: str) -> None:
        for child in list(getattr(folder, "children", None) or []):
            name = str(getattr(child, "name", "") or "")
            path = f"{prefix}/{name}" if prefix else name
            raw_id = folder_id(child)
            grandchildren = list(getattr(child, "children", None) or [])
            if raw_id and raw_id not in seen:
                seen.add(raw_id)
                rows.append({
                    "ews_id": raw_id,
                    "name": name,
                    "path": path,
                    "wk": wk_by_raw.get(raw_id),
                    "total": int(getattr(child, "total_count", None) or 0),
                    "unread": int(getattr(child, "unread_count", None) or 0),
                    "children": len(grandchildren),
                })
                walk(child, path)

    walk(root if root is not None else account.msg_folder_root, "")
    return rows


def row_for_folder(rows: list[dict[str, Any]], raw_id: str) -> dict[str, Any] | None:
    return next((row for row in rows if row["ews_id"] == raw_id), None)
