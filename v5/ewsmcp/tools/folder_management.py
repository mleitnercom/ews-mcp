"""Atomic mail-folder mutations.

The v3 ``manage_folder(action=...)`` tool mixed reversible writes and
destructive deletion under one safety class.  v5 deliberately exposes four
operations so the dispatcher can gate and audit each one accurately.
"""

from __future__ import annotations

import logging
import time
from typing import Any

from exchangelib import Folder
from exchangelib.items import HARD_DELETE, MOVE_TO_DELETED_ITEMS

from ..errors import ToolError
from ..folders import (
    MAIL_FOLDER_CLASS,
    child_named,
    collect_folder_rows,
    descendant_ids,
    folder_id,
    is_mail_parent,
    row_for_folder,
    same_folder,
    subtree_stats,
)
from ..gateway.client import WELL_KNOWN
from .base import Context, ToolSpec
from .writes import _gateway_write, _idempotent_create

logger = logging.getLogger(__name__)

MAX_FOLDER_NAME = 255


def _validated_name(value: str) -> str:
    name = str(value or "").strip()
    if not name:
        raise ToolError("validation", "folder name cannot be empty")
    if len(name) > MAX_FOLDER_NAME:
        raise ToolError(
            "validation", f"folder name cannot exceed {MAX_FOLDER_NAME} characters")
    if any(char in name for char in ("/", "\\")):
        raise ToolError(
            "validation",
            "folder name cannot contain '/' or '\\' because paths use those separators",
        )
    if any(ord(char) < 32 for char in name):
        raise ToolError("validation", "folder name cannot contain control characters")
    return name


def _ensure_mutable(account: Any, folder: Any) -> None:
    if same_folder(folder, getattr(account, "msg_folder_root", None)):
        raise ToolError("protected_folder", "The mailbox root cannot be modified.")
    if bool(getattr(folder, "is_distinguished", False)):
        raise ToolError(
            "protected_folder",
            f"System folder {getattr(folder, 'name', '')!r} cannot be modified.",
            hint="Create or select a normal user mail folder instead.",
        )


def _ensure_mail_parent(account: Any, folder: Any) -> None:
    if not is_mail_parent(account, folder):
        raise ToolError(
            "validation",
            f"{getattr(folder, 'name', 'That folder')!r} cannot contain mail folders.",
            hint="Use f:root, f:inbox, or another IPF.Note mail folder.",
        )


def _well_known_ref(account: Any, folder: Any) -> str | None:
    raw_id = folder_id(folder)
    for alias, attr in WELL_KNOWN.items():
        try:
            if raw_id and raw_id == folder_id(getattr(account, attr, None)):
                return alias
        except Exception as exc:  # noqa: BLE001 — result tagging is best effort
            logger.debug("well-known folder lookup failed for %s: %s", attr, exc)
    return None


def _folder_ref(ctx: Context, account: Any, folder: Any) -> str:
    return (_well_known_ref(account, folder)
            or ctx.aliaser.alias_for(folder_id(folder) or "", "f"))


def _public_folder(ctx: Context, row: dict[str, Any],
                   parent_ref: str | None = None) -> dict[str, Any]:
    result: dict[str, Any] = {
        "id": ctx.aliaser.alias_for(row["ews_id"], "f"),
        "name": row["name"],
        "path": row["path"],
        "total": row["total"],
        "unread": row["unread"],
        "children": row["children"],
    }
    if row.get("wk"):
        result["wk"] = row["wk"]
    if parent_ref is not None:
        result["parent"] = parent_ref
    return result


def _previous_folder(row: dict[str, Any] | None) -> dict[str, Any] | None:
    if row is None:
        return None
    return {"name": row["name"], "path": row["path"]}


def _store_snapshot(ctx: Context, rows: list[dict[str, Any]]) -> list[str]:
    warnings: list[str] = []
    if ctx.cache is None:
        return warnings
    try:
        ctx.cache.replace_folders(rows)
        ctx.cache.set_sync_state("folders", None, time.time())
    except Exception as exc:  # noqa: BLE001 — Exchange write already succeeded
        logger.warning("folder mirror write-through failed: %s", exc)
        warnings.append("folder mirror refresh failed; background sync will repair it")
    return warnings


def _publish_mutation(ctx: Context, account: Any) -> tuple[list[dict[str, Any]], int,
                                                            list[str]]:
    warnings: list[str] = []
    revision = 0
    try:
        if hasattr(ctx.gateway, "mark_folder_tree_changed"):
            revision = int(ctx.gateway.mark_folder_tree_changed(account))
    except Exception as exc:  # noqa: BLE001 — Exchange write already succeeded
        logger.warning("folder revision publish failed: %s", exc)
        warnings.append("cross-process folder invalidation failed; restart stale clients")
    rows = collect_folder_rows(account)
    warnings.extend(_store_snapshot(ctx, rows))
    return rows, revision, warnings


def _snapshot_row(ctx: Context, account: Any, raw_id: str) -> dict[str, Any]:
    rows = collect_folder_rows(account)
    _store_snapshot(ctx, rows)
    row = row_for_folder(rows, raw_id)
    if row is None:
        raise ToolError("not_found", "Folder disappeared while refreshing its result.")
    return row


def _conflict(name: str) -> ToolError:
    return ToolError(
        "conflict",
        f"A sibling folder named {name!r} already exists.",
        hint="Use list_folders(query=..., match='exact') and reuse its folder id.",
    )


async def _create_folder(ctx: Context, *, name: str, parent: str = "f:root",
                         idempotency_key: str | None = None) -> dict[str, Any]:
    clean_name = _validated_name(name)

    def work(account: Any) -> dict[str, Any]:
        parent_folder = ctx.gateway.resolve_folder(account, parent, ctx.aliaser)
        _ensure_mail_parent(account, parent_folder)
        existing = child_named(parent_folder, clean_name)
        if existing is not None:
            raw_id = folder_id(existing)
            if raw_id is None:
                raise ToolError("upstream_error", "Existing folder has no Exchange id.")
            row = _snapshot_row(ctx, account, raw_id)
            return {
                "raw_id": raw_id, "row": row, "created": False,
                "parent_ref": _folder_ref(ctx, account, parent_folder), "revision": 0,
                "warnings": [],
            }

        created = Folder(
            parent=parent_folder, name=clean_name, folder_class=MAIL_FOLDER_CLASS
        ).save()
        raw_id = folder_id(created)
        if raw_id is None:
            raise ToolError("upstream_error", "Exchange created a folder without an id.")
        rows, revision, warnings = _publish_mutation(ctx, account)
        row = row_for_folder(rows, raw_id)
        if row is None:
            raise ToolError("upstream_error", "Created folder is absent from the live tree.")
        return {
            "raw_id": raw_id, "row": row, "created": True,
            "parent_ref": _folder_ref(ctx, account, parent_folder),
            "revision": revision,
            "warnings": warnings,
        }

    payload = {"name": clean_name, "parent": parent}
    outcome = await _idempotent_create(
        ctx, "create_folder", idempotency_key, payload, work)
    if outcome["state"] != "completed":
        return {
            "ok": False,
            "error": outcome.get("error", {
                "code": "operation_in_progress",
                "message": "Folder create is still running; retry with the same key.",
            }),
            "operation": outcome["operation"],
        }
    created = outcome["result"]
    result: dict[str, Any] = {
        "ok": True,
        "created": bool(created["created"]),
        "folder": _public_folder(ctx, created["row"], created["parent_ref"]),
        "folder_revision": created.get("revision", 0),
    }
    if outcome.get("replayed"):
        result["idempotent_replay"] = True
    if created.get("warnings"):
        result["warnings"] = created["warnings"]
    return result


async def _rename_folder(ctx: Context, *, folder: str, new_name: str) -> dict[str, Any]:
    clean_name = _validated_name(new_name)

    def work(account: Any) -> dict[str, Any]:
        target = ctx.gateway.resolve_folder(account, folder, ctx.aliaser)
        _ensure_mutable(account, target)
        raw_id = folder_id(target)
        if raw_id is None:
            raise ToolError("not_found", "Folder has no Exchange id.")
        parent = getattr(target, "parent", None)
        if parent is None:
            raise ToolError("protected_folder", "Folder has no mutable parent.")
        before_rows = collect_folder_rows(account)
        before = row_for_folder(before_rows, raw_id)
        current_name = str(getattr(target, "name", "") or "")
        if current_name == clean_name:
            _store_snapshot(ctx, before_rows)
            return {"changed": False, "before": before, "after": before,
                    "parent_ref": _folder_ref(ctx, account, parent),
                    "revision": 0, "warnings": []}
        if child_named(parent, clean_name, exclude_id=raw_id) is not None:
            raise _conflict(clean_name)
        target.name = clean_name
        target.save(update_fields=["name"])
        rows, revision, warnings = _publish_mutation(ctx, account)
        return {"changed": True, "before": before,
                "after": row_for_folder(rows, raw_id),
                "parent_ref": _folder_ref(ctx, account, parent),
                "revision": revision, "warnings": warnings}

    result = await _gateway_write(ctx, work)
    if result["after"] is None:
        raise ToolError("upstream_error", "Renamed folder is absent from the live tree.")
    out: dict[str, Any] = {
        "ok": True, "changed": result["changed"],
        "folder": _public_folder(ctx, result["after"], result["parent_ref"]),
        "previous": _previous_folder(result["before"]),
        "folder_revision": result["revision"],
    }
    if result["warnings"]:
        out["warnings"] = result["warnings"]
    return out


async def _move_folder(ctx: Context, *, folder: str, parent: str) -> dict[str, Any]:
    def work(account: Any) -> dict[str, Any]:
        target = ctx.gateway.resolve_folder(account, folder, ctx.aliaser)
        destination = ctx.gateway.resolve_folder(account, parent, ctx.aliaser)
        _ensure_mutable(account, target)
        _ensure_mail_parent(account, destination)
        raw_id = folder_id(target)
        destination_id = folder_id(destination)
        if raw_id is None or destination_id is None:
            raise ToolError("not_found", "Folder or destination has no Exchange id.")
        if raw_id == destination_id or destination_id in descendant_ids(target):
            raise ToolError(
                "validation", "A folder cannot be moved into itself or one of its descendants.")
        before_rows = collect_folder_rows(account)
        before = row_for_folder(before_rows, raw_id)
        current_parent = getattr(target, "parent", None)
        if same_folder(current_parent, destination):
            _store_snapshot(ctx, before_rows)
            return {"changed": False, "before": before, "after": before,
                    "parent_ref": _folder_ref(ctx, account, destination),
                    "revision": 0, "warnings": []}
        if child_named(destination, str(getattr(target, "name", "") or ""),
                       exclude_id=raw_id) is not None:
            raise _conflict(str(getattr(target, "name", "") or ""))
        target.move(destination)
        rows, revision, warnings = _publish_mutation(ctx, account)
        return {"changed": True, "before": before,
                "after": row_for_folder(rows, raw_id),
                "parent_ref": _folder_ref(ctx, account, destination),
                "revision": revision, "warnings": warnings}

    result = await _gateway_write(ctx, work)
    if result["after"] is None:
        raise ToolError("upstream_error", "Moved folder is absent from the live tree.")
    out: dict[str, Any] = {
        "ok": True, "changed": result["changed"],
        "folder": _public_folder(ctx, result["after"], result["parent_ref"]),
        "previous": _previous_folder(result["before"]),
        "folder_revision": result["revision"],
    }
    if result["warnings"]:
        out["warnings"] = result["warnings"]
    return out


def _delete_snapshot(ctx: Context, account: Any, folder_ref: str,
                     disposition: str, allow_non_empty: bool) -> dict[str, Any]:
    if disposition not in ("trash", "permanent"):
        raise ToolError("validation", "disposition must be 'trash' or 'permanent'")
    target = ctx.gateway.resolve_folder(account, folder_ref, ctx.aliaser)
    _ensure_mutable(account, target)
    raw_id = folder_id(target)
    if raw_id is None:
        raise ToolError("not_found", "Folder has no Exchange id.")
    rows = collect_folder_rows(account)
    row = row_for_folder(rows, raw_id)
    if row is None:
        raise ToolError("not_found", "Folder is absent from the live mail-folder tree.")
    stats = subtree_stats(target)
    if (stats["items"] or stats["subfolders"]) and not allow_non_empty:
        raise ToolError(
            "validation",
            f"Folder is not empty ({stats['items']} items, "
            f"{stats['subfolders']} subfolders).",
            hint="Inspect the folder, then retry with allow_non_empty=true if intended.",
        )
    return {"target": target, "raw_id": raw_id, "row": row, "stats": stats,
            "disposition": disposition}


async def _delete_folder_preview(ctx: Context, kwargs: dict[str, Any]) -> dict[str, Any]:
    folder_ref = str(kwargs.get("folder") or "")
    disposition = str(kwargs.get("disposition") or "trash")
    allow_non_empty = bool(kwargs.get("allow_non_empty", False))

    def work(account: Any) -> dict[str, Any]:
        snap = _delete_snapshot(
            ctx, account, folder_ref, disposition, allow_non_empty)
        row, stats = snap["row"], snap["stats"]
        return {
            "action": "delete_folder",
            "folder": ctx.aliaser.alias_for(snap["raw_id"], "f"),
            "name": row["name"], "path": row["path"],
            "items": stats["items"], "subfolders": stats["subfolders"],
            "disposition": disposition,
            "recoverable": disposition == "trash",
            "_confirm_binding": {
                "id": snap["raw_id"],
                "changekey": getattr(snap["target"], "changekey", None),
                "path": row["path"], "stats": stats,
                "disposition": disposition,
                "allow_non_empty": allow_non_empty,
            },
        }

    return await ctx.gateway.call(work)


async def _delete_folder(ctx: Context, *, folder: str,
                         disposition: str = "trash",
                         allow_non_empty: bool = False) -> dict[str, Any]:
    def work(account: Any) -> dict[str, Any]:
        snap = _delete_snapshot(
            ctx, account, folder, disposition, allow_non_empty)
        row, stats, raw_id = snap["row"], snap["stats"], snap["raw_id"]
        delete_type = (MOVE_TO_DELETED_ITEMS if disposition == "trash"
                       else HARD_DELETE)
        snap["target"].delete(delete_type=delete_type)
        _rows, revision, warnings = _publish_mutation(ctx, account)
        return {"raw_id": raw_id, "row": row, "stats": stats,
                "revision": revision, "warnings": warnings}

    deleted = await _gateway_write(ctx, work)
    out: dict[str, Any] = {
        "ok": True,
        "deleted": {
            "id": ctx.aliaser.alias_for(deleted["raw_id"], "f"),
            "name": deleted["row"]["name"],
            "path": deleted["row"]["path"],
            **deleted["stats"],
            "disposition": disposition,
            "recoverable": disposition == "trash",
        },
        "folder_revision": deleted["revision"],
    }
    if deleted["warnings"]:
        out["warnings"] = deleted["warnings"]
    return out


def _obj(props: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    return {"type": "object", "properties": props,
            "required": required or [], "additionalProperties": False}


_FOLDER_REF = {
    "type": "string", "minLength": 1,
    "description": "Folder alias (f7/f:inbox/f:root), exact path, or raw id.",
}
_NAME = {
    "type": "string", "minLength": 1, "maxLength": MAX_FOLDER_NAME,
    "description": "One folder name; '/' and '\\' are not allowed.",
}


TOOLS: list[ToolSpec] = [
    ToolSpec(
        name="create_folder",
        description=(
            "Ensure a normal IPF.Note mail folder exists directly under parent. "
            "An exact case-insensitive sibling match is returned with created=false; "
            "optional idempotency_key makes timeout retries durable across processes."
        ),
        side_effect_class="write", requires_ews=True,
        input_schema=_obj({
            "name": dict(_NAME),
            "parent": {**_FOLDER_REF, "default": "f:root",
                       "description": _FOLDER_REF["description"]
                       + " Default: mailbox root."},
            "idempotency_key": {
                "type": "string", "minLength": 1, "maxLength": 200,
                "description": "Stable key for safe retries of this create.",
            },
        }, ["name"]),
        handler=_create_folder,
    ),
    ToolSpec(
        name="rename_folder",
        description=(
            "Rename one normal user mail folder. System folders are protected; "
            "a case-insensitive sibling name collision is refused."
        ),
        side_effect_class="write", requires_ews=True,
        input_schema=_obj({"folder": dict(_FOLDER_REF), "new_name": dict(_NAME)},
                          ["folder", "new_name"]),
        handler=_rename_folder,
    ),
    ToolSpec(
        name="move_folder",
        description=(
            "Move one normal user mail folder under another mail folder or f:root. "
            "Self/descendant moves, system targets and sibling collisions are refused."
        ),
        side_effect_class="write", requires_ews=True,
        input_schema=_obj({"folder": dict(_FOLDER_REF),
                           "parent": dict(_FOLDER_REF)}, ["folder", "parent"]),
        handler=_move_folder,
    ),
    ToolSpec(
        name="delete_folder",
        description=(
            "Delete one normal user mail folder. Default disposition=trash uses "
            "Exchange MoveToDeletedItems and is recoverable; permanent uses "
            "HardDelete. Non-empty trees require allow_non_empty=true. Always "
            "uses a content-bound two-phase confirmation."
        ),
        side_effect_class="destructive", requires_ews=True, confirm=True,
        input_schema=_obj({
            "folder": dict(_FOLDER_REF),
            "disposition": {
                "type": "string", "enum": ["trash", "permanent"],
                "default": "trash",
                "description": "Recoverable trash move (default) or HardDelete.",
            },
            "allow_non_empty": {
                "type": "boolean", "default": False,
                "description": "Explicitly acknowledge deleting contained items/subfolders.",
            },
        }, ["folder"]),
        handler=_delete_folder, preview=_delete_folder_preview,
    ),
]
