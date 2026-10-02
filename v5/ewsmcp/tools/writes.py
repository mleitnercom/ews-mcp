"""Tool pack: writes — the 12 mutating tools (DESIGN.md §Tools, §Safety).

Safety lives in the dispatcher (tools/base.py): kill-switch, tier gate,
recipient guard, two-phase confirm, send rate cap and alias resolution all
run *before* any handler here executes. Specs below only DECLARE
``side_effect_class`` and ``confirm``; handlers receive raw EWS ids and
never see ``confirm_token``.

The single exception (documented inline): ``create_event``/``update_event``
are class "write" for *tier* purposes (a no-invite event is a reversible
calendar write), but invitations/notifications LEAVE the org. The
dispatcher's kill-switch only covers class "send", so those two handlers
re-check ``settings.send_enabled`` themselves when (and only when) the call
would notify attendees.

exchangelib 5.0.3 pins honoured here (verified against the installed lib):
- ``Item.delete()`` takes NO disposal kwarg — it IS HardDelete. Reversible
  paths are ``move_to_trash()`` / ``soft_delete()``.
- ``OofSettings.internal_reply``/``external_reply`` are PLAIN strings (the
  4.x ``OofReply`` wrapper is gone); state constants are
  ``OofSettings.ENABLED/SCHEDULED/DISABLED``.
- ``Item.save(update_fields=[...], send_meeting_invitations=...)``;
  invitation constants live in ``exchangelib.items``.
- ``Message.create_reply(subject, body, ...)`` / ``create_reply_all`` /
  ``Item.create_forward(subject, body, to_recipients, ...)`` return reply
  items persisted via ``.save(account.drafts)`` (Exchange quotes the
  original server-side).
- ``Item.move(folder)`` returns ``None`` and updates ``item.id`` in place.
- There is NO follow-up-flag parameter anywhere: 5.0.3 has no first-class
  flag field; pretending would be mock-drift bait (categories instead).
- ``BaseReplyItem.save(folder)`` (reply/reply_all/forward) returns a
  ``BulkCreateResult``, NOT a ``Message``: ids and an attachment list, none
  of the Item API. Attaching to it raises AttributeError, so that path
  re-fetches the draft by id first (``_saved_draft_item``).
- ``Item.attach(att)`` on a SAVED item issues CreateAttachment immediately
  and writes the new root changekey back onto the parent in place. On an
  unsaved item it only appends locally — hence attachments are always
  applied AFTER ``.save()`` here (the v4 regression this restores was the
  mirror image: append + save, which never reached the server).
"""

import asyncio
import html as _html
import inspect
import logging
import mimetypes
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from exchangelib import CalendarItem, FileAttachment, HTMLBody, Message, OofSettings
from exchangelib.items import (
    SEND_TO_ALL_AND_SAVE_COPY,
    SEND_TO_CHANGED_AND_SAVE_COPY,
    SEND_TO_NONE,
)

from ..calendar_status import SHOW_AS_VALUES, to_ews_show_as, validate_show_as
from ..coordination import ProcessCoordinator
from ..dates import parse_when
from ..dto import fmt_dt
from ..errors import ToolError, map_exception
from .base import Context, ToolSpec

logger = logging.getLogger(__name__)

MAX_BULK_IDS = 50


async def _gateway_write(ctx: Context, fn: Any) -> Any:
    parameters = inspect.signature(ctx.gateway.call).parameters
    if "lane" in parameters:
        return await ctx.gateway.call(
            fn, lane="write", timeout_s=float(ctx.settings.request_timeout))
    return await ctx.gateway.call(fn)


def _coordinator(ctx: Context) -> ProcessCoordinator:
    if ctx.coordinator is None:
        ctx.coordinator = ProcessCoordinator(ctx.settings.data_dir,
                                             ctx.settings.ews_email)
    return ctx.coordinator


async def _idempotent_create(ctx: Context, kind: str, key: Optional[str],
                             payload: Dict[str, Any], fn: Any) -> Dict[str, Any]:
    """Run a create in the worker that owns its durable operation transition.

    This placement is deliberate: cancelling the MCP coroutine cannot cancel a
    Python thread already inside exchangelib, but that thread will still record
    ``completed`` and make a retry with the same key safe.
    """
    if not key:
        return {"state": "completed", "result": await _gateway_write(ctx, fn)}
    coord = _coordinator(ctx)
    op, created = coord.reserve_operation(kind, key, payload)
    if not created:
        ctx.audit.record(tool=kind, side_effect_class="write",
                         outcome=op["status"], latency_ms=0,
                         transport="operation-journal",
                         detail={"operation_id": op["operation_id"]})
        if op["status"] == "completed":
            return {"state": "completed", "result": op.get("result"),
                    "operation": op, "replayed": True}
        if op["status"] == "failed":
            raise ToolError(
                "operation_failed",
                op.get("error") or f"Prior {kind} operation failed",
                hint="Use a new idempotency_key only after verifying Exchange state.",
            )
        return {"state": op["status"], "operation": op}
    op_id = op["operation_id"]
    ctx.audit.record(tool=kind, side_effect_class="write", outcome="queued",
                     latency_ms=0, transport="operation-journal",
                     detail={"operation_id": op_id})

    def durable(account: Any) -> Any:
        started = time.time()
        ctx.audit.record(tool=kind, side_effect_class="write", outcome="running",
                         latency_ms=0, transport="operation-journal",
                         detail={"operation_id": op_id})
        try:
            return coord.execute_operation(op_id, lambda: fn(account))
        finally:
            current = coord.operation(op_id) or {"status": "unknown_outcome"}
            ctx.audit.record(
                tool=kind, side_effect_class="write", outcome=current["status"],
                latency_ms=int((time.time() - started) * 1000),
                transport="operation-journal", detail={"operation_id": op_id})

    try:
        result = await _gateway_write(ctx, durable)
        return {"state": "completed", "result": result,
                "operation": coord.operation(op_id)}
    except asyncio.CancelledError:
        coord.mark_unknown(op_id, "client cancelled while Exchange call was running")
        raise
    except ToolError as exc:
        coord.mark_unknown(op_id, f"{exc.code}: {exc.message}")
        current = coord.operation(op_id) or op
        return {"state": current["status"], "operation": current,
                "error": exc.to_dict()["error"]}
    except Exception as exc:
        mapped = map_exception(exc)
        current = coord.operation(op_id) or op
        return {"state": current["status"], "operation": current,
                "error": mapped.to_dict()["error"]}


def _write_through(ctx: Context, method: str, *args: Any) -> None:
    """Patch the mirror after a successful EWS mutation. Cache failures are
    logged and swallowed — the delta sync repairs the mirror within one
    cycle, and a cache hiccup must never fail a write that already
    succeeded upstream."""
    cache = getattr(ctx, "cache", None)
    if cache is None:
        return
    try:
        getattr(cache, method)(*args)
    except Exception as exc:
        logger.warning("write-through %s failed (sync will repair): %s",
                       method, exc)
_DRAFT_MODES = ("new", "reply", "reply_all", "forward")
_RESPONSE_METHOD = {"accept": "accept", "tentative": "tentatively_accept", "decline": "decline"}
_OOF_STATE = {
    "disabled": OofSettings.DISABLED,
    "enabled": OofSettings.ENABLED,
    "scheduled": OofSettings.SCHEDULED,
}

# --- send_draft idempotency (DESIGN.md §Safety, Stripe semantics) ---------------
# Bounded: entries expire after 48h and the store caps at 512 keys (oldest
# evicted first) so a long-lived process can't grow it without limit.
_IDEMPOTENT_SENDS: Dict[str, Dict[str, Any]] = {}
_IDEMPOTENT_LOCK = threading.Lock()
_IDEMPOTENCY_TTL_S = 48 * 3600
_IDEMPOTENCY_MAX = 512


def reset_idempotency_store() -> None:
    """Tests only."""
    with _IDEMPOTENT_LOCK:
        _IDEMPOTENT_SENDS.clear()


def _idempotency_get(key: str) -> Optional[Dict[str, Any]]:
    now = time.time()
    with _IDEMPOTENT_LOCK:
        for stale in [k for k, v in _IDEMPOTENT_SENDS.items()
                      if now - v.get("ts", 0) > _IDEMPOTENCY_TTL_S]:
            del _IDEMPOTENT_SENDS[stale]
        return _IDEMPOTENT_SENDS.get(key)


def _idempotency_put(key: str, record: Dict[str, Any]) -> None:
    with _IDEMPOTENT_LOCK:
        while len(_IDEMPOTENT_SENDS) >= _IDEMPOTENCY_MAX:
            oldest = min(_IDEMPOTENT_SENDS, key=lambda k: _IDEMPOTENT_SENDS[k].get("ts", 0))
            del _IDEMPOTENT_SENDS[oldest]
        _IDEMPOTENT_SENDS[key] = record


# --- shared helpers ----------------------------------------------------------


def _wrap_html(text: str) -> HTMLBody:
    """Plain text → minimal HTML (<p> per blank-line paragraph, <br/> inside)."""
    paragraphs = (text or "").split("\n\n")
    rendered = "".join(
        "<p>" + _html.escape(p).replace("\n", "<br/>") + "</p>" for p in paragraphs
    ) or "<p></p>"
    return HTMLBody(f"<html><body>{rendered}</body></html>")


def _email_list(recipients: Any) -> List[str]:
    out: List[str] = []
    for r in recipients or []:
        addr = getattr(r, "email_address", None)
        out.append(addr if isinstance(addr, str) else str(r))
    return out


def _own_address(account: Any) -> str:
    return str(getattr(account, "primary_smtp_address", "") or "").lower()


def _drop_self(recipients: Any, own: str) -> List[Any]:
    """Recipients minus the mailbox owner, comparison case-insensitive."""
    kept: List[Any] = []
    for r in recipients or []:
        addr = getattr(r, "email_address", None) or str(r)
        if own and isinstance(addr, str) and addr.lower() == own:
            continue
        kept.append(r)
    return kept


def _prefixed(prefix: str, original_subject: Any) -> str:
    base = original_subject if isinstance(original_subject, str) else ""
    if base.lower().startswith(prefix.strip().lower()):
        return base
    return f"{prefix}{base}"


def _attachment_root_error(path: Path, roots: List[Path]) -> ToolError:
    listed = ", ".join(str(r) for r in roots)
    return ToolError(
        "validation",
        f"attachment path is outside the allowed roots: {path}",
        hint=f"Allowed roots: {listed}. Add more via ATTACH_ALLOWED_DIRS "
             "(os.pathsep-separated) or move the file under one of them.",
    )


def _resolve_attachment_path(raw: str, roots: List[Path]) -> Path:
    """Resolve one caller-supplied path and confine it to the allowlist.

    ``resolve()`` collapses ``..`` and follows symlinks BEFORE the containment
    check, so neither traversal nor a symlink planted inside an allowed root
    can reach outside it."""
    if not isinstance(raw, str) or not raw.strip():
        raise ToolError("validation", "attachment paths must be non-empty strings")
    try:
        path = Path(raw).expanduser().resolve()
    except OSError as exc:
        raise ToolError("validation", f"cannot resolve attachment path {raw!r}: {exc}")
    if not any(path == root or root in path.parents for root in roots):
        raise _attachment_root_error(path, roots)
    if not path.exists():
        raise ToolError("not_found", f"attachment file does not exist: {path}")
    if not path.is_file():
        raise ToolError("validation", f"attachment path is not a file: {path}")
    return path


def _load_attachments(settings: Any, paths: List[str]) -> List[FileAttachment]:
    """Read the files into FileAttachment objects, enforcing count and size
    caps BEFORE anything touches EWS."""
    if len(paths) > settings.attach_max_files:
        raise ToolError(
            "validation",
            f"at most {settings.attach_max_files} attachments per draft "
            f"(got {len(paths)})",
        )
    roots = settings.attachment_roots()
    resolved = [_resolve_attachment_path(p, roots) for p in paths]
    total = 0
    for path in resolved:
        total += path.stat().st_size
        if total > settings.attach_max_bytes:
            raise ToolError(
                "validation",
                f"attachments exceed the {settings.attach_max_bytes} byte cap "
                f"for one draft",
                hint="Raise ATTACH_MAX_BYTES, or send fewer/smaller files. "
                     "Exchange enforces its own message size limit on top.",
            )
    out: List[FileAttachment] = []
    for path in resolved:
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        out.append(FileAttachment(name=path.name,
                                  content=path.read_bytes(),
                                  content_type=content_type))
    return out


def _attach_all(account: Any, msg: Any, atts: List[FileAttachment]) -> List[Dict[str, Any]]:
    """Attach to an ALREADY SAVED draft — each attach() is one CreateAttachment
    round trip that rolls the parent changekey forward in place.

    ``msg`` MUST be a real ``Item``. The reply/forward path cannot pass the
    result of ``BaseReplyItem.save()`` here: that returns a
    ``BulkCreateResult``, a BaseItem carrying ids only, with no ``attach()``
    — see ``_saved_draft_item``.

    On failure the half-built draft is moved to trash: a call that reports an
    error must not leave an orphan behind for an unattended run to accumulate.
    """
    applied: List[Dict[str, Any]] = []
    try:
        for att in atts:
            msg.attach(att)
            applied.append({"name": att.name, "size_bytes": len(att.content or b"")})
    except Exception:
        try:
            msg.move_to_trash()  # reversible — never HardDelete on an error path
        except Exception:
            logger.warning("could not discard the draft after a failed attach",
                           exc_info=True)
        raise
    return applied


def _saved_draft_item(account: Any, saved: Any) -> Any:
    """Re-fetch a reply/forward draft as a real Item.

    ``Message.create_reply(...).save(folder)`` returns a ``BulkCreateResult``
    ("a dummy class to store return values"), not a ``Message``. It carries
    id/changekey/attachments but none of the Item API, so attaching to it
    raises AttributeError. Fetching by id yields the genuine item that
    ``update_draft`` already attaches to successfully.
    """
    return _fetch_one(account, str(getattr(saved, "id", saved)))


def _fetch_one(account: Any, raw_id: str) -> Any:
    """account.fetch yields items in input order; missing ids surface as
    in-line exception instances (exchangelib bulk semantics)."""
    fetched = list(account.fetch([(raw_id, None)]))
    if not fetched:
        raise ToolError("not_found", "item not found; the id may be stale",
                        hint="Re-run search_messages/list_events for a fresh id.")
    item = fetched[0]
    if isinstance(item, Exception):
        raise item  # map_exception classifies (ErrorItemNotFound → not_found)
    return item


def _fetch_many(account: Any, raw_ids: List[str],
                only: Optional[List[str]] = None) -> List[Any]:
    """Bulk fetch with a projection: without only_fields, exchangelib pulls
    the FULL item — including base64 MIME — so a 'mark 50 read' on messages
    with attachments used to move megabytes to flip one boolean each."""
    return list(account.fetch([(i, None) for i in raw_ids], only_fields=only))


def _require_bulk(ids: List[str]) -> None:
    if not ids:
        raise ToolError("validation", "ids must contain at least one id")
    if len(ids) > MAX_BULK_IDS:
        raise ToolError("validation",
                        f"at most {MAX_BULK_IDS} ids per call (got {len(ids)})")


def _item_folder_id(item: Any) -> Optional[str]:
    parent = getattr(item, "parent_folder_id", None)
    if parent is not None and getattr(parent, "id", None):
        return parent.id
    return getattr(getattr(item, "folder", None), "id", None)


def _require_in_drafts(item: Any, account: Any, tool: str) -> None:
    """Input validation mandated by the spec (#15, #18) — not a safety gate."""
    drafts_id = getattr(account.drafts, "id", None)
    if drafts_id is None or _item_folder_id(item) != drafts_id:
        raise ToolError(
            "validation",
            f"{tool}: the item is not a draft (it is not in f:drafts)",
            hint="Only items in the Drafts folder are valid here; "
                 "use create_draft to author one.",
        )


def _failed_entry(ctx: Context, raw_id: str, exc: Exception) -> Dict[str, str]:
    return {"id": ctx.aliaser.alias_for(raw_id, "m"),
            "error": f"{type(exc).__name__}: {exc}"}


def _draft_preview(to: List[str], cc: List[str], subject: str,
                   body: Optional[str]) -> Dict[str, Any]:
    return {"to": to, "cc": cc, "subject": subject,
            "body_snippet": (body or "")[:200]}


# --- WRITE class (reversible; tier ≥ draft) ----------------------------------


async def _create_draft(ctx: Context, *, body: str, mode: str = "new",
                        reply_to: Optional[str] = None,
                        to: Optional[List[str]] = None,
                        cc: Optional[List[str]] = None,
                        bcc: Optional[List[str]] = None,
                        subject: Optional[str] = None,
                        importance: Optional[str] = None,
                        attachments: Optional[List[str]] = None,
                        idempotency_key: Optional[str] = None) -> Dict[str, Any]:
    if mode not in _DRAFT_MODES:
        raise ToolError("validation", f"mode must be one of {_DRAFT_MODES}, got {mode!r}")
    if mode != "new" and not reply_to:
        raise ToolError("validation", f"reply_to (a message id) is required for mode={mode!r}")
    if mode in ("new", "forward") and not to:
        raise ToolError("validation", f"to is required for mode={mode!r}")
    if importance not in (None, "normal", "high"):
        raise ToolError("validation", "importance must be 'normal' or 'high'")
    html_body = _wrap_html(body)
    # Read and validate the files up front: a bad path must fail before a
    # draft exists, not leave a half-populated one behind.
    atts = _load_attachments(ctx.settings, attachments) if attachments else []

    def work(account: Any) -> Dict[str, Any]:
        if mode == "new":
            msg = Message(
                account=account, folder=account.drafts,
                subject=subject or "", body=html_body,
                to_recipients=list(to),
                cc_recipients=list(cc) if cc else None,
                bcc_recipients=list(bcc) if bcc else None,
            )
            if importance == "high":
                msg.importance = "High"
            msg.save()  # SAVE_ONLY into Drafts — nothing leaves the mailbox
            applied = _attach_all(account, msg, atts) if atts else []
            return {"raw_id": str(msg.id), "changekey": getattr(msg, "changekey", None),
                    "to": list(to), "cc": list(cc or []), "subject": subject or "",
                    "attachments": applied}
        original = _fetch_one(account, reply_to)
        extra: Dict[str, Any] = {}
        fallback = None
        if mode == "reply":
            final_subject = subject or _prefixed("Re: ", getattr(original, "subject", None))
            if to:
                extra["to_recipients"] = list(to)
            if cc:
                extra["cc_recipients"] = list(cc)
            if bcc:
                extra["bcc_recipients"] = list(bcc)
            method = getattr(original, "create_reply", None)
            reply = method(final_subject, html_body, **extra) if callable(method) else None
        elif mode == "reply_all":
            final_subject = subject or _prefixed("Re: ", getattr(original, "subject", None))
            method = getattr(original, "create_reply_all", None)
            reply = method(final_subject, html_body) if callable(method) else None
        else:  # forward — create_forward REQUIRES to_recipients (5.0.3)
            final_subject = subject or _prefixed("Fwd: ", getattr(original, "subject", None))
            reply = original.create_forward(
                final_subject, html_body, list(to),
                cc_recipients=list(cc) if cc else None,
                bcc_recipients=list(bcc) if bcc else None,
            )
        if mode in ("reply", "reply_all") and reply is None:
            # MeetingResponse has no create_reply/create_reply_all in
            # exchangelib 5.0.3. Reconstruct a transparent draft instead of
            # leaking AttributeError. The sender is the primary reply target;
            # reply_all additionally retains original To/Cc recipients except
            # the mailbox itself. Explicit caller recipients still win.
            sender = getattr(original, "sender", None)
            sender_address = getattr(sender, "email_address", None)
            recipients = list(to or ([] if not sender_address else [sender_address]))
            copied = list(cc or [])
            if mode == "reply_all" and not to:
                own = str(getattr(account, "primary_smtp_address", "") or "").lower()
                for recipient in [*(getattr(original, "to_recipients", None) or []),
                                  *(getattr(original, "cc_recipients", None) or [])]:
                    address = getattr(recipient, "email_address", None) or str(recipient)
                    if address and address.lower() != own and address not in recipients + copied:
                        copied.append(address)
            if not recipients and not copied:
                raise ToolError(
                    "unsupported_item_type",
                    f"Cannot reconstruct {mode} recipients for "
                    f"{type(original).__name__}",
                    hint="Pass explicit to/cc recipients or reply from Outlook.",
                )
            reply = Message(
                account=account, folder=account.drafts, subject=final_subject,
                body=html_body, to_recipients=recipients,
                cc_recipients=copied or None,
                bcc_recipients=list(bcc) if bcc else None,
            )
            fallback = f"reconstructed_{type(original).__name__}"

        if mode == "reply_all" and fallback is None:
            # create_reply_all mirrors the original To/Cc verbatim, so the
            # mailbox owner — who was a recipient of the original — ends up
            # addressed in their own reply. Outlook drops the owner here; do
            # the same. Guarded: if the owner is the ONLY party (a self-sent
            # message), keep Exchange's recipients rather than save a draft
            # that can reach nobody.
            own = _own_address(account)
            pruned_to = _drop_self(getattr(reply, "to_recipients", None), own)
            pruned_cc = _drop_self(getattr(reply, "cc_recipients", None), own)
            if pruned_to or pruned_cc:
                reply.to_recipients = pruned_to or None
                reply.cc_recipients = pruned_cc or None

        # Exchange quotes the original server-side, keeping threading headers.
        if fallback:
            reply.save()
            saved = reply
        else:
            saved = reply.save(account.drafts)
        applied = []
        if atts:
            # NOT saved.attach(...): save() hands back a BulkCreateResult.
            item = _saved_draft_item(account, saved)
            applied = _attach_all(account, item, atts)
            saved = item  # carries the changekey attach() just rolled forward
        prev_to = _email_list(getattr(reply, "to_recipients", None)) or list(to or [])
        prev_cc = _email_list(getattr(reply, "cc_recipients", None)) or list(cc or [])
        return {"raw_id": str(saved.id), "changekey": getattr(saved, "changekey", None),
                "to": prev_to, "cc": prev_cc, "subject": final_subject,
                "attachments": applied, "fallback": fallback}

    payload = {"mode": mode, "reply_to": reply_to, "to": to, "cc": cc, "bcc": bcc,
               "subject": subject, "body": body, "importance": importance,
               "attachments": attachments}
    outcome = await _idempotent_create(ctx, "create_draft", idempotency_key,
                                       payload, work)
    if outcome["state"] != "completed":
        return {"ok": False, "error": outcome.get("error", {
                    "code": "operation_in_progress",
                    "message": "Draft create is still running; retry with the same key.",
                }), "operation": outcome["operation"]}
    created = outcome["result"]
    alias = ctx.aliaser.alias_for(str(created["raw_id"]), "d",
                                  changekey=created.get("changekey"))
    preview = _draft_preview(created.get("to") or [], created.get("cc") or [],
                             created.get("subject") or "", body)
    applied = created.get("attachments") or []
    if applied:
        preview["attachments"] = applied
    result = {
        "ok": True,
        "draft_id": alias,
        "folder": "f:drafts",
        "preview": preview,
        "note": "saved as draft — NOT sent",
    }
    if created.get("fallback"):
        result["fallback"] = created["fallback"]
        result["note"] += "; recipients reconstructed because the source item has no reply API"
    if outcome.get("operation"):
        result["operation"] = outcome["operation"]
    if outcome.get("replayed"):
        result["idempotent_replay"] = True
    return result


async def _update_draft(ctx: Context, *, draft_id: str,
                        to: Optional[List[str]] = None,
                        cc: Optional[List[str]] = None,
                        bcc: Optional[List[str]] = None,
                        subject: Optional[str] = None,
                        body: Optional[str] = None,
                        attachments: Optional[List[str]] = None) -> Dict[str, Any]:
    if (to is None and cc is None and bcc is None and subject is None
            and body is None and not attachments):
        raise ToolError(
            "validation",
            "nothing to update: pass to/cc/bcc/subject/body/attachments")
    atts = _load_attachments(ctx.settings, attachments) if attachments else []

    def work(account: Any) -> Dict[str, Any]:
        msg = _fetch_one(account, draft_id)
        changed: List[str] = []  # update_fields uses exchangelib field names
        for value, field in ((to, "to_recipients"), (cc, "cc_recipients"),
                             (bcc, "bcc_recipients")):
            if value is not None:
                setattr(msg, field, list(value))
                changed.append(field)
        if subject is not None:
            msg.subject = subject
            changed.append("subject")
        if body is not None:
            msg.body = _wrap_html(body)
            changed.append("body")
        if changed:  # attachments-only calls must not issue an empty update
            msg.save(update_fields=changed)
        # AFTER save: attach() rolls the changekey forward, so a save that
        # followed it would collide with its own stale copy.
        applied = _attach_all(account, msg, atts) if atts else []
        prev_subject = msg.subject if isinstance(msg.subject, str) else None
        preview = {
            "to": _email_list(getattr(msg, "to_recipients", None)),
            "cc": _email_list(getattr(msg, "cc_recipients", None)),
            "subject": prev_subject,
            "body_snippet": (body or "")[:200] if body is not None else None,
        }
        if applied:
            preview["attachments_added"] = applied
        return preview

    preview = await _gateway_write(ctx, work)
    return {"ok": True, "draft_id": ctx.aliaser.alias_for(draft_id, "d"),
            "preview": preview}


async def _delete_draft(ctx: Context, *, draft_id: str) -> Dict[str, Any]:
    def work(account: Any) -> Dict[str, Any]:
        msg = _fetch_one(account, draft_id)
        _require_in_drafts(msg, account, "delete_draft")
        msg.move_to_trash()  # reversible — hence write class, no confirm
        return {"deleted": True}

    return await _gateway_write(ctx, work)


async def _update_messages(ctx: Context, *, ids: List[str],
                           set_read: Optional[bool] = None,
                           categories_add: Optional[List[str]] = None,
                           categories_remove: Optional[List[str]] = None) -> Dict[str, Any]:
    # NOTE: there is deliberately NO set_flag parameter — exchangelib 5.0.3
    # exposes no first-class follow-up flag field, and pretending would be
    # mock-drift bait. Categories are the visible marker (see description).
    _require_bulk(ids)
    if set_read is None and not categories_add and not categories_remove:
        raise ToolError("validation",
                        "nothing to update: pass set_read and/or categories_add/_remove")
    remove = set(categories_remove or [])
    ok_ids: List[str] = []
    final_cats: List[tuple] = []

    def work(account: Any) -> Dict[str, Any]:
        updated, failed = 0, []
        fetched = _fetch_many(account, ids,
                              only=["id", "changekey", "is_read", "categories"])
        for raw_id, item in zip(ids, fetched):
            try:  # per-item isolation: one failure never aborts the batch
                if isinstance(item, Exception):
                    raise item
                changed: List[str] = []
                if set_read is not None:
                    item.is_read = bool(set_read)
                    changed.append("is_read")
                if categories_add or remove:
                    cats = [c for c in (item.categories or []) if c not in remove]
                    for c in categories_add or []:
                        if c not in cats:
                            cats.append(c)
                    item.categories = cats or None
                    changed.append("categories")
                    final_cats.append((raw_id, cats))
                item.save(update_fields=changed)
                updated += 1
                ok_ids.append(raw_id)
            except Exception as exc:
                failed.append(_failed_entry(ctx, raw_id, exc))
        return {"updated": updated, "failed": failed}

    result = await _gateway_write(ctx, work)
    if ok_ids and set_read is not None:
        _write_through(ctx, "set_read_flag", ok_ids, bool(set_read))
    for raw_id, cats in final_cats:
        _write_through(ctx, "apply_categories", raw_id, cats)
    return result


async def _move_messages(ctx: Context, *, ids: List[str],
                         to_folder: str) -> Dict[str, Any]:
    _require_bulk(ids)

    moved_old_ids: List[str] = []

    def work(account: Any) -> Dict[str, Any]:
        folder = ctx.gateway.resolve_folder(account, to_folder, ctx.aliaser)
        moved, failed = [], []
        fetched = _fetch_many(account, ids, only=["id", "changekey"])
        for raw_id, item in zip(ids, fetched):
            try:
                if isinstance(item, Exception):
                    raise item
                item.move(folder)  # returns None; updates item.id in place (5.0.3)
                # 5.0.3 may leave item.id None when the response carries no
                # new id — never crash on it; fall back to the old binding.
                new_raw = str(item.id) if getattr(item, "id", None) else None
                alias = None
                if new_raw and new_raw != raw_id:
                    # EWS ids change on move — rebind so the model's alias
                    # keeps working (DESIGN.md §Ids).
                    alias = ctx.aliaser.rebind(raw_id, new_raw)
                if alias is None:
                    alias = ctx.aliaser.alias_for(new_raw or raw_id, "m")
                row: Dict[str, Any] = {"id": alias}
                if new_raw is None:
                    row["note"] = ("moved, but Exchange returned no new id — "
                                   "re-search before reusing this id")
                moved.append(row)
                moved_old_ids.append(raw_id)
            except Exception as exc:
                failed.append(_failed_entry(ctx, raw_id, exc))
        return {"moved": len(moved), "items": moved, "failed": failed}

    result = await _gateway_write(ctx, work)
    # Tombstone the moved rows: their old ids are gone from the source
    # folder; the delta sync re-adds them under the new id if the target
    # folder is mirrored.
    _write_through(ctx, "tombstone_messages", moved_old_ids)
    return result


async def _create_event(ctx: Context, *, subject: str, start: str, end: str,
                        attendees: Optional[List[str]] = None,
                        location: Optional[str] = None,
                        body: Optional[str] = None,
                        send_invitations: bool = False,
                        show_as: Optional[str] = None,
                        idempotency_key: Optional[str] = None) -> Dict[str, Any]:
    # The ONE handler-side safety check in this pack, by design: this tool is
    # class "write" so a no-invite event stays available at draft tier, but
    # the dispatcher kill-switch only covers class "send". Invitations leave
    # the org, so the same switch must hold here. (Confirm is already gated
    # by the dispatcher via confirm=lambda kw: send_invitations.)
    if send_invitations and not ctx.settings.send_enabled:
        raise ToolError(
            "kill_switch",
            "create_event with send_invitations=true is blocked: "
            "SEND_ENABLED=false on this server.",
            hint="Save the event without invitations, or have the operator "
                 "flip SEND_ENABLED.",
        )
    start_dt = parse_when(start, "start", ctx.settings.ews_tz)
    end_dt = parse_when(end, "end", ctx.settings.ews_tz)
    if end_dt <= start_dt:
        raise ToolError("validation", "end must be after start")
    normalized_show_as = validate_show_as(show_as)

    def work(account: Any) -> Dict[str, Any]:
        item = CalendarItem(account=account, folder=account.calendar,
                            subject=subject, start=start_dt, end=end_dt)
        if attendees:
            # AttendeesField coerces plain email strings to Attendee (5.0.3).
            item.required_attendees = list(attendees)
        if location:
            item.location = location
        if body is not None:
            item.body = _wrap_html(body)
        if normalized_show_as is not None:
            item.legacy_free_busy_status = to_ews_show_as(normalized_show_as)
        item.save(send_meeting_invitations=SEND_TO_ALL_AND_SAVE_COPY
                  if send_invitations else SEND_TO_NONE)
        return {"raw_id": str(item.id), "show_as": normalized_show_as}

    payload = {"subject": subject, "start": start, "end": end,
               "attendees": attendees, "location": location, "body": body,
               "send_invitations": send_invitations, "show_as": normalized_show_as}
    outcome = await _idempotent_create(ctx, "create_event", idempotency_key,
                                       payload, work)
    if outcome["state"] != "completed":
        return {"ok": False, "error": outcome.get("error", {
                    "code": "operation_in_progress",
                    "message": "Event create is still running; retry with the same key.",
                }), "operation": outcome["operation"]}
    created = outcome["result"]
    result = {"ok": True,
              "event_id": ctx.aliaser.alias_for(created["raw_id"], "e"),
              "invitations_sent": bool(send_invitations)}
    if created.get("show_as"):
        result["show_as"] = created["show_as"]
    if outcome.get("operation"):
        result["operation"] = outcome["operation"]
    if outcome.get("replayed"):
        result["idempotent_replay"] = True
    _write_through(ctx, "upsert_events", [{
        "ews_id": created["raw_id"], "changekey": None, "subject": subject,
        "start_ts": int(start_dt.timestamp()),
        "start_iso": fmt_dt(start_dt, ctx.settings.ews_tz),
        "end_ts": int(end_dt.timestamp()),
        "end_iso": fmt_dt(end_dt, ctx.settings.ews_tz),
        "location": location, "organizer": None, "is_recurring": 0,
        "my_response": None, "show_as": normalized_show_as,
    }])
    return result


async def _update_event(ctx: Context, *, event_id: str,
                        subject: Optional[str] = None,
                        start: Optional[str] = None,
                        end: Optional[str] = None,
                        location: Optional[str] = None,
                        body: Optional[str] = None,
                        show_as: Optional[str] = None,
                        notify_attendees: bool = False) -> Dict[str, Any]:
    # Same documented exception as create_event: write-class for tier, but
    # notifications leave the org → the send kill-switch applies.
    if notify_attendees and not ctx.settings.send_enabled:
        raise ToolError(
            "kill_switch",
            "update_event with notify_attendees=true is blocked: "
            "SEND_ENABLED=false on this server.",
            hint="Update without notifying, or have the operator flip SEND_ENABLED.",
        )
    if all(v is None for v in (subject, start, end, location, body, show_as)):
        raise ToolError("validation",
                        "nothing to update: pass subject/start/end/location/body/show_as")
    normalized_show_as = validate_show_as(show_as)
    start_dt = parse_when(start, "start", ctx.settings.ews_tz) if start else None
    end_dt = parse_when(end, "end", ctx.settings.ews_tz) if end else None

    def work(account: Any) -> Dict[str, Any]:
        item = _fetch_one(account, event_id)
        changed: List[str] = []
        for value, field in ((subject, "subject"), (start_dt, "start"),
                             (end_dt, "end"), (location, "location")):
            if value is not None:
                setattr(item, field, value)
                changed.append(field)
        if body is not None:
            item.body = _wrap_html(body)
            changed.append("body")
        if normalized_show_as is not None:
            item.legacy_free_busy_status = to_ews_show_as(normalized_show_as)
            changed.append("legacy_free_busy_status")
        item.save(update_fields=changed,
                  send_meeting_invitations=SEND_TO_CHANGED_AND_SAVE_COPY
                  if notify_attendees else SEND_TO_NONE)
        return {"updated_fields": changed, "attendees_notified": bool(notify_attendees)}

    result = await _gateway_write(ctx, work)
    cache_fields: Dict[str, Any] = {}
    if subject is not None:
        cache_fields["subject"] = subject
    if start_dt is not None:
        cache_fields.update({
            "start_ts": int(start_dt.timestamp()),
            "start_iso": fmt_dt(start_dt, ctx.settings.ews_tz),
        })
    if end_dt is not None:
        cache_fields.update({
            "end_ts": int(end_dt.timestamp()),
            "end_iso": fmt_dt(end_dt, ctx.settings.ews_tz),
        })
    if location is not None:
        cache_fields["location"] = location
    if normalized_show_as is not None:
        cache_fields["show_as"] = normalized_show_as
    _write_through(ctx, "patch_event", event_id, cache_fields)
    return result


# --- SEND class (leaves the mailbox; tier full, two-phase, kill-switch) ------


def _draft_content(msg: Any) -> Dict[str, Any]:
    """The draft fields the confirm token binds and the preview shows.

    Recipients are sorted so reordering doesn't invalidate a token, but
    adding/removing one does (a content change worth re-previewing).
    """
    subject = getattr(msg, "subject", None)
    body_text = getattr(msg, "text_body", None)
    if not isinstance(body_text, str):
        body = getattr(msg, "body", None)
        body_text = str(body) if body is not None else ""
    return {
        "subject": subject if isinstance(subject, str) else "",
        "to": sorted(_email_list(getattr(msg, "to_recipients", None))),
        "cc": sorted(_email_list(getattr(msg, "cc_recipients", None))),
        "bcc": sorted(_email_list(getattr(msg, "bcc_recipients", None))),
        "body_text": body_text,
        "attachment_count": len(getattr(msg, "attachments", None) or []),
    }


async def _send_draft_preview(ctx: Context, kwargs: Dict[str, Any]) -> Dict[str, Any]:
    """Confirm-gate hook: fetch the draft's CURRENT content (both phases).

    Phase 1 hashes it into the token and shows it as the preview; phase 2
    refetches and the hash mismatch kills the token if the draft changed in
    between (update_draft TOCTOU) — ported from the v3.5 implementation.
    """
    draft_id = kwargs.get("draft_id")
    if not isinstance(draft_id, str) or not draft_id:
        raise ToolError("validation", "draft_id is required")
    try:
        raw_id = ctx.aliaser.resolve(draft_id)
    except KeyError as e:
        raise ToolError("validation", str(e.args[0] if e.args else e))

    def work(account: Any) -> Dict[str, Any]:
        msg = _fetch_one(account, raw_id)
        _require_in_drafts(msg, account, "send_draft")
        return _draft_content(msg)

    content = await _gateway_write(ctx, work)
    content["draft_id"] = draft_id  # echo the caller's handle in the preview
    return content


def _send_confirm_needed(kwargs: Dict[str, Any]) -> bool:
    """Skip the confirm gate on an idempotent REPLAY (v3.5 semantics).

    A retry with the same idempotency_key + draft_id returns the cached
    receipt without sending anything, so demanding a fresh (single-use)
    token there would break exactly the retry-after-timeout case the key
    exists for. A mismatched draft_id still goes through the gate and is
    rejected by the handler.
    """
    key = kwargs.get("idempotency_key")
    if isinstance(key, str) and key:
        prior = _idempotency_get(key)
        if prior is not None and prior.get("draft_id") == kwargs.get("draft_id"):
            return False
    return True


async def _send_draft(ctx: Context, *, draft_id: str,
                      idempotency_key: Optional[str] = None) -> Dict[str, Any]:
    if idempotency_key:
        prior = _idempotency_get(idempotency_key)
        if prior is not None:
            if prior["draft_id"] != draft_id:
                raise ToolError(
                    "validation",
                    "idempotency_key was already used for a different draft",
                    hint="Use a fresh idempotency_key per distinct send.",
                )
            return {**prior["result"], "replayed": True}

    def work(account: Any) -> Dict[str, Any]:
        msg = _fetch_one(account, draft_id)
        _require_in_drafts(msg, account, "send_draft")
        # Draft with an id → SendItem; Exchange moves it from Drafts to Sent.
        msg.send(save_copy=True)
        imid = getattr(msg, "message_id", None)
        return {"sent": True,
                "internet_message_id": imid if isinstance(imid, str) else None}

    result = await _gateway_write(ctx, work)
    if idempotency_key:
        _idempotency_put(idempotency_key, {"draft_id": draft_id,
                                           "result": dict(result),
                                           "ts": time.time()})
    return result


async def _respond_to_event(ctx: Context, *, event_id: str, response: str,
                            message: Optional[str] = None) -> Dict[str, Any]:
    method = _RESPONSE_METHOD.get(response)
    if method is None:
        raise ToolError("validation",
                        f"response must be one of {sorted(_RESPONSE_METHOD)}, got {response!r}")

    def work(account: Any) -> Dict[str, Any]:
        item = _fetch_one(account, event_id)
        kwargs = {"body": message} if message else {}
        # 5.0.3: accept()/tentatively_accept()/decline() build the meeting
        # reply item and send it to the organizer.
        getattr(item, method)(**kwargs)
        return {"responded": response}

    return await _gateway_write(ctx, work)


async def _set_oof(ctx: Context, *, state: str,
                   internal_reply: Optional[str] = None,
                   external_reply: Optional[str] = None,
                   start: Optional[str] = None,
                   end: Optional[str] = None) -> Dict[str, Any]:
    target_state = _OOF_STATE.get(state)
    if target_state is None:
        raise ToolError("validation",
                        f"state must be one of {sorted(_OOF_STATE)}, got {state!r}")
    if state != "disabled" and not internal_reply:
        raise ToolError("validation", f"internal_reply is required when state={state!r}")
    if state == "scheduled" and not (start and end):
        raise ToolError("validation", "start and end are required when state='scheduled'")
    start_dt = parse_when(start, "start", ctx.settings.ews_tz) if start else None
    end_dt = parse_when(end, "end", ctx.settings.ews_tz) if end else None

    def work(account: Any) -> Dict[str, Any]:
        kwargs: Dict[str, Any] = {"state": target_state}
        if internal_reply is not None:
            # 5.0.3: plain strings — the OofReply wrapper class is gone.
            kwargs["internal_reply"] = internal_reply
            # OofSettings.clean requires BOTH replies when state != disabled.
            kwargs["external_reply"] = external_reply or internal_reply
        if start_dt is not None:
            kwargs["start"] = start_dt
        if end_dt is not None:
            kwargs["end"] = end_dt
        account.oof_settings = OofSettings(**kwargs)
        return {"state": state}

    return await _gateway_write(ctx, work)


# --- DESTRUCTIVE class (tier full) -------------------------------------------


async def _delete_event(ctx: Context, *, event_id: str) -> Dict[str, Any]:
    """Move an own, attendee-less calendar entry to the trash.

    Separate from cancel_event on purpose: cancel_event is a MEETING
    operation that mails a cancellation to every attendee, and Exchange
    rejects it outright when there is nobody to notify ("keine Empfänger").
    A personal blocker has no attendees, so the only correct disposal is a
    plain move to Deleted Items — no message leaves the mailbox, and the
    item stays recoverable.
    """
    def work(account: Any) -> Dict[str, Any]:
        item = _fetch_one(account, event_id)
        if not isinstance(item, CalendarItem):
            raise ToolError(
                "validation",
                f"{event_id} is a {type(item).__name__}, not a calendar event.",
                hint="Use delete_messages for mail, delete_draft for drafts.",
            )
        attendees = [*(getattr(item, "required_attendees", None) or []),
                     *(getattr(item, "optional_attendees", None) or []),
                     *(getattr(item, "resources", None) or [])]
        if attendees:
            raise ToolError(
                "validation",
                f"'{getattr(item, 'subject', '') or event_id}' has "
                f"{len(attendees)} attendee(s) — deleting it silently would "
                "leave them with a meeting that no longer exists.",
                hint="Use cancel_event, which notifies them (needs "
                     "SEND_ENABLED=true).",
            )
        own = _own_address(account)
        organizer = getattr(item, "organizer", None)
        organizer_address = str(getattr(organizer, "email_address", "") or "").lower()
        # No organizer at all = a plain appointment, which is ours by
        # definition. A FOREIGN organizer means we only hold a copy.
        if organizer_address and own and organizer_address != own:
            raise ToolError(
                "validation",
                f"'{getattr(item, 'subject', '') or event_id}' is organized by "
                f"{getattr(organizer, 'email_address', 'someone else')}, not by you.",
                hint="Decline it with respond_to_event instead of deleting it.",
            )
        item_type = str(getattr(item, "type", "") or "")
        if item_type == "RecurringMaster":
            raise ToolError(
                "validation",
                f"'{getattr(item, 'subject', '') or event_id}' is the master of "
                "a recurring series — deleting it would remove every occurrence.",
                hint="Delete a single occurrence by its own id, or remove the "
                     "series in Outlook.",
            )
        subject = getattr(item, "subject", "") or ""
        start = fmt_dt(getattr(item, "start", None), ctx.settings.ews_tz)
        item.move_to_trash()  # reversible — never HardDelete
        return {"deleted": True, "disposition": "trash",
                "subject": subject, "start": start,
                "occurrence": item_type in ("Occurrence", "Exception")}

    return await _gateway_write(ctx, work)


async def _cancel_event(ctx: Context, *, event_id: str,
                        message: Optional[str] = None) -> Dict[str, Any]:
    def work(account: Any) -> Dict[str, Any]:
        item = _fetch_one(account, event_id)
        kwargs = {"new_body": message} if message else {}
        # 5.0.3 CalendarItem.cancel() exists (CancelCalendarItem service);
        # Exchange itself rejects the call when we are not the organizer.
        item.cancel(**kwargs)
        return {"cancelled": True}

    return await _gateway_write(ctx, work)


async def _delete_messages(ctx: Context, *, ids: List[str],
                           disposition: str = "trash") -> Dict[str, Any]:
    if disposition not in ("trash", "soft", "permanent"):
        raise ToolError("validation",
                        "disposition must be 'trash', 'soft' or 'permanent'")
    _require_bulk(ids)

    deleted_ids: List[str] = []

    def work(account: Any) -> Dict[str, Any]:
        deleted, failed = 0, []
        fetched = _fetch_many(account, ids, only=["id", "changekey"])
        for raw_id, item in zip(ids, fetched):
            try:
                if isinstance(item, Exception):
                    raise item
                if disposition == "trash":
                    item.move_to_trash()
                elif disposition == "soft":
                    item.soft_delete()
                else:
                    # 5.0.3: Item.delete() IS HardDelete — no disposal kwarg.
                    item.delete()
                deleted += 1
                deleted_ids.append(raw_id)
            except Exception as exc:
                failed.append(_failed_entry(ctx, raw_id, exc))
        return {"deleted": deleted, "disposition": disposition, "failed": failed}

    result = await _gateway_write(ctx, work)
    _write_through(ctx, "tombstone_messages", deleted_ids)
    return result


# --- schemas ------------------------------------------------------------------


def _obj(props: Dict[str, Any], required: Optional[List[str]] = None) -> Dict[str, Any]:
    return {"type": "object", "properties": props,
            "required": required or [], "additionalProperties": False}


_STR = {"type": "string"}
_BOOL = {"type": "boolean"}
_EMAILS = {"type": "array", "items": {"type": "string"}}
_IDS = {"type": "array", "items": {"type": "string"},
        "minItems": 1, "maxItems": MAX_BULK_IDS}
def _paths(description: str) -> Dict[str, Any]:
    return {"type": "array", "items": {"type": "string"},
            "description": description}


_ATTACH_NOTE = ("Absolute paths to files on the SERVER host. Confined to "
                "DATA_DIR plus ATTACH_ALLOWED_DIRS; anything else is refused. "
                "To attach a file you first read from another message, call "
                "get_attachment(mode='save') and pass its saved_path back.")


# --- the pack -----------------------------------------------------------------

TOOLS: List[ToolSpec] = [
    ToolSpec(
        name="create_draft",
        description=(
            "Author mail as a DRAFT (never sends). mode=new needs to+body; "
            "reply/reply_all/forward need reply_to (the original message id) "
            "and quote the original server-side; forward also needs to. "
            "body is plain text (the server wraps it in minimal HTML). "
            "importance applies to mode=new only. A forward/reply draft "
            "ALREADY carries the original message's attachments — do not "
            "re-attach them or they go out twice. attachments takes local "
            "file paths on the SERVER host and attaches them to the draft "
            "(confined to DATA_DIR plus ATTACH_ALLOWED_DIRS). "
            "Use send_draft to send."
            " MeetingResponse items without a native reply API use a transparent "
            "recipient-reconstruction fallback. Optional idempotency_key makes "
            "create retries safe across processes."
        ),
        side_effect_class="write",
        input_schema=_obj({
            "mode": {"type": "string", "enum": list(_DRAFT_MODES), "default": "new"},
            "reply_to": _STR, "to": _EMAILS, "cc": _EMAILS, "bcc": _EMAILS,
            "subject": _STR, "body": _STR,
            "importance": {"type": "string", "enum": ["normal", "high"]},
            "attachments": _paths("Files to attach to the new draft. " + _ATTACH_NOTE),
            "idempotency_key": {"type": "string", "minLength": 1,
                                "description": "Deduplicate safe retries across processes."},
        }, required=["body"]),
        handler=_create_draft,
        confirm=False,
    ),
    ToolSpec(
        name="update_draft",
        description=("Update fields of an existing draft (to/cc/bcc/subject/body). "
                     "Only the supplied fields change. attachments takes local "
                     "file paths on the SERVER host and ADDS them to the draft "
                     "(existing attachments are kept)."),
        side_effect_class="write",
        input_schema=_obj({
            "draft_id": _STR, "to": _EMAILS, "cc": _EMAILS, "bcc": _EMAILS,
            "subject": _STR, "body": _STR,
            "attachments": _paths("Files to ADD to the draft (existing "
                                  "attachments are kept). " + _ATTACH_NOTE),
        }, required=["draft_id"]),
        handler=_update_draft,
        confirm=False,
    ),
    ToolSpec(
        name="delete_draft",
        description=("Move a draft to trash (recoverable). Refuses items that "
                     "are not in f:drafts."),
        side_effect_class="write",
        input_schema=_obj({"draft_id": _STR}, required=["draft_id"]),
        handler=_delete_draft,
        confirm=False,
    ),
    ToolSpec(
        name="update_messages",
        description=(
            "Bulk-update up to 50 messages: set_read and/or categories_add/"
            "categories_remove. Per-item failures are isolated and reported. "
            "There is no follow-up-flag support in this backend — use "
            "categories_add (e.g. ['Follow up']) as the visible marker."
        ),
        side_effect_class="write",
        input_schema=_obj({
            "ids": _IDS, "set_read": _BOOL,
            "categories_add": _EMAILS, "categories_remove": _EMAILS,
        }, required=["ids"]),
        handler=_update_messages,
        confirm=False,
    ),
    ToolSpec(
        name="move_messages",
        description=("Move up to 50 messages to a folder (id, f:alias or path). "
                     "Message ids are re-bound automatically: the SAME id keeps "
                     "working after the move."),
        side_effect_class="write",
        input_schema=_obj({"ids": _IDS, "to_folder": _STR},
                          required=["ids", "to_folder"]),
        handler=_move_messages,
        confirm=False,
    ),
    ToolSpec(
        name="create_event",
        description=(
            "Create a calendar event. Default send_invitations=false saves a "
            "placeholder WITHOUT emailing attendees (no confirm needed). With "
            "send_invitations=true, invitations leave the org: two-phase "
            "confirm + send kill-switch apply."
            " Optional show_as includes working_elsewhere (Outlook: An einem "
            "anderen Ort); idempotency_key makes create retries safe."
        ),
        side_effect_class="write",  # tier: draft; invite path confirm-gated below
        input_schema=_obj({
            "subject": _STR, "start": _STR, "end": _STR,
            "attendees": _EMAILS, "location": _STR, "body": _STR,
            "show_as": {"type": "string", "enum": list(SHOW_AS_VALUES),
                        "description": "Outlook free/busy display status."},
            "idempotency_key": {"type": "string", "minLength": 1,
                                "description": "Deduplicate safe retries across processes."},
            "send_invitations": {"type": "boolean", "default": False},
        }, required=["subject", "start", "end"]),
        handler=_create_event,
        confirm=lambda kw: bool(kw.get("send_invitations")),
    ),
    ToolSpec(
        name="update_event",
        description=(
            "Update event fields (subject/start/end/location/body/show_as). Default "
            "notify_attendees=false is a silent metadata edit; true notifies "
            "changed attendees → two-phase confirm + send kill-switch apply."
        ),
        side_effect_class="write",
        input_schema=_obj({
            "event_id": _STR, "subject": _STR, "start": _STR, "end": _STR,
            "location": _STR, "body": _STR,
            "show_as": {"type": "string", "enum": list(SHOW_AS_VALUES),
                        "description": "Outlook free/busy display status."},
            "notify_attendees": {"type": "boolean", "default": False},
        }, required=["event_id"]),
        handler=_update_event,
        confirm=lambda kw: bool(kw.get("notify_attendees")),
    ),
    ToolSpec(
        name="send_draft",
        description=(
            "Send an existing draft (the ONLY way mail leaves this mailbox). "
            "Two-phase: the first call fetches the draft and returns its REAL "
            "recipients/subject/body as a preview + confirm_token bound to "
            "that content; the second call with the token re-verifies the "
            "content and sends (editing the draft in between invalidates the "
            "token). Optional idempotency_key dedupes retries (replays return "
            "the stored result)."
        ),
        side_effect_class="send",
        input_schema=_obj({"draft_id": _STR, "idempotency_key": _STR},
                          required=["draft_id"]),
        handler=_send_draft,
        confirm=_send_confirm_needed,
        preview=_send_draft_preview,
    ),
    ToolSpec(
        name="respond_to_event",
        description=("Accept/tentative/decline a meeting — this SENDS a response "
                     "to the organizer, so it is two-phase confirmed."),
        side_effect_class="send",
        input_schema=_obj({
            "event_id": _STR,
            "response": {"type": "string", "enum": sorted(_RESPONSE_METHOD)},
            "message": _STR,
        }, required=["event_id", "response"]),
        handler=_respond_to_event,
        confirm=True,
    ),
    ToolSpec(
        name="delete_event",
        description=(
            "Move your own calendar entry to the trash WITHOUT notifying "
            "anyone — the disposal path for personal blockers and holds. "
            "Refuses an event that has attendees (use cancel_event), one "
            "organized by somebody else (decline it instead), and the master "
            "of a recurring series (delete a single occurrence by its own id). "
            "Recoverable from Deleted Items; no message leaves the mailbox, so "
            "it works with SEND_ENABLED=false."
        ),
        side_effect_class="destructive",
        input_schema=_obj({"event_id": _STR}, required=["event_id"]),
        handler=_delete_event,
    ),
    ToolSpec(
        name="cancel_event",
        description=("Cancel a meeting you organize — sends cancellations to all "
                     "attendees (destructive, two-phase confirmed). For an "
                     "attendee-less entry of your own use delete_event: Exchange "
                     "rejects a cancellation that has no recipients."),
        side_effect_class="destructive",
        input_schema=_obj({"event_id": _STR, "message": _STR},
                          required=["event_id"]),
        handler=_cancel_event,
        confirm=True,
    ),
    ToolSpec(
        name="delete_messages",
        description=(
            "Delete up to 50 messages. disposition: 'trash' (default, "
            "recoverable), 'soft' (dumpster) or 'permanent' (UNRECOVERABLE — "
            "requires two-phase confirm). Class is destructive, so ALL "
            "dispositions need the full tier in v5.0 (conservative)."
        ),
        side_effect_class="destructive",
        input_schema=_obj({
            "ids": _IDS,
            "disposition": {"type": "string",
                            "enum": ["trash", "soft", "permanent"],
                            "default": "trash"},
        }, required=["ids"]),
        handler=_delete_messages,
        confirm=lambda kw: kw.get("disposition") == "permanent",
    ),
    ToolSpec(
        name="set_oof",
        description=(
            "Set out-of-office auto-replies (externally visible → send class, "
            "two-phase confirmed). state: disabled|enabled|scheduled; "
            "scheduled needs start+end. external_reply defaults to "
            "internal_reply when omitted."
        ),
        side_effect_class="send",
        input_schema=_obj({
            "state": {"type": "string", "enum": sorted(_OOF_STATE)},
            "internal_reply": _STR, "external_reply": _STR,
            "start": _STR, "end": _STR,
        }, required=["state"]),
        handler=_set_oof,
        confirm=True,
    ),
]
