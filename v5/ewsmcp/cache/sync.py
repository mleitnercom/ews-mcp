"""Background delta-sync: exchangelib's native SyncFolderItems primitives.

Owned by the ConnectionManager's warm state: the engine starts when the
connection first warms up and loops forever. Every cycle applies item
deltas for the hot folders (resumable ``sync_state`` tokens persisted in
the store — a container restart resumes instead of re-downloading), and a
slower cadence refreshes the folder tree (which structurally fixes the
frozen unread/total-counts finding), the expanded calendar window, and
the tasks folder.

Failure posture: ANY exception marks the engine degraded and is retried
next cycle; tools always fall back to live EWS — the cache accelerates,
never gates. All EWS work runs on the gateway's bounded pool.
"""

from __future__ import annotations

import asyncio
import json
import logging
import time
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional
from zoneinfo import ZoneInfo

from ..bodyclean import clean_body
from ..calendar_status import from_ews_show_as
from ..dto import fmt_dt
from ..folders import collect_folder_rows
from .store import CacheStore

logger = logging.getLogger(__name__)

# Fields pulled per synced message — everything a card/full DTO needs,
# never the MIME. Cleaning happens here, once, at sync time.
ITEM_FIELDS = [
    "id", "changekey", "subject", "sender", "datetime_received", "is_read",
    "has_attachments", "importance", "categories", "conversation_id",
    "message_id", "to_recipients", "text_body",
]
TASK_FIELDS = ["id", "changekey", "subject", "due_date", "is_complete", "status"]
BODY_CLEAN_MAX = 20_000
# Superseded by EWS_CALENDAR_WINDOW_DAYS / _MAX_ITEMS; kept so an
# out-of-tree importer does not break.
CALENDAR_WINDOW_DAYS = 14
CALENDAR_MAX_ITEMS = 200


def _ts(dt: Any) -> Optional[int]:
    try:
        return int(dt.timestamp())
    except (AttributeError, OSError, OverflowError, ValueError, TypeError):
        return None


def row_from_message(item: Any, folder_key: str, tz: str) -> Dict[str, Any]:
    """Message item → mirror row. Body cleaned HERE (the 700× paid once)."""
    sender = getattr(item, "sender", None)
    sender_email = getattr(sender, "email_address", None) or ""
    sender_name = getattr(sender, "name", None) or ""
    subject = getattr(item, "subject", None) or ""
    text = getattr(item, "text_body", None) or ""
    body_clean = ""
    if text:
        try:
            body_clean = clean_body(text, max_chars=BODY_CLEAN_MAX)["text"]
        except Exception:
            body_clean = text[:BODY_CLEAN_MAX]
    to = [
        r.email_address
        for r in (getattr(item, "to_recipients", None) or [])
        if getattr(r, "email_address", None)
    ]
    received = getattr(item, "datetime_received", None)
    conv = getattr(item, "conversation_id", None)
    imid = getattr(item, "message_id", None)
    return {
        "ews_id": str(item.id),
        "changekey": getattr(item, "changekey", None),
        "folder": folder_key,
        "conversation_id": getattr(conv, "id", None),
        "sender_name": sender_name,
        "sender_email": sender_email,
        "to_json": json.dumps(to, ensure_ascii=False),
        "subject": subject,
        "date_ts": _ts(received),
        "date_iso": fmt_dt(received, tz),
        "is_read": 1 if getattr(item, "is_read", True) else 0,
        "has_attachments": 1 if getattr(item, "has_attachments", False) else 0,
        "importance": str(getattr(item, "importance", "") or "") or None,
        "categories_json": json.dumps(list(getattr(item, "categories", None) or []),
                                      ensure_ascii=False),
        "body_clean": body_clean,
        "internet_message_id": imid if isinstance(imid, str) else None,
        "norm_text": CacheStore.norm_for_row(subject, sender_name,
                                             sender_email, body_clean),
    }


def row_from_event(item: Any, tz: str) -> Dict[str, Any]:
    organizer = getattr(item, "organizer", None)
    return {
        "ews_id": str(getattr(item, "id", "") or ""),
        "changekey": getattr(item, "changekey", None),
        "subject": getattr(item, "subject", "") or "",
        "start_ts": _ts(getattr(item, "start", None)),
        "start_iso": fmt_dt(getattr(item, "start", None), tz),
        "end_ts": _ts(getattr(item, "end", None)),
        "end_iso": fmt_dt(getattr(item, "end", None), tz),
        "location": str(getattr(item, "location", None) or "") or None,
        "organizer": getattr(organizer, "email_address", None),
        "is_recurring": 1 if (getattr(item, "is_recurring", False)
                              or getattr(item, "recurrence", None)) else 0,
        "my_response": str(getattr(item, "my_response_type", None) or "") or None,
        "show_as": from_ews_show_as(
            getattr(item, "legacy_free_busy_status", None)),
    }


def row_from_task(item: Any, tz: str) -> Dict[str, Any]:
    due = getattr(item, "due_date", None)
    return {
        "ews_id": str(getattr(item, "id", "") or ""),
        "changekey": getattr(item, "changekey", None),
        "subject": getattr(item, "subject", "") or "",
        "due_ts": _ts(due),
        "due_iso": fmt_dt(due, tz) if hasattr(due, "astimezone") else (
            due.isoformat() if due is not None else None),
        "is_complete": 1 if getattr(item, "is_complete", False) else 0,
        "status": str(getattr(item, "status", None) or "") or None,
    }


class SyncEngine:
    def __init__(self, settings: Any, gateway: Any, store: CacheStore,
                 semantic: Any = None, coordinator: Any = None):
        self.settings = settings
        self.gateway = gateway
        self.store = store
        self.semantic = semantic  # optional vector tier; failures degrade
        self.coordinator = coordinator
        self.folder_keys = [
            k.strip().lower()
            for k in (settings.ews_cache_folders or "").split(",") if k.strip()
        ]
        self.last_error: Optional[str] = None
        self.last_cycle_ts: Optional[float] = None
        self.cycles = 0
        self._task: Optional[asyncio.Task] = None
        self._stopped = False
        self._last_slow_ts = 0.0

    # ------------------------------------------------------------ lifecycle

    async def start(self) -> None:
        if self._task is None:
            self._task = asyncio.create_task(self._loop(), name="cache-sync")
            logger.info("cache sync engine started (folders=%s, every %ss)",
                        self.folder_keys, self.settings.ews_cache_sync_seconds)

    async def stop(self) -> None:
        self._stopped = True
        if self._task is not None and not self._task.done():
            self._task.cancel()
            try:
                await self._task
            except (asyncio.CancelledError, Exception):
                pass

    def status(self) -> Dict[str, Any]:
        return {
            "cycles": self.cycles,
            "last_cycle_age_s": (int(time.time() - self.last_cycle_ts)
                                 if self.last_cycle_ts else None),
            "last_error": self.last_error,
            "leader": (self.coordinator.is_sync_leader
                       if self.coordinator is not None else True),
        }

    async def _loop(self) -> None:
        while not self._stopped:
            try:
                if self.coordinator is not None:
                    # EVERY instance ticks here, leader or not. Without this
                    # the heartbeat was only written when someone happened to
                    # call get_server_status, so instance counts were noise:
                    # 11 stale rows one day, 0 live ones the next.
                    self.coordinator.heartbeat()
                if self.coordinator is None or self.coordinator.try_become_sync_leader():
                    await self._cycle()
                self.last_error = None
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                # Degrade, never die: tools keep answering from live EWS
                # (or from the last good mirror state) while we retry.
                self.last_error = f"{type(exc).__name__}: {exc}"[:500]
                logger.warning("cache sync cycle failed: %s", self.last_error)
            self.last_cycle_ts = time.time()
            self.cycles += 1
            await asyncio.sleep(max(5, int(self.settings.ews_cache_sync_seconds)))

    # --------------------------------------------------------------- cycle

    async def _cycle(self) -> None:
        await self._gateway_call(self._sync_mail_folders)
        slow_every = max(60, int(self.settings.ews_cache_hierarchy_seconds))
        if time.time() - self._last_slow_ts >= slow_every:
            await self._gateway_call(self._sync_slow_lane)
            self._last_slow_ts = time.time()

    async def _gateway_call(self, fn: Any) -> Any:
        # Test gateways and older adapters only accept ``call(fn)``.
        import inspect
        if "lane" in inspect.signature(self.gateway.call).parameters:
            # A sync job performs multiple individually bounded EWS requests.
            # Applying the single-request timeout to the entire job reports a
            # false failure after 30s while its isolated worker keeps updating
            # the mirror in the background.  The sync lane has its own worker
            # and interprocess lock, so let the complete job drain.
            return await self.gateway.call(fn, lane="sync", timeout_s=0)
        return await self.gateway.call(fn)

    def _sync_mail_folders(self, account: Any) -> None:
        """Apply item deltas for each hot folder (runs on the EWS pool)."""
        tz = self.settings.ews_tz
        window_floor = (
            datetime.now(ZoneInfo(tz))
            - timedelta(days=int(self.settings.ews_cache_window_days))
        )
        for key in self.folder_keys:
            folder = getattr(account, key, None)
            if folder is None:
                continue
            token = self.store.get_sync_state(f"item:{key}")
            upserts: List[Dict[str, Any]] = []
            deletes: List[str] = []
            read_flags: List[tuple] = []
            for change_type, payload in folder.sync_items(
                sync_state=token, only_fields=ITEM_FIELDS,
            ):
                if change_type in ("create", "update"):
                    received = getattr(payload, "datetime_received", None)
                    if received is not None and received < window_floor:
                        continue  # outside the mirror window — skip storing
                    if getattr(payload, "id", None):
                        upserts.append(row_from_message(payload, key, tz))
                elif change_type == "delete":
                    deletes.append(str(payload.id))
                elif change_type == "read_flag_change":
                    item_id, is_read = payload
                    read_flags.append((str(item_id.id), bool(is_read)))
            self.store.upsert_messages(upserts)
            self.store.delete_messages_by_id(deletes)
            for ews_id, is_read in read_flags:
                self.store.set_read_flag([ews_id], is_read)
            self.store.set_sync_state(f"item:{key}", folder.item_sync_state,
                                      time.time())
            if self.semantic is not None and (upserts or deletes):
                try:  # embeddings ride the sync, never gate it
                    self.semantic.add([
                        {"ews_id": r["ews_id"],
                         "text": f"{r['subject']}\n{r['body_clean']}"}
                        for r in upserts
                    ])
                    self.semantic.delete(deletes)
                except Exception as exc:
                    logger.warning("semantic indexing skipped this cycle: %s", exc)

    def _sync_slow_lane(self, account: Any) -> None:
        """Folder tree + expanded calendar window + tasks (every ~10 min)."""
        tz = self.settings.ews_tz
        # Folder tree with fresh counts — the frozen-counts fix.
        try:
            if hasattr(self.gateway, "prepare_folder_tree"):
                self.gateway.prepare_folder_tree(account)
            rows = collect_folder_rows(account)
            self.store.replace_folders(rows)
            self.store.set_sync_state("folders", None, time.time())
        except Exception as exc:
            logger.debug("folder tree sync failed: %s", exc)

        # Expanded calendar occurrences for the overview window. The window
        # reaches BACKWARDS as well: a briefing that includes yesterday, or a
        # weekly review, would otherwise miss the mirror by definition and
        # fall through to a live CalendarView.
        try:
            now = datetime.now(ZoneInfo(tz))
            day_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
            window_start = day_start - timedelta(
                days=int(self.settings.ews_calendar_backfill_days))
            window_end = day_start + timedelta(
                days=int(self.settings.ews_calendar_window_days))
            events = list(account.calendar.view(
                start=window_start,
                end=window_end,
                max_items=int(self.settings.ews_calendar_max_items),
            ))
            self.store.replace_events(
                [row_from_event(ev, tz) for ev in events
                 if getattr(ev, "id", None)])
            self.store.set_sync_state("events", None, time.time())
            self.store.set_sync_state("events:start",
                                      str(int(window_start.timestamp())),
                                      time.time())
            self.store.set_sync_state("events:end",
                                      str(int(window_end.timestamp())),
                                      time.time())
        except Exception as exc:
            logger.debug("calendar window sync failed: %s", exc)

        # Tasks folder (small; delta-synced like mail).
        try:
            tasks_folder = getattr(account, "tasks", None)
            if tasks_folder is not None:
                token = self.store.get_sync_state("item:tasks")
                upserts: List[Dict[str, Any]] = []
                deletes: List[str] = []
                for change_type, payload in tasks_folder.sync_items(
                    sync_state=token, only_fields=TASK_FIELDS,
                ):
                    if change_type in ("create", "update"):
                        if getattr(payload, "id", None):
                            upserts.append(row_from_task(payload, tz))
                    elif change_type == "delete":
                        deletes.append(str(payload.id))
                self.store.upsert_tasks(upserts)
                self.store.delete_tasks_by_id(deletes)
                self.store.set_sync_state("item:tasks",
                                          tasks_folder.item_sync_state,
                                          time.time())
        except Exception as exc:
            logger.debug("tasks sync failed: %s", exc)
