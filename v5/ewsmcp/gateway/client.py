"""EWS gateway — one Account per process, EWS off the event loop.

Auth rule (verified live 2026-06-12): NEVER pin auth_type against this
Exchange; only exchangelib auto-negotiation works, and during front-door
lockdown windows nothing fresh authenticates at all — which is why the
ConnectionManager treats connecting as a state, not a failure.
"""

import asyncio
import logging
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, Dict, List, Optional, Tuple

from exchangelib import DELEGATE, Account, Configuration, Credentials, EWSTimeZone
from exchangelib.protocol import (
    BaseProtocol,
    CachingProtocol,
    FaultTolerance,
    NoVerifyHTTPAdapter,
)

from ..config import Settings
from ..errors import ToolError

logger = logging.getLogger(__name__)

WELL_KNOWN = {
    "f:root": "msg_folder_root",
    "f:inbox": "inbox", "f:sent": "sent", "f:drafts": "drafts",
    "f:trash": "trash", "f:junk": "junk", "f:outbox": "outbox",
    "f:calendar": "calendar", "f:contacts": "contacts", "f:tasks": "tasks",
}


class EWSGateway:
    def __init__(self, settings: Settings, coordinator: Any = None):
        self.settings = settings
        self.coordinator = coordinator
        self._account: Optional[Account] = None
        self._account_lock = threading.Lock()
        self._pool = ThreadPoolExecutor(
            max_workers=max(1, settings.ews_max_concurrency),
            thread_name_prefix="ews",
        )
        # Calendar reads, writes and background sync must not starve one
        # another when exchangelib keeps a cancelled request alive in a
        # worker.  Each lane has a bounded executor; the coordinator also
        # enforces the same lane across local processes.
        self._lane_pools = {
            "calendar": ThreadPoolExecutor(max_workers=1, thread_name_prefix="ews-calendar"),
            "write": ThreadPoolExecutor(max_workers=1, thread_name_prefix="ews-write"),
            "sync": ThreadPoolExecutor(max_workers=1, thread_name_prefix="ews-sync"),
        }
        self.last_connection_error: Optional[str] = None
        self._folder_cache: Dict[str, Any] = {}
        self._folder_cache_ts = 0.0
        self._folder_revision_seen: Optional[int] = None
        if settings.ews_insecure_skip_verify:
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
            logger.warning("TLS verification DISABLED for Exchange traffic")

    # ------------------------------------------------------------- account

    @property
    def account(self) -> Account:
        with self._account_lock:
            if self._account is None:
                self._account = self._build_account()
            return self._account

    def _build_account(self) -> Account:
        s = self.settings
        BaseProtocol.TIMEOUT = s.request_timeout
        kwargs: Dict[str, Any] = dict(
            service_endpoint=s.ews_server_url,
            credentials=Credentials(s.ews_username or s.ews_email, s.ews_password or ""),
            retry_policy=FaultTolerance(max_wait=s.ews_retry_max_wait_seconds),
        )
        if s.ews_auth_type_force:  # escape hatch for a DIFFERENT Exchange only
            logger.warning(
                "auth_type FORCED to %s — the primary Exchange requires auto-negotiation",
                s.ews_auth_type_force,
            )
            kwargs["auth_type"] = s.ews_auth_type_force
        config = Configuration(**kwargs)
        return Account(
            primary_smtp_address=s.ews_email,
            config=config,
            autodiscover=False,
            access_type=DELEGATE,
            default_timezone=EWSTimeZone(s.ews_tz),
        )

    def reset(self) -> None:
        """Drop the cached account AND exchangelib's protocol-cache entry.

        Dropping only our Account is not enough: ``CachingProtocol`` hands
        the same wedged Protocol (with its already-negotiated auth type)
        right back on the next build, so a session that died mid-outage
        would never renegotiate. Clearing the cache forces a genuinely
        fresh session + auth negotiation on the next access.
        """
        with self._account_lock:
            if self._account is not None:
                try:
                    self._account.protocol.close()
                except Exception:
                    pass
                self._account = None
        try:
            CachingProtocol.clear_cache()
        except Exception as e:
            logger.debug("protocol cache clear failed: %s", e)
        self._folder_cache.clear()
        self._folder_cache_ts = 0.0
        self._folder_revision_seen = None

    def test_connection(self) -> bool:
        """Real network probe — must round-trip on EVERY call.

        ``inbox.total_count`` is a cached property after its first read, so
        probing it reported warm forever once it had succeeded once (the
        false-warm bug: /readyz lied through outages and the reset ladder
        never ran). ``root.refresh()`` issues a GetFolder request each time.
        """
        try:
            self.account.root.refresh()
            self.last_connection_error = None
            return True
        except Exception as e:
            self.last_connection_error = f"{type(e).__name__}: {e}"
            logger.error("connection test failed: %s", self.last_connection_error)
            return False

    # ------------------------------------------------------------- calling

    async def call(self, fn: Callable[[Account], Any], *, lane: str = "read",
                   timeout_s: Optional[float] = None) -> Any:
        """Run blocking EWS work on the bounded pool; the pool size IS the
        EWS concurrency cap (polite guest on the per-user throttle budget).

        A timed-out asyncio caller does not stop the underlying HTTP request.
        Separate pools contain that unavoidable Python limitation, while the
        interprocess lane lock prevents retries from multiplying the load.
        """
        loop = asyncio.get_running_loop()
        timeout = float(timeout_s if timeout_s is not None else self.settings.request_timeout)
        # A non-positive explicit timeout disables only the *outer* asyncio
        # deadline.  This is reserved for isolated background work such as
        # cache sync, which consists of several bounded EWS requests and can
        # legitimately take longer than one request timeout in total.  The
        # exchangelib HTTP timeout and retry policy still bound each upstream
        # request, and the dedicated sync pool prevents user-facing calls from
        # being starved.
        no_outer_deadline = timeout_s is not None and timeout <= 0
        pool = self._lane_pools.get(lane, self._pool)

        queued_at = time.time()

        def invoke() -> Any:
            # Three phases, timed separately: how long the job sat in this
            # process's lane pool, how long it waited for the SHARED lane
            # lock, and how long Exchange actually took. A slow call looks
            # identical from outside no matter which of the three it was,
            # which is exactly what made 2026-10-09/12 hard to pin down.
            pool_wait = time.time() - queued_at
            lock_wait = 0.0
            started = time.time()
            try:
                if self.coordinator is None:
                    return fn(self.account)
                lock_timeout = (float(self.settings.request_timeout)
                                if no_outer_deadline else timeout)
                before_lock = time.time()
                with self.coordinator.live_call(lane, lock_timeout):
                    lock_wait = time.time() - before_lock
                    started = time.time()
                    return fn(self.account)
            except Exception as exc:
                if self.coordinator is not None:
                    backoff = self.coordinator.backoff_seconds(exc)
                    if backoff is not None:
                        self.coordinator.set_cooldown(
                            backoff, f"{type(exc).__name__}: {exc}")
                raise
            finally:
                total = time.time() - queued_at
                if total >= 2.0:
                    logger.warning(
                        "slow %s lane call: %.1fs total "
                        "(pool wait %.1fs, lock wait %.1fs, exchange %.1fs)",
                        lane, total, pool_wait, lock_wait,
                        max(0.0, time.time() - started))

        future = loop.run_in_executor(pool, invoke)
        if no_outer_deadline:
            return await asyncio.shield(future)
        try:
            # Shield keeps the executor future alive after caller cancellation;
            # it will release the interprocess lock when exchangelib returns.
            return await asyncio.wait_for(asyncio.shield(future), timeout=timeout)
        except asyncio.TimeoutError:
            raise ToolError(
                "upstream_timeout",
                f"Exchange {lane} call exceeded {int(timeout)} seconds; "
                "the isolated worker may still finish in the background",
                retry_after_s=1,
            ) from None

    # ------------------------------------------------------------- folders

    def prepare_folder_tree(self, account: Account) -> None:
        """Invalidate this process's tree after a mutation in another one."""
        coordinator = self.coordinator
        if coordinator is None or not hasattr(coordinator, "folder_revision"):
            return
        revision = int(coordinator.folder_revision())
        if self._folder_revision_seen is None:
            self._folder_revision_seen = revision
            return
        if revision == self._folder_revision_seen:
            return
        try:
            account.msg_folder_root.clear_cache()
        except Exception as exc:
            logger.debug("exchangelib folder cache clear failed: %s", exc)
        self._folder_cache.clear()
        self._folder_cache_ts = 0.0
        self._folder_revision_seen = revision

    def mark_folder_tree_changed(self, account: Optional[Account] = None) -> int:
        """Publish a local folder mutation to every process sharing DATA_DIR."""
        self._folder_cache.clear()
        self._folder_cache_ts = 0.0
        coordinator = self.coordinator
        if coordinator is not None and hasattr(coordinator, "bump_folder_revision"):
            revision = int(coordinator.bump_folder_revision())
            self._folder_revision_seen = revision
            return revision
        return 0

    def _folder_map(self, account: Account) -> Dict[str, Any]:
        """{raw_id|lower_path: Folder} cache, rebuilt every 300s (sync)."""
        self.prepare_folder_tree(account)
        now = time.time()
        if self._folder_cache and now - self._folder_cache_ts < 300:
            return self._folder_cache
        cache: Dict[str, Any] = {}
        try:
            for folder in account.msg_folder_root.walk():
                if getattr(folder, "id", None):
                    cache[folder.id] = folder
                try:
                    path = "/".join(
                        p.name for p in folder.parts[2:]  # drop root/Top of Info Store
                    ) or folder.name
                except Exception:
                    path = folder.name
                cache[path.lower()] = folder
        except Exception as e:
            logger.warning("folder walk failed: %s", e)
        if cache:
            self._folder_cache = cache
            self._folder_cache_ts = now
        return cache

    def resolve_folder(self, account: Account, ref: Optional[str], aliaser) -> Any:
        """well-known alias | folder alias (f12) | path | raw id → Folder (sync)."""
        self.prepare_folder_tree(account)
        if not ref:
            return account.inbox
        key = ref.strip()
        attr = WELL_KNOWN.get(key.lower()) or WELL_KNOWN.get(f"f:{key.lower()}")
        if attr:
            return getattr(account, attr)
        try:
            key = aliaser.resolve(key)  # f12 → raw id; raw/path pass through
        except KeyError as e:
            raise ToolError("validation", str(e.args[0] if e.args else e))
        cache = self._folder_map(account)
        folder = cache.get(key) or cache.get(key.lower())
        if folder is None:
            raise ToolError(
                "not_found", f"No folder matches {ref!r}.",
                hint="Use list_folders and pass one of its ids or paths.",
            )
        return folder


def paginate(query: Any, *, offset: int, limit: int,
             chunk: int = 50) -> Tuple[List[Any], Optional[int]]:
    """Materialize query[offset:offset+limit] in chunks (sync, raises on
    mid-iteration failure — the caller's error mapper classifies it).

    Returns ``(items, next_offset)``. NEVER calls ``QuerySet.count()`` —
    in exchangelib that iterates every matching id server-side, so a 20k
    inbox paid ~20k ids of round trips on every "read 10 emails". Whether
    another page exists comes from a one-item lookahead instead; callers
    that want an exact total use a refreshed ``folder.total_count`` (only
    valid for unfiltered listings) or a local mirror count.
    """
    offset = max(0, offset)
    limit = max(0, limit)
    lookahead = limit + 1
    items: List[Any] = []
    cursor = offset
    chunk = max(1, min(chunk, 250))
    while len(items) < lookahead:
        want = min(chunk, lookahead - len(items))
        batch = list(query[cursor:cursor + want])
        if not batch:
            break
        items.extend(batch)
        cursor += len(batch)
        if len(batch) < want:
            break
    next_offset = offset + limit if len(items) > limit else None
    return items[:limit], next_offset
