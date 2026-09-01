"""Regression coverage for calendar mirror and multi-process safety."""

from __future__ import annotations

import asyncio
import sqlite3
import threading
import time
from datetime import datetime
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pytest
from conftest import make_settings

from ewsmcp.audit import AuditLog
from ewsmcp.cache.store import CacheStore
from ewsmcp.coordination import ProcessCoordinator
from ewsmcp.dates import parse_when
from ewsmcp.errors import ToolError
from ewsmcp.gateway.client import EWSGateway
from ewsmcp.ids import get_aliaser
from ewsmcp.tools import calendar_people
from ewsmcp.tools.base import Context, dispatch

SPEC = {spec.name: spec for spec in calendar_people.TOOLS}
VIENNA = ZoneInfo("Europe/Vienna")


def _event_row(raw_id: str, start: datetime, end: datetime, **overrides):
    row = {
        "ews_id": raw_id,
        "changekey": "CK",
        "subject": raw_id,
        "start_ts": int(start.timestamp()),
        "start_iso": start.isoformat(timespec="minutes"),
        "end_ts": int(end.timestamp()),
        "end_iso": end.isoformat(timespec="minutes"),
        "location": None,
        "organizer": None,
        "is_recurring": 0,
        "my_response": None,
        "show_as": "busy",
    }
    row.update(overrides)
    return row


class _Gateway:
    def __init__(self, account=None, error=None):
        self.account = account
        self.error = error
        self.calls = 0

    async def call(self, fn):
        self.calls += 1
        if self.error:
            raise self.error
        return fn(self.account)


def _context(tmp_path, gateway, store):
    settings = make_settings(ews_tz="Europe/Vienna")
    return Context(
        settings=settings,
        gateway=gateway,
        manager=None,
        aliaser=get_aliaser(str(tmp_path / "aliases")),
        audit=AuditLog(str(tmp_path / "data")),
        cache=store,
    )


def _run(ctx, args):
    return asyncio.run(dispatch(ctx, SPEC["list_events"], dict(args)))


def test_cache_first_calendar_paging_overlap_and_recurrence(tmp_path):
    store = CacheStore(tmp_path / "mirror.db")
    day = datetime(2026, 9, 4, tzinfo=VIENNA)
    store.replace_events(
        [
            _event_row("E0", day.replace(hour=7), day.replace(hour=8)),
            _event_row(
                "E1",
                day.replace(hour=8),
                day.replace(hour=9),
                is_recurring=1,
                show_as="working_elsewhere",
            ),
            _event_row("E2", day.replace(hour=9), day.replace(hour=10)),
            _event_row("E3", day.replace(hour=10), day.replace(hour=11)),
        ]
    )
    store.set_sync_state("events", None, time.time())
    gateway = _Gateway(error=AssertionError("live Exchange must not run"))
    ctx = _context(tmp_path, gateway, store)
    result = _run(
        ctx,
        {
            "start": "2026-09-04T08:30+02:00",
            "end": "2026-09-04T11:00+02:00",
            "offset": 1,
            "limit": 1,
        },
    )
    assert result["source"] == "cache" and gateway.calls == 0
    assert result["total_available"] == 3 and result["next_offset"] == 2
    assert result["items"][0]["subject"] == "E2"

    first = _run(
        ctx, {"start": "2026-09-04T08:30+02:00", "end": "2026-09-04T11:00+02:00", "limit": 1}
    )
    assert first["items"][0]["recurring"] is True
    assert first["items"][0]["show_as"] == "working_elsewhere"


def test_fresh_calendar_falls_back_to_warm_mirror(tmp_path):
    store = CacheStore(tmp_path / "mirror.db")
    day = datetime(2026, 9, 4, tzinfo=VIENNA)
    store.replace_events([_event_row("E1", day.replace(hour=8), day.replace(hour=9))])
    store.set_sync_state("events", None, time.time())
    ctx = _context(tmp_path, _Gateway(error=TimeoutError("slow view")), store)
    result = _run(ctx, {"start": "2026-09-04", "end": "2026-09-05", "fresh": True})
    assert result["source"] == "cache"
    assert result["fallback"] == "live_failed"
    assert "slow view" in result["warning"]


def test_fresh_calendar_uses_live_when_available(tmp_path):
    store = CacheStore(tmp_path / "mirror.db")
    day = datetime(2026, 9, 4, tzinfo=VIENNA)
    store.replace_events([_event_row("CACHE", day.replace(hour=8), day.replace(hour=9))])
    store.set_sync_state("events", None, time.time())
    live = SimpleNamespace(
        id="LIVE",
        subject="Live event",
        start=day.replace(hour=10),
        end=day.replace(hour=11),
        location=None,
        organizer=None,
        my_response_type=None,
        is_recurring=False,
        recurrence=None,
        legacy_free_busy_status="WorkingElsewhere",
    )
    account = SimpleNamespace(calendar=SimpleNamespace(view=lambda **kwargs: [live]))
    ctx = _context(tmp_path, _Gateway(account=account), store)
    result = _run(ctx, {"start": "2026-09-04", "end": "2026-09-05", "fresh": True})
    assert result["source"] == "live"
    assert result["items"][0]["show_as"] == "working_elsewhere"


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("2026-09-04T00:00:00+02:00", "2026-09-04T00:00:00+02:00"),
        ("2026-01-04T00:00:00+01:00", "2026-01-04T00:00:00+01:00"),
        ("2026-09-03T22:00:00Z", "2026-09-04T00:00:00+02:00"),
    ],
)
def test_iso_offsets_normalize_to_named_vienna_zone(raw, expected):
    parsed = parse_when(raw, "start", "Europe/Vienna")
    assert parsed.isoformat() == expected
    assert str(parsed.tzinfo) == "Europe/Vienna"


def test_event_schema_migrates_show_as_column(tmp_path):
    db = tmp_path / "legacy.db"
    with sqlite3.connect(db) as conn:
        conn.execute("""CREATE TABLE events (
            ews_id TEXT PRIMARY KEY, changekey TEXT, subject TEXT,
            start_ts INTEGER, start_iso TEXT, end_ts INTEGER, end_iso TEXT,
            location TEXT, organizer TEXT, is_recurring INTEGER DEFAULT 0,
            my_response TEXT)""")
    store = CacheStore(db)
    with store._read() as conn:
        columns = {row[1] for row in conn.execute("PRAGMA table_info(events)")}
    assert "show_as" in columns


def test_shared_leader_cooldown_and_operation_journal(tmp_path):
    data_dir = str(tmp_path / "shared")
    one = ProcessCoordinator(data_dir, "exec@corp.example")
    two = ProcessCoordinator(data_dir, "exec@corp.example")
    try:
        assert one.is_sync_leader is not two.is_sync_leader
        one.set_cooldown(2, "test backoff")
        assert two.cooldown()["retry_after_s"] >= 1

        op, created = one.reserve_operation("create_event", "key-1", {"subject": "A"})
        assert created and op["status"] == "queued"
        result = one.execute_operation(op["operation_id"], lambda: {"raw_id": "E1"})
        assert result == {"raw_id": "E1"}
        replay, created = two.reserve_operation("create_event", "key-1", {"subject": "A"})
        assert not created and replay["status"] == "completed"
        with pytest.raises(ToolError) as conflict:
            two.reserve_operation("create_event", "key-1", {"subject": "B"})
        assert conflict.value.code == "idempotency_conflict"

        assert one.folder_revision() == 0
        assert one.bump_folder_revision() == 1
        assert two.folder_revision() == 1
        assert two.bump_folder_revision() == 2
        assert one.folder_revision() == 2
    finally:
        one.close()
        two.close()


def test_slow_calendar_timeout_does_not_starve_write_lane(tmp_path):
    settings = make_settings(ews_tz="Europe/Vienna", request_timeout=1)
    gateway = EWSGateway(settings)
    gateway._account = object()
    finished = threading.Event()

    def slow(_account):
        time.sleep(0.2)
        finished.set()

    async def scenario():
        calendar = asyncio.create_task(gateway.call(slow, lane="calendar", timeout_s=0.03))
        await asyncio.sleep(0.01)
        started = time.monotonic()
        write = await gateway.call(lambda _account: "written", lane="write", timeout_s=0.2)
        elapsed = time.monotonic() - started
        with pytest.raises(ToolError) as timed_out:
            await calendar
        assert timed_out.value.code == "upstream_timeout"
        return write, elapsed

    write, elapsed = asyncio.run(scenario())
    assert write == "written" and elapsed < 0.1
    assert finished.wait(1.0)  # isolated calendar worker eventually drains


def test_background_sync_is_not_clipped_by_single_request_timeout():
    settings = make_settings(ews_tz="Europe/Vienna", request_timeout=1)
    gateway = EWSGateway(settings)
    gateway._account = object()

    def multi_request_sync(_account):
        time.sleep(0.06)
        return "synced"

    async def scenario():
        # timeout_s=0 is the explicit background-job contract: individual EWS
        # requests remain bounded by exchangelib, but the aggregate job is not
        # incorrectly cut off at one request's deadline.
        return await gateway.call(
            multi_request_sync, lane="sync", timeout_s=0)

    assert asyncio.run(scenario()) == "synced"
