"""Gateway probe/reset behavior — the false-warm fix (Phase B critical #4).

``inbox.total_count`` is a cached property after its first read, so probing
it reported warm forever (the /readyz lie). The probe must round-trip every
call, and reset() must evict exchangelib's protocol cache so a wedged
session actually renegotiates auth.
"""

from unittest.mock import MagicMock

from conftest import make_settings

from ewsmcp.gateway import client as client_mod
from ewsmcp.gateway.client import EWSGateway


def _gateway_with_mock_account():
    gw = EWSGateway(make_settings())
    account = MagicMock(name="account")
    gw._account = account
    return gw, account


def test_probe_round_trips_on_every_call():
    gw, account = _gateway_with_mock_account()
    assert gw.test_connection() is True
    assert gw.test_connection() is True
    assert account.root.refresh.call_count == 2  # not a cached-property read


def test_probe_failure_is_reported_not_cached():
    gw, account = _gateway_with_mock_account()
    account.root.refresh.side_effect = ConnectionError("front door down")
    assert gw.test_connection() is False
    assert "front door down" in gw.last_connection_error
    account.root.refresh.side_effect = None
    assert gw.test_connection() is True  # recovery observed immediately
    assert gw.last_connection_error is None


def test_reset_evicts_protocol_cache_and_folder_cache(monkeypatch):
    cleared = []
    monkeypatch.setattr(client_mod.CachingProtocol, "clear_cache",
                        lambda: cleared.append(True))
    gw, account = _gateway_with_mock_account()
    gw._folder_cache = {"path": object()}
    gw._folder_cache_ts = 123.0
    gw._folder_revision_seen = 7
    gw.reset()
    assert gw._account is None
    account.protocol.close.assert_called_once()
    assert cleared == [True]  # without this, the wedged Protocol comes back
    assert gw._folder_cache == {}
    assert gw._folder_cache_ts == 0.0
    assert gw._folder_revision_seen is None


def test_reset_survives_close_and_clear_failures(monkeypatch):
    monkeypatch.setattr(
        client_mod.CachingProtocol, "clear_cache",
        lambda: (_ for _ in ()).throw(RuntimeError("cache locked")))
    gw, account = _gateway_with_mock_account()
    account.protocol.close.side_effect = OSError("socket gone")
    gw.reset()  # must not raise
    assert gw._account is None


def test_remote_folder_revision_clears_exchangelib_and_gateway_caches():
    class Coordinator:
        revision = 1

        def folder_revision(self):
            return self.revision

        def bump_folder_revision(self):
            self.revision += 1
            return self.revision

    coordinator = Coordinator()
    gw = EWSGateway(make_settings(), coordinator=coordinator)
    root = MagicMock(name="msg_folder_root")
    account = MagicMock(name="account")
    account.msg_folder_root = root
    gw._folder_revision_seen = 1
    gw._folder_cache = {"stale": object()}
    gw._folder_cache_ts = 123.0

    coordinator.revision = 2
    gw.prepare_folder_tree(account)

    root.clear_cache.assert_called_once()
    assert gw._folder_cache == {}
    assert gw._folder_cache_ts == 0.0
    assert gw._folder_revision_seen == 2


def test_local_folder_mutation_publishes_revision_without_clearing_root():
    class Coordinator:
        revision = 0

        def bump_folder_revision(self):
            self.revision += 1
            return self.revision

    coordinator = Coordinator()
    gw = EWSGateway(make_settings(), coordinator=coordinator)
    account = MagicMock(name="account")
    gw._folder_cache = {"old": object()}

    assert gw.mark_folder_tree_changed(account) == 1
    assert gw._folder_revision_seen == 1
    assert gw._folder_cache == {}
    account.msg_folder_root.clear_cache.assert_not_called()

# --- recovery after a degraded probe -----------------------------------------


def test_degraded_recovery_drops_the_session_pool_before_reconnecting():
    """A probe failure means the cached Protocol's session pool may hold
    sockets that died with the connection. Reconnecting does not discard
    them, so the next caller draws a dead session and blocks until the OS
    times out the TCP connection — minutes, not seconds.

    The old ladder only reset every third FAILED attempt, so a recovery that
    succeeded on the first attempt kept the poisoned pool. Observed live on
    2026-10-09: probe failed 07:49:31, reconnect succeeded 07:49:31, and
    calls hung for ~2 minutes each for hours afterwards.
    """
    import asyncio

    from ewsmcp.gateway.connection import ConnectionManager

    order = []
    client = MagicMock(name="client")
    client.reset.side_effect = lambda: order.append("reset")
    client.test_connection.side_effect = lambda: (order.append("connect"), True)[1]

    manager = ConnectionManager(client, heartbeat_seconds=3600)
    try:
        asyncio.run(manager._warmup_loop(reset_first=True))
    finally:
        manager._stopped = True

    assert order[:2] == ["reset", "connect"], (
        f"pool must be dropped BEFORE the reconnect attempt, got {order}")
    assert manager.state == "warm"


def test_plain_warmup_does_not_reset_on_the_first_attempt():
    """Boot is not a recovery: a cold start has no stale pool to discard,
    and resetting there would throw away the negotiation it just did."""
    import asyncio

    from ewsmcp.gateway.connection import ConnectionManager

    client = MagicMock(name="client")
    client.test_connection.return_value = True
    manager = ConnectionManager(client, heartbeat_seconds=3600)
    try:
        asyncio.run(manager._warmup_loop())
    finally:
        manager._stopped = True

    client.reset.assert_not_called()

