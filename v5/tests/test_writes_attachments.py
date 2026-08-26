"""Outbound draft attachments (create_draft/update_draft ``attachments``).

Restores the v4 capability (``add_attachment``, ``attach_email_to_draft``)
that the 4.5 rewrite dropped, as a parameter on the draft tools rather than
as separate tools — draft-first stays the only authoring path.

Two invariants are worth pinning: attachments are applied only AFTER the
item is saved (an unsaved ``attach()`` never reaches EWS — the exact shape
of the v3 bug fixed in v4), and every path is confined to the configured
roots BEFORE any file is read.
"""
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

from ewsmcp.tools import writes

from test_writes import FakeDraftMessage, call, make_account, make_ctx


def _ctx(tmp_path, account, **overrides):
    root = tmp_path / "outbox"
    root.mkdir(exist_ok=True)
    overrides.setdefault("attach_allowed_dirs", str(root))
    return make_ctx(tmp_path, account, **overrides), root


def _file(root: Path, name: str, data: bytes = b"payload") -> str:
    p = root / name
    p.write_bytes(data)
    return str(p)


# --- create_draft ------------------------------------------------------------


def test_create_draft_attaches_after_save(tmp_path, monkeypatch):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    monkeypatch.setattr(writes, "Message", FakeDraftMessage)
    path = _file(root, "register.pdf", b"%PDF-1.7 x")

    res = call(ctx, "create_draft",
               {"to": ["legal@corp.example"], "subject": "Register",
                "body": "anbei", "attachments": [path]})

    msg = FakeDraftMessage.last
    assert res["ok"] is True
    assert msg.save_calls == [{}]                                 # saved first …
    assert [a.name for a in msg.attached] == ["register.pdf"]     # … then attached
    assert msg.attached[0].content == b"%PDF-1.7 x"
    assert msg.attached[0].content_type == "application/pdf"
    assert res["preview"]["attachments"] == [
        {"name": "register.pdf", "size_bytes": 10}]


def test_create_draft_without_attachments_is_unchanged(tmp_path, monkeypatch):
    account = make_account()
    ctx, _ = _ctx(tmp_path, account)
    monkeypatch.setattr(writes, "Message", FakeDraftMessage)
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "subject": "Hi", "body": "Hello"})
    assert FakeDraftMessage.last.attached == []
    assert "attachments" not in res["preview"]


def test_create_draft_forward_attaches_to_the_saved_reply(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    original = MagicMock()
    original.subject = "Belege"
    fwd = MagicMock()
    fwd.to_recipients = [SimpleNamespace(email_address="m.h@corp.example")]
    fwd.cc_recipients = None
    saved = MagicMock()
    saved.id, saved.changekey = "RAW-FWD", "CK-FWD"
    fwd.save.return_value = saved
    original.create_forward.return_value = fwd
    account._by_id["RAW-ORIG"] = original
    alias = ctx.aliaser.alias_for("RAW-ORIG", "m")
    path = _file(root, "beleg.pdf")

    res = call(ctx, "create_draft",
               {"mode": "forward", "reply_to": alias, "to": ["m.h@corp.example"],
                "body": "anbei", "attachments": [path]})

    assert res["ok"] is True
    assert saved.attach.call_count == 1     # the SAVED item, not the template
    assert fwd.attach.call_count == 0


# --- update_draft ------------------------------------------------------------


def test_update_draft_attachments_only_skips_the_field_update(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    draft = MagicMock()
    draft.subject = "Register"
    draft.to_recipients = []
    draft.cc_recipients = []
    account._by_id["RAW-DRAFT"] = draft
    alias = ctx.aliaser.alias_for("RAW-DRAFT", "d")

    res = call(ctx, "update_draft",
               {"draft_id": alias, "attachments": [_file(root, "anlage.xlsx")]})

    draft.save.assert_not_called()          # no empty update_fields=[] call
    assert draft.attach.call_count == 1
    assert res["preview"]["attachments_added"] == [
        {"name": "anlage.xlsx", "size_bytes": 7}]


def test_update_draft_saves_fields_before_attaching(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    order = []
    draft = MagicMock()
    draft.subject = "neu"
    draft.to_recipients = []
    draft.cc_recipients = []
    draft.save.side_effect = lambda **kw: order.append("save")
    draft.attach.side_effect = lambda a: order.append("attach")
    account._by_id["RAW-DRAFT"] = draft
    alias = ctx.aliaser.alias_for("RAW-DRAFT", "d")

    call(ctx, "update_draft",
         {"draft_id": alias, "subject": "neu",
          "attachments": [_file(root, "a.txt")]})

    assert order == ["save", "attach"]      # attach() rolls the changekey last


def test_update_draft_still_rejects_an_empty_call(tmp_path):
    account = make_account()
    ctx, _ = _ctx(tmp_path, account)
    res = call(ctx, "update_draft", {"draft_id": "d1"})
    assert res["ok"] is False and res["error"]["code"] == "validation"


# --- containment and caps ----------------------------------------------------


def test_path_outside_the_allowed_roots_is_refused(tmp_path, monkeypatch):
    account = make_account()
    ctx, _ = _ctx(tmp_path, account)
    monkeypatch.setattr(writes, "Message", FakeDraftMessage)
    outside = tmp_path / "secrets.txt"
    outside.write_bytes(b"nope")
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x", "attachments": [str(outside)]})
    assert res["ok"] is False and res["error"]["code"] == "validation"
    assert "outside the allowed roots" in res["error"]["message"]


def test_traversal_out_of_an_allowed_root_is_refused(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    (tmp_path / "secrets.txt").write_bytes(b"nope")
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x",
                "attachments": [str(root / ".." / "secrets.txt")]})
    assert res["ok"] is False and res["error"]["code"] == "validation"


def test_data_dir_is_always_an_allowed_root(tmp_path, monkeypatch):
    """get_attachment(mode='save') writes under DATA_DIR/attachments — the
    download then re-attach round trip must work with no extra config."""
    account = make_account()
    ctx = make_ctx(tmp_path, account)          # no attach_allowed_dirs at all
    monkeypatch.setattr(writes, "Message", FakeDraftMessage)
    saved = Path(ctx.settings.data_dir) / "attachments"
    saved.mkdir(parents=True, exist_ok=True)
    path = _file(saved, "vom-server.pdf")
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x", "attachments": [path]})
    assert res["preview"]["attachments"][0]["name"] == "vom-server.pdf"


def test_missing_file_is_not_found(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x",
                "attachments": [str(root / "ghost.pdf")]})
    assert res["error"]["code"] == "not_found"


def test_directory_is_refused(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account)
    (root / "folder").mkdir()
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x",
                "attachments": [str(root / "folder")]})
    assert "not a file" in res["error"]["message"]


def test_file_count_cap(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account, attach_max_files=2)
    paths = [_file(root, f"f{i}.txt") for i in range(3)]
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x", "attachments": paths})
    assert "at most 2 attachments" in res["error"]["message"]


def test_total_size_cap_counts_across_files(tmp_path):
    account = make_account()
    ctx, root = _ctx(tmp_path, account, attach_max_bytes=10)
    paths = [_file(root, "a.bin", b"123456"), _file(root, "b.bin", b"123456")]
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x", "attachments": paths})
    assert "byte cap" in res["error"]["message"]


def test_bad_path_fails_before_a_draft_is_created(tmp_path, monkeypatch):
    """Validation happens outside gateway.call — no orphan draft on failure."""
    account = make_account()
    ctx, _ = _ctx(tmp_path, account)
    monkeypatch.setattr(writes, "Message", FakeDraftMessage)
    FakeDraftMessage.last = None
    res = call(ctx, "create_draft",
               {"to": ["a@x.com"], "body": "x", "attachments": ["/etc/passwd"]})
    assert res["ok"] is False
    assert FakeDraftMessage.last is None
