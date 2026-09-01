"""Folder-management tools: CRUD, safety gates and mirror coherence."""

import asyncio
from types import SimpleNamespace

from conftest import make_settings

from ewsmcp.audit import AuditLog
from ewsmcp.cache.store import CacheStore
from ewsmcp.errors import ToolError
from ewsmcp.ids import get_aliaser
from ewsmcp.tools import folder_management
from ewsmcp.tools.base import Context, dispatch

SPECS = {spec.name: spec for spec in folder_management.TOOLS}


class FakeFolder:
    next_id = 1

    def __init__(self, *, parent=None, name="", folder_class="IPF.Note",
                 raw_id=None, distinguished=False, total=0):
        self.parent = parent
        self.name = name
        self.folder_class = folder_class
        self.id = raw_id
        self.changekey = "CK1"
        self.is_distinguished = distinguished
        self.total_count = total
        self.unread_count = 0
        self.children = []
        self.deleted_with = None
        self.saved_fields = []

    def save(self, update_fields=None):
        self.saved_fields.append(update_fields)
        if self.id is None:
            self.id = f"NEW-{FakeFolder.next_id}"
            FakeFolder.next_id += 1
            self.parent.children.append(self)
        self.changekey = f"{self.changekey}-saved"
        return self

    def move(self, to_folder):
        self.parent.children.remove(self)
        self.parent = to_folder
        to_folder.children.append(self)
        self.changekey = f"{self.changekey}-moved"

    def delete(self, delete_type):
        self.deleted_with = delete_type
        self.parent.children.remove(self)
        self.id = None


def attach(parent, child):
    child.parent = parent
    parent.children.append(child)
    return child


def mailbox():
    root = FakeFolder(name="root", raw_id="ROOT", folder_class="IPM.Root")
    inbox = attach(root, FakeFolder(
        name="Inbox", raw_id="INBOX", distinguished=True))
    trash = attach(root, FakeFolder(
        name="Deleted Items", raw_id="TRASH", distinguished=True))
    projects = attach(root, FakeFolder(name="Projects", raw_id="PROJECTS"))
    archive = attach(root, FakeFolder(name="Archive", raw_id="ARCHIVE"))
    return SimpleNamespace(
        msg_folder_root=root, inbox=inbox, trash=trash,
        projects=projects, archive=archive,
    )


class Gateway:
    def __init__(self, account):
        self.account = account
        self.revision = 0
        self.calls = []

    async def call(self, fn, lane="read", timeout_s=None):
        self.calls.append(lane)
        return fn(self.account)

    def resolve_folder(self, account, ref, aliaser):
        if ref in (None, "f:inbox"):
            return account.inbox
        if ref == "f:root":
            return account.msg_folder_root
        if ref == "f:trash":
            return account.trash
        try:
            key = aliaser.resolve(ref)
        except KeyError as exc:
            raise ToolError("validation", str(exc)) from exc
        matches = {}

        def walk(node, prefix=""):
            for child in node.children:
                path = f"{prefix}/{child.name}" if prefix else child.name
                matches[str(child.id)] = child
                matches[path.casefold()] = child
                walk(child, path)

        walk(account.msg_folder_root)
        found = matches.get(key) or matches.get(str(key).casefold())
        if found is None:
            raise ToolError("not_found", f"No folder matches {ref!r}.")
        return found

    def mark_folder_tree_changed(self, _account):
        self.revision += 1
        return self.revision


def ctx_for(tmp_path, account=None, **overrides):
    account = account or mailbox()
    gateway = Gateway(account)
    cache = CacheStore(tmp_path / "mirror.db")
    ctx = Context(
        settings=make_settings(ews_capability_tier="full", **overrides),
        gateway=gateway,
        manager=None,
        aliaser=get_aliaser(str(tmp_path / "alias")),
        audit=AuditLog(str(tmp_path / "audit")),
        cache=cache,
    )
    return ctx, account, gateway, cache


def run(ctx, tool_name, **kwargs):
    return asyncio.run(dispatch(ctx, SPECS[tool_name], dict(kwargs)))


def test_pack_has_four_atomic_specs():
    assert set(SPECS) == {
        "create_folder", "rename_folder", "move_folder", "delete_folder"}
    assert SPECS["create_folder"].side_effect_class == "write"
    assert SPECS["rename_folder"].side_effect_class == "write"
    assert SPECS["move_folder"].side_effect_class == "write"
    assert SPECS["delete_folder"].side_effect_class == "destructive"
    assert SPECS["delete_folder"].confirm is True


def test_create_is_idempotent_by_exact_sibling_and_refreshes_mirror(tmp_path,
                                                                    monkeypatch):
    FakeFolder.next_id = 1
    monkeypatch.setattr(folder_management, "Folder", FakeFolder)
    ctx, account, gateway, cache = ctx_for(tmp_path)

    created = run(ctx, "create_folder", name="Board", parent="f:root")
    assert created["ok"] is True and created["created"] is True
    assert created["folder"]["path"] == "Board"
    assert created["folder"]["parent"] == "f:root"
    assert created["folder_revision"] == 1
    assert {row["path"] for row in cache.folder_rows()} >= {"Board", "Projects"}

    replay = run(ctx, "create_folder", name="board", parent="f:root")
    assert replay["ok"] is True and replay["created"] is False
    assert len([f for f in account.msg_folder_root.children
                if f.name.casefold() == "board"]) == 1
    assert gateway.calls.count("write") == 2


def test_create_rejects_path_separator(tmp_path):
    ctx, _account, _gateway, _cache = ctx_for(tmp_path)
    result = run(ctx, "create_folder", name="Parent/Child")
    assert result["error"]["code"] == "validation"


def test_rename_is_noop_or_conflict_and_protects_system_folders(tmp_path):
    ctx, account, gateway, _cache = ctx_for(tmp_path)
    no_op = run(ctx, "rename_folder", folder="PROJECTS", new_name="Projects")
    assert no_op["changed"] is False and gateway.revision == 0

    conflict = run(ctx, "rename_folder", folder="PROJECTS", new_name="Archive")
    assert conflict["error"]["code"] == "conflict"

    protected = run(ctx, "rename_folder", folder="f:inbox", new_name="Incoming")
    assert protected["error"]["code"] == "protected_folder"

    renamed = run(ctx, "rename_folder", folder="PROJECTS", new_name="Initiatives")
    assert renamed["changed"] is True
    assert account.projects.name == "Initiatives"
    assert renamed["previous"] == {"name": "Projects", "path": "Projects"}


def test_move_rejects_descendant_and_moves_to_new_parent(tmp_path):
    ctx, account, gateway, _cache = ctx_for(tmp_path)
    child = attach(account.projects, FakeFolder(name="2026", raw_id="YEAR"))

    cycle = run(ctx, "move_folder", folder="PROJECTS", parent="YEAR")
    assert cycle["error"]["code"] == "validation"

    moved = run(ctx, "move_folder", folder="YEAR", parent="ARCHIVE")
    assert moved["changed"] is True
    assert child.parent is account.archive
    assert moved["folder"]["path"] == "Archive/2026"
    assert gateway.revision == 1


def test_delete_uses_two_phase_confirm_and_recoverable_default(tmp_path):
    ctx, account, gateway, _cache = ctx_for(tmp_path)
    target = attach(account.msg_folder_root, FakeFolder(name="Empty", raw_id="EMPTY"))

    phase1 = run(ctx, "delete_folder", folder="EMPTY")
    assert phase1["requires_confirmation"] is True
    assert phase1["preview"]["recoverable"] is True
    assert target in account.msg_folder_root.children

    deleted = run(ctx, "delete_folder", folder="EMPTY",
                  confirm_token=phase1["confirm_token"])
    assert deleted["ok"] is True
    assert deleted["deleted"]["recoverable"] is True
    assert target.deleted_with == folder_management.MOVE_TO_DELETED_ITEMS
    assert target not in account.msg_folder_root.children
    assert gateway.revision == 1


def test_delete_non_empty_requires_ack_and_permanent_is_hard_delete(tmp_path):
    ctx, account, _gateway, _cache = ctx_for(tmp_path)
    target = attach(account.msg_folder_root, FakeFolder(
        name="Old", raw_id="OLD", total=3))

    refused = run(ctx, "delete_folder", folder="OLD", disposition="permanent")
    assert refused["error"]["code"] == "validation"
    assert "allow_non_empty" in refused["error"]["hint"]

    phase1 = run(ctx, "delete_folder", folder="OLD", disposition="permanent",
                 allow_non_empty=True)
    assert phase1["preview"]["items"] == 3
    deleted = run(ctx, "delete_folder", folder="OLD", disposition="permanent",
                  allow_non_empty=True, confirm_token=phase1["confirm_token"])
    assert deleted["deleted"]["recoverable"] is False
    assert target.deleted_with == folder_management.HARD_DELETE


def test_delete_confirmation_is_invalidated_when_contents_change(tmp_path):
    ctx, account, _gateway, _cache = ctx_for(tmp_path)
    target = attach(account.msg_folder_root, FakeFolder(name="Temp", raw_id="TEMP"))
    phase1 = run(ctx, "delete_folder", folder="TEMP", allow_non_empty=True)
    target.total_count = 1
    stale = run(ctx, "delete_folder", folder="TEMP",
                confirm_token=phase1["confirm_token"], allow_non_empty=True)
    # Arguments are identical; only the resolved folder contents changed.
    assert stale["error"]["code"] == "confirm_invalid"
    assert target in account.msg_folder_root.children


def test_delete_protects_distinguished_folder_before_preview(tmp_path):
    ctx, _account, _gateway, _cache = ctx_for(tmp_path)
    result = run(ctx, "delete_folder", folder="f:inbox")
    assert result["error"]["code"] == "protected_folder"
