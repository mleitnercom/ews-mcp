"""Personal-contact CRUD, duplicate safety and destructive confirmation."""

import asyncio
from types import SimpleNamespace

from conftest import make_settings
from exchangelib.indexed_properties import EmailAddress, PhoneNumber, PhysicalAddress

from ewsmcp.audit import AuditLog
from ewsmcp.ids import get_aliaser
from ewsmcp.tools import calendar_people, contact_management
from ewsmcp.tools.base import Context, dispatch

SPECS = {spec.name: spec for spec in contact_management.TOOLS}


class ContactQuery(list):
    def only(self, *_fields):
        return self


class ContactsFolder:
    def __init__(self):
        self.items = ContactQuery()

    def all(self):
        return self.items


class FakeContact:
    next_id = 1

    def __init__(self, *, account=None, folder=None, raw_id=None,
                 given_name=None, surname=None, display_name=None,
                 file_as=None, email=None):
        self.account = account
        self.folder = folder
        self.id = raw_id
        self.changekey = "CK1"
        self.given_name = given_name
        self.surname = surname
        self.display_name = display_name or email or ""
        self.file_as = file_as
        self.email_addresses = (
            [EmailAddress(label="EmailAddress1", email=email)] if email else [])
        self.phone_numbers = []
        self.physical_addresses = []
        self.company_name = None
        self.job_title = None
        self.department = None
        self.office = None
        self.business_homepage = None
        self.categories = []
        self.saved_fields = []
        self.trashed = False
        self.permanently_deleted = False

    def save(self, update_fields=None):
        self.saved_fields.append(update_fields)
        if self.id is None:
            self.id = f"CONTACT-{FakeContact.next_id}"
            FakeContact.next_id += 1
            self.folder.items.append(self)
        self.changekey = f"{self.changekey}-saved"
        return self

    def move_to_trash(self):
        self.trashed = True
        self.folder.items.remove(self)

    def delete(self):
        self.permanently_deleted = True
        self.folder.items.remove(self)


def existing(folder, raw_id="CONTACT-OLD", name="Alex Example",
             email="alex@example.com"):
    contact = FakeContact(
        folder=folder, raw_id=raw_id, display_name=name, email=email,
        given_name=name.split()[0], surname=name.split()[-1],
    )
    folder.items.append(contact)
    return contact


class Gateway:
    def __init__(self, account):
        self.account = account
        self.calls = []

    async def call(self, fn, lane="read", timeout_s=None):
        self.calls.append(lane)
        return fn(self.account)


def ctx_for(tmp_path, **overrides):
    contacts = ContactsFolder()
    account = SimpleNamespace(contacts=contacts)

    def fetch(*, ids):
        raw = str(ids[0][0])
        return [item for item in contacts.items if str(item.id) == raw]

    account.fetch = fetch
    gateway = Gateway(account)
    ctx = Context(
        settings=make_settings(
            data_dir=str(tmp_path / "data"), ews_capability_tier="full", **overrides),
        gateway=gateway,
        manager=None,
        aliaser=get_aliaser(str(tmp_path / "alias")),
        audit=AuditLog(str(tmp_path / "audit")),
    )
    return ctx, account, gateway


def run(ctx, tool_name, **kwargs):
    return asyncio.run(dispatch(ctx, SPECS[tool_name], dict(kwargs)))


def test_pack_has_three_atomic_specs():
    assert set(SPECS) == {"create_contact", "update_contact", "delete_contact"}
    assert SPECS["create_contact"].side_effect_class == "write"
    assert SPECS["update_contact"].side_effect_class == "write"
    assert SPECS["delete_contact"].side_effect_class == "destructive"
    assert SPECS["delete_contact"].confirm is True


def test_create_contact_maps_business_fields_and_aliases(tmp_path, monkeypatch):
    FakeContact.next_id = 1
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, gateway = ctx_for(tmp_path)
    result = run(
        ctx, "create_contact", email="ALEX@EXAMPLE.COM",
        given_name="Alex", surname="Example", company="Example GmbH",
        job_title="Director", business_phone="+43 1 555",
        mobile_phone="+43 699 555", categories=["VIP", "Customer"],
        business_address={"street": "Main Street 1", "city": "Vienna",
                          "postal_code": "1010", "country": "Austria"},
    )
    assert result["ok"] is True and result["created"] is True
    assert result["contact"]["id"] == "p1"
    assert result["contact"]["email"] == "alex@example.com"
    assert result["contact"]["company"] == "Example GmbH"
    saved = account.contacts.items[0]
    assert saved.display_name == "Alex Example"
    assert {entry.label for entry in saved.phone_numbers} == {
        "BusinessPhone", "MobilePhone"}
    assert saved.physical_addresses[0].label == "Business"
    assert gateway.calls == ["write"]


def test_create_returns_exact_email_duplicate_unless_explicitly_allowed(
    tmp_path, monkeypatch,
):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, _gateway = ctx_for(tmp_path)
    original = existing(account.contacts)

    reused = run(ctx, "create_contact", email="ALEX@example.com")
    assert reused["ok"] is True and reused["created"] is False
    assert reused["contact"]["id"] == "p1"
    assert account.contacts.items == [original]

    created = run(
        ctx, "create_contact", email="alex@example.com", allow_duplicate=True)
    assert created["created"] is True
    assert len(account.contacts.items) == 2


def test_create_idempotency_key_replays_without_second_contact(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, _gateway = ctx_for(tmp_path)
    first = run(
        ctx, "create_contact", email="new@example.com", idempotency_key="contact-new")
    replay = run(
        ctx, "create_contact", email="new@example.com", idempotency_key="contact-new")
    assert first["created"] is True
    assert replay["idempotent_replay"] is True
    assert replay["contact"]["id"] == first["contact"]["id"]
    assert len(account.contacts.items) == 1


def test_contact_operation_status_never_leaks_raw_exchange_id(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, _account, _gateway = ctx_for(tmp_path)
    run(ctx, "create_contact", email="journal@example.com",
        idempotency_key="contact-journal")
    operation_id = ctx.coordinator.recent_operations(limit=1)[0]["operation_id"]
    status_spec = next(
        spec for spec in calendar_people.TOOLS if spec.name == "get_operation_status")
    status = asyncio.run(dispatch(
        ctx, status_spec, {"operation_id": operation_id}))
    stored = status["operation"]["result"]
    assert stored["item_id"].startswith("p")
    assert "raw_id" not in stored


def test_create_rejects_invalid_email(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, _account, _gateway = ctx_for(tmp_path)
    result = run(ctx, "create_contact", email="not-an-address")
    assert result["error"]["code"] == "validation"


def test_update_preserves_other_indexed_values_and_can_clear(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, gateway = ctx_for(tmp_path)
    contact = existing(account.contacts)
    contact.email_addresses.append(
        EmailAddress(label="EmailAddress2", email="private@example.net"))
    contact.phone_numbers = [
        PhoneNumber(label="HomePhone", phone_number="HOME"),
        PhoneNumber(label="BusinessPhone", phone_number="OLD"),
        PhoneNumber(label="MobilePhone", phone_number="MOBILE"),
    ]
    contact.physical_addresses = [PhysicalAddress(
        label="Business", street="Old Street", city="Old City")]
    alias = ctx.aliaser.alias_for(contact.id, "p")

    result = run(
        ctx, "update_contact", id=alias, email="new@example.com",
        business_phone="NEW", company="New Company",
        clear_fields=["mobile_phone", "business_address"],
    )
    assert result["ok"] is True
    assert result["contact"]["email"] == "new@example.com"
    assert result["contact"]["company"] == "New Company"
    assert {(entry.label, entry.email) for entry in contact.email_addresses} == {
        ("EmailAddress1", "new@example.com"),
        ("EmailAddress2", "private@example.net"),
    }
    assert {(entry.label, entry.phone_number) for entry in contact.phone_numbers} == {
        ("HomePhone", "HOME"), ("BusinessPhone", "NEW")}
    assert contact.physical_addresses == []
    assert gateway.calls == ["write"]


def test_update_rejects_noop_duplicate_and_gal_alias(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, _gateway = ctx_for(tmp_path)
    first = existing(account.contacts)
    existing(account.contacts, raw_id="CONTACT-OTHER", name="Other Person",
             email="other@example.com")
    alias = ctx.aliaser.alias_for(first.id, "p")

    noop = run(ctx, "update_contact", id=alias)
    assert noop["error"]["code"] == "validation"
    conflict = run(ctx, "update_contact", id=alias, email="other@example.com")
    assert conflict["error"]["code"] == "conflict"

    gal = ctx.aliaser.alias_for("directory@example.com", "p")
    protected = run(ctx, "update_contact", id=gal, company="Nope")
    assert protected["error"]["code"] == "protected_contact"


def test_delete_uses_two_phase_confirm_and_trash_by_default(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, _gateway = ctx_for(tmp_path)
    contact = existing(account.contacts)
    alias = ctx.aliaser.alias_for(contact.id, "p")

    phase1 = run(ctx, "delete_contact", id=alias)
    assert phase1["requires_confirmation"] is True
    assert phase1["preview"]["recoverable"] is True
    assert contact in account.contacts.items

    deleted = run(
        ctx, "delete_contact", id=alias, confirm_token=phase1["confirm_token"])
    assert deleted["ok"] is True
    assert deleted["deleted"]["recoverable"] is True
    assert contact.trashed is True
    assert contact not in account.contacts.items


def test_delete_confirmation_invalidates_when_contact_changes(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, _gateway = ctx_for(tmp_path)
    contact = existing(account.contacts)
    alias = ctx.aliaser.alias_for(contact.id, "p")
    phase1 = run(ctx, "delete_contact", id=alias)
    contact.display_name = "Changed after preview"
    contact.changekey = "CK2"

    stale = run(
        ctx, "delete_contact", id=alias, confirm_token=phase1["confirm_token"])
    assert stale["error"]["code"] == "confirm_invalid"
    assert contact in account.contacts.items


def test_delete_permanent_and_gal_protection(tmp_path, monkeypatch):
    monkeypatch.setattr(contact_management, "Contact", FakeContact)
    ctx, account, _gateway = ctx_for(tmp_path)
    contact = existing(account.contacts)
    alias = ctx.aliaser.alias_for(contact.id, "p")
    phase1 = run(ctx, "delete_contact", id=alias, disposition="permanent")
    deleted = run(
        ctx, "delete_contact", id=alias, disposition="permanent",
        confirm_token=phase1["confirm_token"],
    )
    assert deleted["deleted"]["recoverable"] is False
    assert contact.permanently_deleted is True

    gal = ctx.aliaser.alias_for("directory@example.com", "p")
    protected = run(ctx, "delete_contact", id=gal)
    assert protected["error"]["code"] == "protected_contact"
