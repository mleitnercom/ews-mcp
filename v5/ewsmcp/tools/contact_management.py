"""Personal-contact create, update and delete tools.

GAL/directory people are read-only.  These tools operate exclusively on the
authenticated mailbox's personal Contacts folder and keep destructive delete
behind the dispatcher's content-bound two-phase confirmation gate.
"""

from __future__ import annotations

import logging
import re
from itertools import islice
from typing import Any

from exchangelib import Contact
from exchangelib.indexed_properties import EmailAddress, PhoneNumber, PhysicalAddress

from ..contacts import contact_details, replace_indexed
from ..errors import ToolError
from .base import Context, ToolSpec
from .writes import _gateway_write, _idempotent_create

MAX_CONTACT_SCAN = 2000
MAX_TEXT = 255
MAX_CATEGORIES = 50
EMAIL_LABEL = "EmailAddress1"
BUSINESS_PHONE_LABEL = "BusinessPhone"
MOBILE_PHONE_LABEL = "MobilePhone"
BUSINESS_ADDRESS_LABEL = "Business"

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+$")
_DIRECT_FIELDS = {
    "company": "company_name",
    "job_title": "job_title",
    "department": "department",
    "office": "office",
    "business_homepage": "business_homepage",
}
logger = logging.getLogger(__name__)

_CLEARABLE_FIELDS = (
    "business_phone", "mobile_phone", *_DIRECT_FIELDS,
    "business_address", "categories",
)


def _text(value: Any, field: str, *, limit: int = MAX_TEXT) -> str:
    clean = str(value or "").strip()
    if not clean:
        raise ToolError("validation", f"'{field}' cannot be empty")
    if len(clean) > limit:
        raise ToolError("validation", f"'{field}' cannot exceed {limit} characters")
    if any(ord(char) < 32 for char in clean):
        raise ToolError("validation", f"'{field}' cannot contain control characters")
    return clean


def _email(value: Any) -> str:
    clean = _text(value, "email").lower()
    if not _EMAIL_RE.fullmatch(clean):
        raise ToolError(
            "validation", f"invalid email address: {value!r}",
            hint="Use one SMTP address such as person@example.com.",
        )
    return clean


def _categories(values: Any) -> list[str]:
    if not isinstance(values, list):
        raise ToolError("validation", "'categories' must be an array of strings")
    if len(values) > MAX_CATEGORIES:
        raise ToolError(
            "validation", f"at most {MAX_CATEGORIES} contact categories are allowed")
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        category = _text(value, "categories[]")
        key = category.casefold()
        if key not in seen:
            seen.add(key)
            out.append(category)
    return out


def _address(value: Any) -> dict[str, str]:
    if not isinstance(value, dict):
        raise ToolError("validation", "'business_address' must be an object")
    allowed = ("street", "city", "state", "postal_code", "country")
    extra = sorted(set(value) - set(allowed))
    if extra:
        raise ToolError(
            "validation", f"unsupported business_address fields: {', '.join(extra)}")
    result = {key: _text(raw, f"business_address.{key}")
              for key, raw in value.items() if raw is not None}
    if not result:
        raise ToolError(
            "validation",
            "'business_address' must contain at least one address field; "
            "use clear_fields=['business_address'] to remove it",
        )
    return result


def _display_name(given_name: str | None, surname: str | None,
                  explicit: str | None, email: str) -> str:
    if explicit:
        return explicit
    joined = " ".join(part for part in (given_name, surname) if part).strip()
    return joined or email.split("@", 1)[0]


def _raw_id(contact: Any) -> str:
    raw = getattr(contact, "id", None)
    if not raw:
        raise ToolError("upstream_error", "Exchange returned a contact without an id.")
    return str(raw)


def _public_contact(ctx: Context, raw_id: str,
                    details: dict[str, Any]) -> dict[str, Any]:
    return {"id": ctx.aliaser.alias_for(raw_id, "p"), **details}


def _ensure_personal_ref(raw_id: str) -> None:
    # GAL p-aliases are intentionally bound to the SMTP address because GAL
    # objects have no personal Contact item id and must never be mutated.
    if "@" in raw_id:
        raise ToolError(
            "protected_contact",
            "Directory/GAL entries are read-only and cannot be changed here.",
            hint="Create a personal contact first, then use its contacts-source p-alias.",
        )


def _fetch_contact(account: Any, raw_id: str) -> Any:
    _ensure_personal_ref(raw_id)
    fetched = list(account.fetch(ids=[(raw_id, None)]))
    if not fetched:
        raise ToolError("not_found", "Personal contact not found; the id may be stale.",
                        hint="Run find_people(source='contacts') for a fresh p-alias.")
    contact = fetched[0]
    if isinstance(contact, Exception):
        raise contact
    if not isinstance(contact, Contact):
        raise ToolError(
            "validation", "The supplied id does not refer to a personal contact.",
            hint="Use a p-alias returned with source='contacts', not a GAL result.",
        )
    return contact


def _primary_email(contact: Any) -> str | None:
    entries = list(getattr(contact, "email_addresses", None) or [])
    primary = next(
        (entry for entry in entries if str(getattr(entry, "label", "")) == EMAIL_LABEL),
        entries[0] if entries else None,
    )
    value = getattr(primary, "email", None) if primary is not None else None
    return str(value).lower() if value else None


def _duplicate_by_email(account: Any, email: str,
                        *, exclude_id: str | None = None) -> Any | None:
    contacts = account.contacts.all()
    try:
        contacts = contacts.only("display_name", "email_addresses", "given_name",
                                 "surname", "company_name", "job_title",
                                 "department", "office", "business_homepage",
                                 "phone_numbers", "physical_addresses", "categories")
    except Exception as exc:  # noqa: BLE001 — projection is an optimization only
        logger.debug("contact projection unavailable: %s", exc)
    for contact in islice(contacts, MAX_CONTACT_SCAN):
        raw = str(getattr(contact, "id", None) or "")
        if raw and raw != exclude_id and _primary_email(contact) == email:
            return contact
    return None


def _prepared_fields(**values: Any) -> dict[str, Any]:
    prepared: dict[str, Any] = {}
    for key, value in values.items():
        if value is None:
            continue
        if key == "email":
            prepared[key] = _email(value)
        elif key == "categories":
            prepared[key] = _categories(value)
        elif key == "business_address":
            prepared[key] = _address(value)
        else:
            prepared[key] = _text(value, key)
    return prepared


def _new_address(value: dict[str, str]) -> PhysicalAddress:
    return PhysicalAddress(
        label=BUSINESS_ADDRESS_LABEL,
        street=value.get("street"), city=value.get("city"),
        state=value.get("state"), zipcode=value.get("postal_code"),
        country=value.get("country"),
    )


async def _create_contact(
    ctx: Context, *, email: str, given_name: str | None = None,
    surname: str | None = None, display_name: str | None = None,
    business_phone: str | None = None, mobile_phone: str | None = None,
    company: str | None = None, job_title: str | None = None,
    department: str | None = None, office: str | None = None,
    business_homepage: str | None = None,
    business_address: dict[str, Any] | None = None,
    categories: list[str] | None = None, allow_duplicate: bool = False,
    idempotency_key: str | None = None,
) -> dict[str, Any]:
    prepared = _prepared_fields(
        email=email, given_name=given_name, surname=surname,
        display_name=display_name, business_phone=business_phone,
        mobile_phone=mobile_phone, company=company, job_title=job_title,
        department=department, office=office,
        business_homepage=business_homepage,
        business_address=business_address, categories=categories,
    )
    clean_email = prepared["email"]
    display = _display_name(
        prepared.get("given_name"), prepared.get("surname"),
        prepared.get("display_name"), clean_email,
    )

    def work(account: Any) -> dict[str, Any]:
        if not allow_duplicate:
            existing = _duplicate_by_email(account, clean_email)
            if existing is not None:
                return {
                    "raw_id": _raw_id(existing), "created": False,
                    "contact": contact_details(existing),
                }
        contact = Contact(
            account=account, folder=account.contacts,
            given_name=prepared.get("given_name"),
            surname=prepared.get("surname"), display_name=display,
            file_as=display,
        )
        contact.email_addresses = [EmailAddress(label=EMAIL_LABEL, email=clean_email)]
        phones = []
        if prepared.get("business_phone"):
            phones.append(PhoneNumber(
                label=BUSINESS_PHONE_LABEL,
                phone_number=prepared["business_phone"],
            ))
        if prepared.get("mobile_phone"):
            phones.append(PhoneNumber(
                label=MOBILE_PHONE_LABEL,
                phone_number=prepared["mobile_phone"],
            ))
        if phones:
            contact.phone_numbers = phones
        for public, attr in _DIRECT_FIELDS.items():
            if prepared.get(public):
                setattr(contact, attr, prepared[public])
        if prepared.get("business_address"):
            contact.physical_addresses = [_new_address(prepared["business_address"])]
        if "categories" in prepared:
            contact.categories = prepared["categories"]
        contact.save()
        return {
            "raw_id": _raw_id(contact), "created": True,
            "contact": contact_details(contact),
        }

    payload = {**prepared, "allow_duplicate": bool(allow_duplicate)}
    outcome = await _idempotent_create(
        ctx, "create_contact", idempotency_key, payload, work)
    if outcome["state"] != "completed":
        return {
            "ok": False,
            "error": outcome.get("error", {
                "code": "operation_in_progress",
                "message": "Contact create is still running; retry with the same key.",
            }),
            "operation": outcome["operation"],
        }
    created = outcome["result"]
    result: dict[str, Any] = {
        "ok": True, "created": bool(created["created"]),
        "contact": _public_contact(ctx, created["raw_id"], created["contact"]),
    }
    if outcome.get("replayed"):
        result["idempotent_replay"] = True
    return result


async def _update_contact(
    ctx: Context, *, id: str, email: str | None = None,
    given_name: str | None = None, surname: str | None = None,
    display_name: str | None = None, business_phone: str | None = None,
    mobile_phone: str | None = None, company: str | None = None,
    job_title: str | None = None, department: str | None = None,
    office: str | None = None, business_homepage: str | None = None,
    business_address: dict[str, Any] | None = None,
    categories: list[str] | None = None,
    clear_fields: list[str] | None = None, allow_duplicate: bool = False,
) -> dict[str, Any]:
    prepared = _prepared_fields(
        email=email, given_name=given_name, surname=surname,
        display_name=display_name, business_phone=business_phone,
        mobile_phone=mobile_phone, company=company, job_title=job_title,
        department=department, office=office,
        business_homepage=business_homepage,
        business_address=business_address, categories=categories,
    )
    clears = list(dict.fromkeys(clear_fields or []))
    invalid = [field for field in clears if field not in _CLEARABLE_FIELDS]
    if invalid:
        raise ToolError("validation", f"unsupported clear_fields: {', '.join(invalid)}")
    overlap = sorted(set(clears) & set(prepared))
    if overlap:
        raise ToolError(
            "validation", f"fields cannot be set and cleared together: {', '.join(overlap)}")
    if not prepared and not clears:
        raise ToolError("validation", "Provide at least one field to update or clear.")

    def work(account: Any) -> dict[str, Any]:
        contact = _fetch_contact(account, id)
        raw_id = _raw_id(contact)
        if prepared.get("email") and not allow_duplicate:
            duplicate = _duplicate_by_email(
                account, prepared["email"], exclude_id=raw_id)
            if duplicate is not None:
                raise ToolError(
                    "conflict",
                    f"Another personal contact already uses {prepared['email']!r}.",
                    hint="Use that contact, or retry with allow_duplicate=true if intentional.",
                )
        changed_fields: list[str] = []

        def changed(field: str) -> None:
            if field not in changed_fields:
                changed_fields.append(field)

        for public, attr in (("given_name", "given_name"), ("surname", "surname")):
            if public in prepared:
                setattr(contact, attr, prepared[public])
                changed(attr)
        if "display_name" in prepared or {"given_name", "surname"} & set(prepared):
            display = _display_name(
                getattr(contact, "given_name", None), getattr(contact, "surname", None),
                prepared.get("display_name"),
                prepared.get("email") or _primary_email(contact) or "contact",
            )
            contact.display_name = display
            contact.file_as = display
            changed("display_name")
            changed("file_as")
        if "email" in prepared:
            contact.email_addresses = replace_indexed(
                getattr(contact, "email_addresses", None), EMAIL_LABEL,
                EmailAddress(label=EMAIL_LABEL, email=prepared["email"]),
            )
            changed("email_addresses")
        for public, label in (("business_phone", BUSINESS_PHONE_LABEL),
                              ("mobile_phone", MOBILE_PHONE_LABEL)):
            if public in prepared or public in clears:
                replacement = None if public in clears else PhoneNumber(
                    label=label, phone_number=prepared[public])
                contact.phone_numbers = replace_indexed(
                    getattr(contact, "phone_numbers", None), label, replacement)
                changed("phone_numbers")
        for public, attr in _DIRECT_FIELDS.items():
            if public in prepared or public in clears:
                setattr(contact, attr, None if public in clears else prepared[public])
                changed(attr)
        if "business_address" in prepared or "business_address" in clears:
            replacement = (None if "business_address" in clears
                           else _new_address(prepared["business_address"]))
            contact.physical_addresses = replace_indexed(
                getattr(contact, "physical_addresses", None),
                BUSINESS_ADDRESS_LABEL, replacement,
            )
            changed("physical_addresses")
        if "categories" in prepared or "categories" in clears:
            contact.categories = [] if "categories" in clears else prepared["categories"]
            changed("categories")
        contact.save(update_fields=changed_fields)
        return {"raw_id": raw_id, "contact": contact_details(contact),
                "updated_fields": changed_fields}

    updated = await _gateway_write(ctx, work)
    return {
        "ok": True,
        "contact": _public_contact(ctx, updated["raw_id"], updated["contact"]),
        "updated_fields": updated["updated_fields"],
    }


def _delete_snapshot(ctx: Context, account: Any, contact_ref: str,
                     disposition: str) -> dict[str, Any]:
    if disposition not in ("trash", "permanent"):
        raise ToolError("validation", "disposition must be 'trash' or 'permanent'")
    try:
        raw_id = ctx.aliaser.resolve(contact_ref)
    except KeyError as exc:
        raise ToolError("validation", str(exc.args[0] if exc.args else exc)) from exc
    contact = _fetch_contact(account, raw_id)
    return {
        "target": contact, "raw_id": raw_id,
        "contact": contact_details(contact), "disposition": disposition,
    }


async def _delete_contact_preview(ctx: Context,
                                  kwargs: dict[str, Any]) -> dict[str, Any]:
    contact_ref = str(kwargs.get("id") or "")
    disposition = str(kwargs.get("disposition") or "trash")

    def work(account: Any) -> dict[str, Any]:
        snap = _delete_snapshot(ctx, account, contact_ref, disposition)
        details = snap["contact"]
        return {
            "action": "delete_contact",
            "id": ctx.aliaser.alias_for(snap["raw_id"], "p"),
            "name": details.get("name"), "email": details.get("email"),
            "disposition": disposition, "recoverable": disposition == "trash",
            "_confirm_binding": {
                "id": snap["raw_id"],
                "changekey": getattr(snap["target"], "changekey", None),
                "name": details.get("name"), "email": details.get("email"),
                "disposition": disposition,
            },
        }

    return await ctx.gateway.call(work)


async def _delete_contact(ctx: Context, *, id: str,
                          disposition: str = "trash") -> dict[str, Any]:
    def work(account: Any) -> dict[str, Any]:
        snap = _delete_snapshot(ctx, account, id, disposition)
        if disposition == "trash":
            snap["target"].move_to_trash()
        else:
            snap["target"].delete()
        return snap

    deleted = await _gateway_write(ctx, work)
    return {
        "ok": True,
        "deleted": {
            "id": ctx.aliaser.alias_for(deleted["raw_id"], "p"),
            "name": deleted["contact"].get("name"),
            "email": deleted["contact"].get("email"),
            "disposition": disposition,
            "recoverable": disposition == "trash",
        },
    }


def _obj(props: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    return {"type": "object", "properties": props,
            "required": required or [], "additionalProperties": False}


def _string(description: str) -> dict[str, Any]:
    return {"type": "string", "minLength": 1, "maxLength": MAX_TEXT,
            "description": description}


_ADDRESS = _obj({
    "street": _string("Street and house number."),
    "city": _string("City."),
    "state": _string("State, province or region."),
    "postal_code": _string("Postal or ZIP code."),
    "country": _string("Country."),
})
_CATEGORIES = {
    "type": "array", "items": _string("One Outlook category."),
    "maxItems": MAX_CATEGORIES,
    "description": "Replace the contact's Outlook categories; [] clears them.",
}
_CONTACT_FIELDS: dict[str, Any] = {
    "email": {**_string("Primary SMTP email address (EmailAddress1)."),
              "format": "email"},
    "given_name": _string("First or given name."),
    "surname": _string("Last or family name."),
    "display_name": _string("Outlook display name; otherwise derived from names."),
    "business_phone": _string("Business phone number."),
    "mobile_phone": _string("Mobile phone number."),
    "company": _string("Company or organization."),
    "job_title": _string("Job title or role."),
    "department": _string("Department."),
    "office": _string("Office or location label."),
    "business_homepage": _string("Business website URL."),
    "business_address": {**_ADDRESS, "description": "Complete business address."},
    "categories": dict(_CATEGORIES),
}
_CONTACT_ID = {
    "type": "string", "minLength": 1,
    "description": "Personal-contact p-alias from find_people(source='contacts'), or raw id.",
}


TOOLS: list[ToolSpec] = [
    ToolSpec(
        name="create_contact",
        description=(
            "Create a personal Outlook contact. An exact primary-email match is "
            "returned with created=false unless allow_duplicate=true; an optional "
            "idempotency_key makes retries durable across processes."
        ),
        side_effect_class="write", requires_ews=True,
        input_schema=_obj({
            **{key: dict(value) for key, value in _CONTACT_FIELDS.items()},
            "allow_duplicate": {
                "type": "boolean", "default": False,
                "description": "Allow another personal contact with the same email.",
            },
            "idempotency_key": {
                "type": "string", "minLength": 1, "maxLength": 200,
                "description": "Stable key for safe retries of this create.",
            },
        }, ["email"]),
        handler=_create_contact,
    ),
    ToolSpec(
        name="update_contact",
        description=(
            "Update one personal Outlook contact by p-alias. GAL entries are "
            "read-only. Omitted fields stay unchanged; categories=[] clears "
            "categories and clear_fields removes selected optional values."
        ),
        side_effect_class="write", requires_ews=True,
        input_schema=_obj({
            "id": dict(_CONTACT_ID),
            **{key: dict(value) for key, value in _CONTACT_FIELDS.items()},
            "clear_fields": {
                "type": "array", "uniqueItems": True,
                "items": {"type": "string", "enum": list(_CLEARABLE_FIELDS)},
                "description": "Optional fields to remove instead of setting.",
            },
            "allow_duplicate": {
                "type": "boolean", "default": False,
                "description": "Allow changing the email to one already in Contacts.",
            },
        }, ["id"]),
        handler=_update_contact,
    ),
    ToolSpec(
        name="delete_contact",
        description=(
            "Delete one personal Outlook contact. Default disposition=trash is "
            "recoverable; permanent uses HardDelete. GAL entries are protected. "
            "Always uses a content-bound two-phase confirmation."
        ),
        side_effect_class="destructive", requires_ews=True, confirm=True,
        input_schema=_obj({
            "id": dict(_CONTACT_ID),
            "disposition": {
                "type": "string", "enum": ["trash", "permanent"],
                "default": "trash",
                "description": "Recoverable trash move (default) or HardDelete.",
            },
        }, ["id"]),
        handler=_delete_contact, preview=_delete_contact_preview,
    ),
]
