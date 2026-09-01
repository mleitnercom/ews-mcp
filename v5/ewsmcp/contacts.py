"""Shared personal-contact DTO and indexed-field helpers.

The read and write packs use the same compact representation so a contact
returned by ``find_people`` / ``get_contact`` can be passed directly to the
contact-management tools without field-name translation.
"""

from __future__ import annotations

from typing import Any


def indexed_label(value: Any) -> str:
    return str(getattr(value, "label", "") or "")


def replace_indexed(values: Any, label: str, replacement: Any | None) -> list[Any]:
    """Replace one labelled EWS indexed property while preserving the rest."""
    kept = [value for value in (values or []) if indexed_label(value) != label]
    if replacement is not None:
        kept.append(replacement)
    return kept


def contact_details(contact: Any) -> dict[str, Any]:
    """Return an alias-free, raw-id-free personal-contact DTO."""
    email_addresses = [
        {"label": indexed_label(entry), "email": str(email)}
        for entry in (getattr(contact, "email_addresses", None) or [])
        if (email := getattr(entry, "email", None))
    ]
    primary = next(
        (entry["email"] for entry in email_addresses
         if entry["label"] == "EmailAddress1"),
        email_addresses[0]["email"] if email_addresses else None,
    )
    display = str(getattr(contact, "display_name", None) or primary or "")
    result: dict[str, Any] = {
        "name": display,
        "email": primary,
        "source": "contacts",
    }

    simple_fields = {
        "given_name": "given_name",
        "surname": "surname",
        "company": "company_name",
        "title": "job_title",
        "department": "department",
        "office": "office",
        "business_homepage": "business_homepage",
    }
    for public, attr in simple_fields.items():
        value = getattr(contact, attr, None)
        if value:
            result[public] = str(value)

    if email_addresses:
        result["email_addresses"] = email_addresses

    phone_numbers = [
        {"label": indexed_label(entry), "number": str(number)}
        for entry in (getattr(contact, "phone_numbers", None) or [])
        if (number := getattr(entry, "phone_number", None))
    ]
    if phone_numbers:
        result["phone_numbers"] = phone_numbers
        result["phone"] = next(
            (entry["number"] for entry in phone_numbers
             if entry["label"] == "BusinessPhone"),
            phone_numbers[0]["number"],
        )

    addresses = []
    for entry in (getattr(contact, "physical_addresses", None) or []):
        address = {
            "label": indexed_label(entry),
            "street": getattr(entry, "street", None),
            "city": getattr(entry, "city", None),
            "state": getattr(entry, "state", None),
            "postal_code": getattr(entry, "zipcode", None),
            "country": getattr(entry, "country", None),
        }
        address = {key: str(value) for key, value in address.items() if value}
        if len(address) > 1:
            addresses.append(address)
    if addresses:
        result["physical_addresses"] = addresses
        business = next(
            (entry for entry in addresses if entry.get("label") == "Business"), None)
        if business:
            result["business_address"] = {
                key: value for key, value in business.items() if key != "label"
            }

    categories = [str(value) for value in (getattr(contact, "categories", None) or [])]
    if categories:
        result["categories"] = categories
    return result


def enrich_from_contact(person: dict[str, Any], contact: Any,
                        *, detailed: bool = False) -> None:
    """Enrich a GAL/person card from an exchangelib Contact object."""
    if contact is None:
        return
    details = contact_details(contact)
    for key in ("title", "company", "phone"):
        if details.get(key):
            person[key] = details[key]
    if detailed:
        for key, value in details.items():
            if key not in ("name", "email", "source") and value is not None:
                person[key] = value


def personal_contact(contact: Any, *, detailed: bool = False
                     ) -> tuple[str, dict[str, Any]] | None:
    """Build the raw alias key and public DTO for a personal contact."""
    details = contact_details(contact)
    email = details.get("email")
    if not email:
        return None
    raw_key = str(getattr(contact, "id", None) or email)
    if not detailed:
        compact = {
            key: value for key, value in details.items()
            if key in ("name", "email", "source", "title", "company", "phone")
        }
        return raw_key, compact
    return raw_key, details
