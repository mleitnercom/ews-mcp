"""Tests for contact tools."""

import pytest
from unittest.mock import MagicMock, patch

from src.tools.contact_tools import (
    CreateContactTool,
    UpdateContactTool,
    DeleteContactTool
)
from src.tools.contact_intelligence_tools import FindPersonTool


@pytest.mark.asyncio
async def test_find_person_gal(mock_ews_client):
    """Test finding person via GAL (replaces resolve_names)."""
    tool = FindPersonTool(mock_ews_client)

    # Mock resolution result
    mock_resolution = MagicMock()
    mock_mailbox = MagicMock()
    mock_mailbox.name = "John Doe"
    mock_mailbox.email_address = "john.doe@example.com"
    mock_mailbox.routing_type = "SMTP"
    mock_mailbox.mailbox_type = "Mailbox"
    mock_resolution.mailbox = mock_mailbox

    mock_ews_client.account.protocol.resolve_names.return_value = [mock_resolution]

    result = await tool.execute(
        query="john",
        source="gal"
    )

    assert result["success"] is True


@pytest.mark.asyncio
async def test_find_person_no_results(mock_ews_client):
    """Test finding person with no matches."""
    tool = FindPersonTool(mock_ews_client)

    mock_ews_client.account.protocol.resolve_names.return_value = []

    result = await tool.execute(query="nonexistent", source="gal")

    assert result["success"] is True


@pytest.mark.asyncio
async def test_create_contact_tool_sets_categories(mock_ews_client):
    """Create contact should persist provided categories."""
    tool = CreateContactTool(mock_ews_client)

    with patch('src.tools.contact_tools.Contact') as mock_contact_cls:
        mock_contact = MagicMock()
        mock_contact.id = "contact-1"
        mock_contact_cls.return_value = mock_contact

        result = await tool.execute(
            given_name="Michael",
            surname="Leitner",
            email_address="michael@example.com",
            categories=["VIP", "Blocker"]
        )

    assert result["success"] is True
    assert mock_contact.categories == ["VIP", "Blocker"]


@pytest.mark.asyncio
async def test_update_contact_tool_sets_categories(mock_ews_client):
    """Update contact should replace categories when provided."""
    tool = UpdateContactTool(mock_ews_client)

    mock_contact = MagicMock()
    mock_ews_client.account.contacts.get.return_value = mock_contact

    result = await tool.execute(
        item_id="contact-1",
        categories=["VIP", "Blocker"]
    )

    assert result["success"] is True
    assert mock_contact.categories == ["VIP", "Blocker"]
