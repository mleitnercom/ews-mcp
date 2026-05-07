from datetime import datetime
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from src.tools.calendar_tools import GetCalendarTool


@pytest.mark.asyncio
async def test_get_calendar_returns_categories():
    mock_client = MagicMock()
    account = MagicMock()
    mock_client.get_account.return_value = account
    mock_client.get_mailbox_info.return_value = "test@example.com"

    organizer = SimpleNamespace(email_address="organizer@example.com")
    attendee = SimpleNamespace(mailbox=SimpleNamespace(email_address="attendee@example.com"))
    event = SimpleNamespace(
        id="event-id",
        subject="Calendar event",
        start=datetime(2026, 5, 7, 9, 0, 0),
        end=datetime(2026, 5, 7, 10, 0, 0),
        location="Room 1",
        organizer=organizer,
        is_all_day=False,
        required_attendees=[attendee],
        categories=["Blocker", "CodexTest"],
    )

    query = MagicMock()
    query.only.return_value = query
    query.order_by.return_value = [event]
    account.calendar.view.return_value = query

    tool = GetCalendarTool(mock_client)
    result = await tool.execute(
        start_date="2026-05-07T00:00:00+02:00",
        end_date="2026-05-08T00:00:00+02:00",
        max_results=10,
    )

    assert result["success"] is True
    assert result["events"][0]["categories"] == ["Blocker", "CodexTest"]
