"""Verbindungs-Check gegen PBS Exchange - liest .env, schreibt nichts.

Aufruf aus v5:  .venv\Scripts\python.exe ews_check.py
"""
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import truststore

truststore.inject_into_ssl()  # interne PBS-PKI, siehe run_server.py

from ewsmcp.config import Settings

s = Settings()
if not s.ews_password:
    sys.exit("FEHLT: EWS_PASSWORD ist in .env noch leer.")

from exchangelib import DELEGATE, Account, Configuration, Credentials

cfg = Configuration(
    service_endpoint=s.ews_server_url,
    credentials=Credentials(s.ews_username or s.ews_email, s.ews_password),
)
acct = Account(
    primary_smtp_address=s.ews_email,
    config=cfg,
    autodiscover=False,
    access_type=DELEGATE,
)

print("OK  Verbindung :", acct.primary_smtp_address)
print("OK  Exchange   :", acct.protocol.version)
print("OK  Auth-Typ   :", acct.protocol.auth_type)
print("OK  Posteingang:", acct.inbox.total_count, "Elemente")

tz = ZoneInfo(s.ews_tz)
start = datetime.now(tz).replace(hour=0, minute=0, second=0, microsecond=0)
items = list(acct.calendar.view(start=start, end=start + timedelta(days=1)))
allday = sum(1 for i in items if getattr(i, "is_all_day", False))
print(f"OK  Kalender heute: {len(items)} Termine (davon ganztags: {allday})")
