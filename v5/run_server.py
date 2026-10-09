"""Wrapper fuer Claude Desktop (MSIX): cwd, Windows-Trust-Store, Datei-Log.

Drei Gruende:

1. Die Store-Variante von Claude Desktop reicht das cwd-Feld aus
   claude_desktop_config.json nicht korrekt an den Kindprozess weiter.
   ews-mcp laedt .env relativ zum Arbeitsverzeichnis -> ohne chdir keine
   Credentials. Gleiches Muster wie beim v3-Setup (Vault: PBS-Exchange-MCP-Setup).

2. owa.pbs-holding.com traegt ein Zertifikat der internen PBS-PKI
   (PBSRootCA2020 -> PBS-IssuingCA1). Windows vertraut dieser Kette, Pythons
   certifi-Bundle nicht -> SSLCertVerificationError. truststore legt den
   Windows-Zertifikatspeicher unter ssl/requests und macht die Verifikation
   damit korrekt, statt sie abzuschalten (EWS_INSECURE_SKIP_VERIFY).

3. Im stdio-Betrieb geht stderr an claude.exe und wird dort NICHT persistiert
   (beide Log-Verzeichnisse der App bleiben leer). Nach dem Vorfall vom
   07.10.2026 liess sich deshalb nicht klaeren, was waehrend der haengenden
   Aufrufe im Server passierte. Dieser Wrapper schreibt deshalb zusaetzlich
   in eine Datei unter DATA_DIR/logs. Der stderr-Kanal bleibt erhalten:
   main() ruft logging.basicConfig(), das wirkungslos ist, sobald der
   Root-Logger bereits Handler hat -- beide werden hier gesetzt.
"""
import logging
import logging.handlers
import os
import sys
from datetime import datetime

import truststore

truststore.inject_into_ssl()

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)


def _env_from_dotenv(key: str) -> str:
    """Read one key from .env without importing Settings (which needs the
    credentials to validate and must not run before logging is up)."""
    try:
        with open(os.path.join(HERE, ".env"), encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    name, _, value = line.partition("=")
                    if name.strip() == key:
                        return value.strip()
    except OSError:
        pass
    return ""


def _setup_logging() -> None:
    """Never let logging setup stop the server from booting."""
    level_name = (os.environ.get("LOG_LEVEL") or _env_from_dotenv("LOG_LEVEL")
                  or "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    root = logging.getLogger()
    root.setLevel(level)
    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s [pid=%(process)d] %(name)s: %(message)s")

    stream = logging.StreamHandler(sys.stderr)  # stdout belongs to MCP
    stream.setFormatter(fmt)
    root.addHandler(stream)

    # exchangelib logs one INFO line PER FIELD for naive datetimes and
    # unknown timezone ids. On a normal mailbox that is thousands of lines
    # an hour and it buried the useful entries on 2026-10-09. Warnings from
    # the same logger still come through.
    logging.getLogger("exchangelib.fields").setLevel(logging.WARNING)

    data_dir = (os.environ.get("DATA_DIR") or _env_from_dotenv("DATA_DIR")
                or os.path.join(HERE, "data"))
    try:
        log_dir = os.path.join(data_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)
        # One file per process: two instances (desktop chat + claude-device)
        # run concurrently and would otherwise interleave into one file.
        path = os.path.join(
            log_dir,
            f"ews-mcp-{datetime.now():%Y%m%d}-{os.getpid()}.log")
        rotating = logging.handlers.RotatingFileHandler(
            path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
        rotating.setFormatter(fmt)
        root.addHandler(rotating)
        root.info("file logging active: %s", path)
    except Exception as exc:  # noqa: BLE001 — boot must survive a bad path
        root.warning("file logging unavailable (%s) — stderr only", exc)


_setup_logging()

from ewsmcp.main import main  # noqa: E402 — logging must exist first

if __name__ == "__main__":
    main()
