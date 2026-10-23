"""stdio -> Streamable-HTTP-Bruecke zum gemeinsamen Server (run_http_server.py).

Claude Desktop und Codex starten MCP-Server per stdio. Statt dort jeweils
einen eigenen Exchange-Server zu starten, startet ihr Config-Eintrag diese
Bruecke: Sie leitet die MCP-Nachrichten unveraendert an
http://127.0.0.1:MCP_PORT/mcp weiter und hat selbst keine Exchange-Verbindung.

Ersetzt `npx mcp-remote` (auf dieser Maschine ist kein Node installiert).
Key und Port kommen aus .env (MCP_API_KEY, MCP_PORT).

stdout gehoert dem MCP-Protokoll; Logs gehen nach stderr und in
DATA_DIR/logs/bridge-<datum>-<pid>.log.
"""
import logging
import os
import sys
from datetime import datetime

import anyio
import httpx
from mcp.client.streamable_http import streamable_http_client
from mcp.server.stdio import stdio_server

HERE = os.path.dirname(os.path.abspath(__file__))


def _env_from_dotenv(key: str) -> str:
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


def _setting(key: str, default: str = "") -> str:
    return os.environ.get(key) or _env_from_dotenv(key) or default


logger = logging.getLogger("http_bridge")


def _setup_logging() -> None:
    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s [pid=%(process)d] %(name)s: %(message)s")
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    stream = logging.StreamHandler(sys.stderr)
    stream.setFormatter(fmt)
    root.addHandler(stream)
    # Eine Zeile pro HTTP-Request waere nur Rauschen.
    logging.getLogger("httpx").setLevel(logging.WARNING)
    try:
        log_dir = os.path.join(
            _setting("DATA_DIR", os.path.join(HERE, "data")), "logs")
        os.makedirs(log_dir, exist_ok=True)
        fh = logging.FileHandler(
            os.path.join(log_dir,
                         f"bridge-{datetime.now():%Y%m%d}-{os.getpid()}.log"),
            encoding="utf-8")
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except Exception as exc:  # noqa: BLE001 — Bruecke muss trotzdem laufen
        logger.warning("file logging unavailable (%s)", exc)


async def _pump(src, dst, label: str) -> None:
    async for msg in src:
        if isinstance(msg, Exception):
            logger.warning("%s: %s", label, msg)
            continue
        await dst.send(msg)


async def main() -> None:
    url = f"http://127.0.0.1:{_setting('MCP_PORT', '8765')}/mcp"
    key = _setting("MCP_API_KEY")
    if not key:
        sys.exit("MCP_API_KEY fehlt in .env")
    logger.info("bridge -> %s", url)

    # Lange Lese-Timeouts: Tool-Aufrufe koennen waehrend eines Exchange-
    # Cooldowns mehrere Minuten auf die Antwort warten.
    http = httpx.AsyncClient(
        headers={"Authorization": f"Bearer {key}"},
        timeout=httpx.Timeout(60, read=900),
        follow_redirects=True,
    )
    async with http, stdio_server() as (stdin, stdout):
        async with streamable_http_client(url, http_client=http) as (
                http_in, http_out, _):
            async with anyio.create_task_group() as tg:
                async def client_to_server():
                    await _pump(stdin, http_out, "client")
                    # Client hat stdin geschlossen -> Bruecke beenden.
                    tg.cancel_scope.cancel()

                tg.start_soon(client_to_server)
                tg.start_soon(_pump, http_in, stdout, "server")
    logger.info("bridge closed")


if __name__ == "__main__":
    _setup_logging()
    anyio.run(main)
