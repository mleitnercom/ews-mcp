"""stdio -> Streamable-HTTP-Bruecke zum gemeinsamen Server (run_http_server.py).

Claude Desktop und Codex starten MCP-Server per stdio. Statt dort jeweils
einen eigenen Exchange-Server zu starten, startet ihr Config-Eintrag diese
Bruecke: Sie leitet die MCP-Nachrichten unveraendert an
http://127.0.0.1:MCP_PORT/mcp weiter und hat selbst keine Exchange-Verbindung.

Ersetzt `npx mcp-remote` (auf dieser Maschine ist kein Node installiert).
Key und Port kommen aus .env (MCP_API_KEY, MCP_PORT).

Reconnect (seit 23.10.2026): Der Supervisor in run_http_server.py startet den
Server neu, wenn sein Listener stirbt. Der neue Prozess kennt die alte
MCP-Session nicht. Die erste Fassung dieser Bruecke nutzte den SDK-Transport,
der danach nur noch "Session terminated" lieferte bzw. bei abgewiesener
Verbindung die ganze Bruecke beendete -- der Konnektor fehlte in der App bis
zu deren Neustart. Deshalb jetzt eine eigene, schlanke Weiterleitung:

- Sie merkt sich das initialize des Clients. Antwortet der Server mit 404
  (Session unbekannt) oder ist er nicht erreichbar, wartet sie bis zu
  RECONNECT_WAIT_S auf ihn, spielt initialize + notifications/initialized
  erneut ein und schickt die Nachricht noch einmal.
- Wiederholt wird nur, was den Server sicher nicht erreicht hat (404 bzw.
  Verbindungsaufbau gescheitert). Bricht die Verbindung WAEHREND eines
  Aufrufs ab, bekommt der Client einen Fehler "Ergebnis unbekannt" -- ein
  create_draft darf nicht doppelt laufen.

stdout gehoert dem MCP-Protokoll; Logs gehen nach stderr und in
DATA_DIR/logs/bridge-<datum>-<pid>.log.
"""
import json
import logging
import os
import sys
import time
from datetime import datetime

import anyio
import httpx
from httpx_sse import EventSource
from mcp.server.stdio import stdio_server
from mcp.shared.message import SessionMessage
from mcp.types import JSONRPCMessage

HERE = os.path.dirname(os.path.abspath(__file__))

RECONNECT_WAIT_S = 180
SESSION_HEADER = "mcp-session-id"
PROTOCOL_HEADER = "mcp-protocol-version"


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
                         f"bridge-{datetime.now().astimezone():%Y%m%d}-"
                         f"{os.getpid()}.log"),
            encoding="utf-8")
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except Exception as exc:  # noqa: BLE001 — Bruecke muss trotzdem laufen
        logger.warning("file logging unavailable (%s)", exc)


class SessionLost(Exception):
    """The server answered 404: it does not know our session (restarted)."""


# Failures where the message demonstrably never reached the server.
NOT_DELIVERED = (SessionLost, httpx.ConnectError, httpx.ConnectTimeout)


def _is_request(obj: dict) -> bool:
    return "method" in obj and "id" in obj


def _is_response(obj: dict) -> bool:
    return "method" not in obj and "id" in obj


class Upstream:
    def __init__(self, url: str, http: httpx.AsyncClient, to_client) -> None:
        self.url = url
        self.http = http
        self.to_client = to_client
        self.session_id: str | None = None
        self.protocol: str | None = None
        self.init_msg: dict | None = None
        self.initialized_msg: dict | None = None
        self.generation = 0
        self.lock = anyio.Lock()

    def _headers(self) -> dict:
        h = {"Accept": "application/json, text/event-stream",
             "Content-Type": "application/json"}
        if self.session_id:
            h[SESSION_HEADER] = self.session_id
        if self.protocol:
            h[PROTOCOL_HEADER] = self.protocol
        return h

    async def _deliver(self, obj: dict) -> None:
        await self.to_client.send(
            SessionMessage(JSONRPCMessage.model_validate(obj)))

    async def _post(self, obj: dict, sink) -> None:
        """POST one message; hand every message the server sends to sink."""
        async with self.http.stream("POST", self.url, json=obj,
                                    headers=self._headers()) as resp:
            if resp.status_code == 404:
                raise SessionLost()
            if resp.status_code == 202:
                return
            resp.raise_for_status()
            if obj.get("method") == "initialize":
                self.session_id = resp.headers.get(SESSION_HEADER)
            ctype = resp.headers.get("content-type", "").lower()
            if ctype.startswith("text/event-stream"):
                async for sse in EventSource(resp).aiter_sse():
                    if not sse.data:
                        continue
                    msg = sse.json()
                    await sink(msg)
                    if _is_response(msg) and msg.get("id") == obj.get("id"):
                        return
                if _is_request(obj):
                    raise httpx.ReadError("event stream ended without a "
                                          "response")
            elif ctype.startswith("application/json"):
                await sink(json.loads(await resp.aread()))

    def _note_protocol(self, msg: dict) -> None:
        version = (msg.get("result") or {}).get("protocolVersion")
        if version:
            self.protocol = version

    async def _wait_for(self, attempt, what: str) -> None:
        """Run attempt() until the server takes it or RECONNECT_WAIT_S ends."""
        deadline = time.monotonic() + RECONNECT_WAIT_S
        delay = 1.0
        while True:
            try:
                return await attempt()
            except NOT_DELIVERED as exc:
                if time.monotonic() > deadline:
                    raise
                logger.info("%s: server not reachable (%s), retry in %.0fs",
                            what, type(exc).__name__, delay)
                await anyio.sleep(delay)
                delay = min(delay * 2, 10)

    async def _handshake(self) -> None:
        self.session_id = None
        replay = dict(self.init_msg,
                      id=f"bridge-reinit-{self.generation + 1}")

        async def swallow(msg: dict) -> None:
            if _is_response(msg) and msg.get("id") == replay["id"]:
                if "error" in msg:
                    raise RuntimeError(f"re-initialize refused: {msg['error']}")
                self._note_protocol(msg)
            else:
                await self._deliver(msg)

        await self._post(replay, swallow)
        if self.initialized_msg:
            await self._post(self.initialized_msg, self._deliver)

    async def reconnect(self, failed_generation: int) -> None:
        async with self.lock:
            if self.generation != failed_generation:
                return  # a concurrent call already reconnected
            if self.init_msg is None:
                raise SessionLost()
            logger.warning("server gone or session unknown -- re-initializing")
            await self._wait_for(self._handshake, "reconnect")
            self.generation += 1
            logger.info("reconnected (generation %d)", self.generation)

    async def _error(self, obj: dict, message: str) -> None:
        if _is_request(obj):
            await self._deliver({"jsonrpc": "2.0", "id": obj["id"],
                                 "error": {"code": -32000,
                                           "message": message}})

    async def forward(self, obj: dict) -> None:
        method = obj.get("method")

        if method == "initialize":
            self.init_msg = obj

            async def sink(msg: dict) -> None:
                if _is_response(msg) and msg.get("id") == obj.get("id"):
                    self._note_protocol(msg)
                await self._deliver(msg)

            try:
                await self._wait_for(lambda: self._post(obj, sink),
                                     "initialize")
            except Exception as exc:  # noqa: BLE001
                logger.error("initialize failed: %r", exc)
                await self._error(obj, f"Exchange-Server nicht erreichbar: "
                                       f"{exc!r}")
            return
        if method == "notifications/initialized":
            self.initialized_msg = obj

        for attempt in (1, 2):
            generation = self.generation
            try:
                await self._post(obj, self._deliver)
                return
            except NOT_DELIVERED as exc:
                if attempt == 2:
                    logger.error("%s not delivered after reconnect: %r",
                                 method or "response", exc)
                    break
                try:
                    await self.reconnect(generation)
                except Exception as exc2:  # noqa: BLE001
                    logger.error("reconnect failed: %r", exc2)
                    break
            except httpx.HTTPStatusError as exc:
                logger.warning("%s: HTTP %s", method, exc.response.status_code)
                await self._error(obj, f"Exchange-Server antwortet mit HTTP "
                                       f"{exc.response.status_code}")
                return
            except (httpx.TransportError, ValueError) as exc:
                # Reached the server, then the line broke: it may have run.
                logger.warning("%s: connection lost mid-call (%r)",
                               method, exc)
                await self._error(obj, "Verbindung zum Exchange-Server waehrend "
                                       "des Aufrufs abgebrochen -- Ergebnis "
                                       "unbekannt, bitte pruefen statt blind "
                                       "wiederholen")
                return
        await self._error(obj, f"Exchange-Server seit {RECONNECT_WAIT_S}s "
                               f"nicht erreichbar")

    async def close(self) -> None:
        if not self.session_id:
            return
        try:
            await self.http.delete(self.url, headers=self._headers(),
                                   timeout=3)
        except Exception as exc:  # noqa: BLE001 — beim Beenden egal
            logger.debug("session delete on close failed: %r", exc)


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
        timeout=httpx.Timeout(60, connect=10, read=900),
        follow_redirects=True,
    )
    async with http, stdio_server() as (stdin, stdout):
        upstream = Upstream(url, http, stdout)
        async with anyio.create_task_group() as tg:
            async for msg in stdin:
                if isinstance(msg, Exception):
                    logger.warning("client: %s", msg)
                    continue
                obj = msg.message.model_dump(by_alias=True, mode="json",
                                             exclude_none=True)
                # Requests concurrently (tool calls can take minutes),
                # everything else in order.
                if _is_request(obj) and obj.get("method") != "initialize":
                    tg.start_soon(upstream.forward, obj)
                else:
                    await upstream.forward(obj)
            # Client hat stdin geschlossen -> Bruecke beenden.
            tg.cancel_scope.cancel()
        await upstream.close()
        # stdio_server waits for its stdout writer, which ends only when
        # this stream is closed -- without it the bridge never exits.
        await stdout.aclose()
    logger.info("bridge closed")


if __name__ == "__main__":
    _setup_logging()
    anyio.run(main)
