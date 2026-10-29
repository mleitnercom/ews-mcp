"""Gemeinsamer HTTP-Server fuer alle Clients (Claude Desktop, Claude Code, Codex).

Statt dass jeder Client per stdio seine eigene Instanz startet (drei
Exchange-Verbindungen, die sich das Throttling-Budget teilen), laeuft genau
eine Instanz dauerhaft auf 127.0.0.1:MCP_PORT. Die Clients verbinden sich
ueber http_bridge.py damit.

Gestartet von der geplanten Aufgabe "ews-mcp HTTP" bei der Anmeldung, mit
pythonw.exe (kein Konsolenfenster).

Supervisor (seit 23.10.2026): Am 23.10.2026 um 08:26 starb bei einem
Netzwerkwechsel der Accept-Loop des Listeners (WinError 64). Der Prozess lief
weiter, Port 8765 hoerte aber nicht mehr -- alle Clients hingen, die geplante
Aufgabe stand trotzdem auf "Running". Dieses Skript startet den Server deshalb
als Kindprozess und fragt /livez ab. Antwortet der Listener FAILS_BEFORE_RESTART
Mal hintereinander nicht, wird der Kindprozess beendet und neu gestartet;
ebenso, wenn er sich selbst beendet. /livez haengt nicht an Exchange, ein
nicht erreichbares Exchange loest also keinen Neustart aus.

Ein zweiter Start findet einen antwortenden Port vor und beendet sich -- es
bleibt bei einer Instanz. Der Kindprozess haengt an einem Job-Objekt mit
KILL_ON_JOB_CLOSE: Wird die Aufgabe beendet, stirbt der Server mit.
Log: DATA_DIR/logs/supervisor.log.

Transport und Host werden hier gesetzt, nicht in .env: .env gilt auch fuer
den stdio-Betrieb, und dort darf MCP_TRANSPORT=http nicht stehen.
"""
import ctypes
import ctypes.wintypes as wt
import logging
import logging.handlers
import os
import runpy
import subprocess
import sys
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))

CHILD_FLAG = "EWSMCP_SUPERVISED_CHILD"
PROBE_INTERVAL_S = 15
PROBE_TIMEOUT_S = 5
FAILS_BEFORE_RESTART = 3
BOOT_GRACE_S = 90
CREATE_NO_WINDOW = 0x08000000


def _run_server() -> None:
    os.environ["MCP_TRANSPORT"] = "http"
    os.environ["MCP_HOST"] = "127.0.0.1"

    # pythonw.exe hat kein stdout/stderr; der StreamHandler in run_server.py
    # braucht aber einen Strom. Das Datei-Log unter DATA_DIR/logs bleibt.
    # Kein with: die Stroeme muessen bis zum Prozessende offen bleiben.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")  # noqa: SIM115
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")  # noqa: SIM115

    runpy.run_path(os.path.join(HERE, "run_server.py"), run_name="__main__")


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


def _logger() -> logging.Logger:
    log = logging.getLogger("supervisor")
    log.setLevel(logging.INFO)
    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s [pid=%(process)d] %(name)s: %(message)s")
    try:
        log_dir = os.path.join(
            _setting("DATA_DIR", os.path.join(HERE, "data")), "logs")
        os.makedirs(log_dir, exist_ok=True)
        handler = logging.handlers.RotatingFileHandler(
            os.path.join(log_dir, "supervisor.log"),
            maxBytes=1024 * 1024, backupCount=3, encoding="utf-8")
    except Exception:  # noqa: BLE001 -- der Server muss trotzdem starten
        handler = logging.NullHandler()
    handler.setFormatter(fmt)
    log.addHandler(handler)
    return log


def _livez(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/livez",
                                    timeout=PROBE_TIMEOUT_S) as resp:
            return resp.status == 200
    except Exception:  # noqa: BLE001 -- jede Art Fehler heisst "nicht da"
        return False


class _IoCounters(ctypes.Structure):
    _fields_ = [(n, ctypes.c_ulonglong) for n in (
        "ReadOperationCount", "WriteOperationCount", "OtherOperationCount",
        "ReadTransferCount", "WriteTransferCount", "OtherTransferCount")]


class _BasicLimits(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", wt.LARGE_INTEGER),
        ("PerJobUserTimeLimit", wt.LARGE_INTEGER),
        ("LimitFlags", wt.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wt.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wt.DWORD),
        ("SchedulingClass", wt.DWORD),
    ]


class _ExtendedLimits(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", _BasicLimits),
        ("IoInfo", _IoCounters),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryUsed", ctypes.c_size_t),
        ("PeakJobMemoryUsed", ctypes.c_size_t),
    ]


def _kill_children_with_me(log: logging.Logger):
    """Put this process (and so every child) into a KILL_ON_JOB_CLOSE job.

    The handle is only held here; when the supervisor dies for whatever
    reason, Windows closes it and takes the server down too. Without this a
    stopped task would leave an orphaned server holding the port."""
    try:
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateJobObjectW.restype = wt.HANDLE
        k32.GetCurrentProcess.restype = wt.HANDLE
        job = k32.CreateJobObjectW(None, None)
        info = _ExtendedLimits()
        info.BasicLimitInformation.LimitFlags = 0x2000  # KILL_ON_JOB_CLOSE
        ok = (job
              and k32.SetInformationJobObject(
                  wt.HANDLE(job), 9, ctypes.byref(info), ctypes.sizeof(info))
              and k32.AssignProcessToJobObject(
                  wt.HANDLE(job), wt.HANDLE(k32.GetCurrentProcess())))
        if not ok:
            log.warning("job object unavailable (winerror %s) -- a stopped "
                        "task may leave the server running",
                        ctypes.get_last_error())
        return job
    except Exception as exc:  # noqa: BLE001
        log.warning("job object unavailable (%s)", exc)
        return None


def _kill_tree(child: subprocess.Popen) -> None:
    # The venv's pythonw.exe is a launcher with the real interpreter as its
    # own child, so the whole tree has to go.
    subprocess.run(["taskkill", "/PID", str(child.pid), "/T", "/F"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   creationflags=CREATE_NO_WINDOW, check=False)
    try:
        child.wait(timeout=15)
    except subprocess.TimeoutExpired:
        child.kill()


def _watch(child: subprocess.Popen, port: int, log: logging.Logger) -> str:
    """Block until the server is gone or deaf; return why."""
    up = False
    fails = 0
    boot_deadline = time.monotonic() + BOOT_GRACE_S
    while True:
        time.sleep(PROBE_INTERVAL_S if up else 2)
        rc = child.poll()
        if rc is not None:
            return f"exited rc={rc}"
        if _livez(port):
            if not up:
                log.info("listener up on 127.0.0.1:%s", port)
            elif fails:
                log.info("listener answers again after %d failed probe(s)",
                         fails)
            up, fails = True, 0
            continue
        if not up:
            if time.monotonic() > boot_deadline:
                return f"listener not up within {BOOT_GRACE_S}s"
            continue
        fails += 1
        log.warning("livez failed (%d/%d)", fails, FAILS_BEFORE_RESTART)
        if fails >= FAILS_BEFORE_RESTART:
            return "listener dead"


def _supervise() -> None:
    log = _logger()
    port = int(_setting("MCP_PORT", "8765"))
    if _livez(port):
        log.info("127.0.0.1:%s answers already -- another instance runs, "
                 "exiting", port)
        return
    job = _kill_children_with_me(log)  # noqa: F841 -- handle must stay open
    log.info("supervisor started (port %s)", port)

    backoff = 5
    while True:
        started = time.monotonic()
        child = subprocess.Popen(
            [sys.executable, os.path.abspath(__file__)],
            env={**os.environ, CHILD_FLAG: "1"}, cwd=HERE,
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL, creationflags=CREATE_NO_WINDOW)
        log.info("server started (pid=%s)", child.pid)

        reason = _watch(child, port, log)
        if child.poll() is None:
            _kill_tree(child)

        # A server that dies right after boot (bad .env, port taken by a
        # foreign process) must not be restarted in a tight loop.
        if time.monotonic() - started > 300:
            backoff = 5
        log.warning("server down (%s) -- restart in %ss", reason, backoff)
        time.sleep(backoff)
        backoff = min(backoff * 2, 300)

        if _livez(port):
            log.warning("127.0.0.1:%s answers without our server -- another "
                        "instance took over, exiting", port)
            return


if __name__ == "__main__":
    if os.environ.get(CHILD_FLAG) == "1":
        _run_server()
    else:
        _supervise()
