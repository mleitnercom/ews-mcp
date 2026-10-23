"""Gemeinsamer HTTP-Server fuer alle Clients (Claude Desktop, Claude Code, Codex).

Statt dass jeder Client per stdio seine eigene Instanz startet (drei
Exchange-Verbindungen, die sich das Throttling-Budget teilen), laeuft genau
eine Instanz dauerhaft auf 127.0.0.1:MCP_PORT. Die Clients verbinden sich
ueber http_bridge.py damit.

Gestartet von der geplanten Aufgabe "ews-mcp HTTP" bei der Anmeldung, mit
pythonw.exe (kein Konsolenfenster). Ein zweiter Start scheitert am belegten
Port und beendet sich -- es bleibt bei einer Instanz.

Transport und Host werden hier gesetzt, nicht in .env: .env gilt auch fuer
den stdio-Betrieb, und dort darf MCP_TRANSPORT=http nicht stehen.
"""
import os
import runpy
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

os.environ["MCP_TRANSPORT"] = "http"
os.environ["MCP_HOST"] = "127.0.0.1"

# pythonw.exe hat kein stdout/stderr; der StreamHandler in run_server.py
# braucht aber einen Strom. Das Datei-Log unter DATA_DIR/logs bleibt.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

runpy.run_path(os.path.join(HERE, "run_server.py"), run_name="__main__")
