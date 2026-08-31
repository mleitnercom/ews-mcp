"""Wrapper fuer Claude Desktop (MSIX): cwd setzen + Windows-Trust-Store nutzen.

Zwei Gruende:

1. Die Store-Variante von Claude Desktop reicht das cwd-Feld aus
   claude_desktop_config.json nicht korrekt an den Kindprozess weiter.
   ews-mcp laedt .env relativ zum Arbeitsverzeichnis -> ohne chdir keine
   Credentials. Gleiches Muster wie beim v3-Setup (Vault: PBS-Exchange-MCP-Setup).

2. owa.pbs-holding.com traegt ein Zertifikat der internen PBS-PKI
   (PBSRootCA2020 -> PBS-IssuingCA1). Windows vertraut dieser Kette, Pythons
   certifi-Bundle nicht -> SSLCertVerificationError. truststore legt den
   Windows-Zertifikatspeicher unter ssl/requests und macht die Verifikation
   damit korrekt, statt sie abzuschalten (EWS_INSECURE_SKIP_VERIFY).
"""
import os
import sys

import truststore

truststore.inject_into_ssl()

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)

from ewsmcp.main import main

if __name__ == "__main__":
    main()
