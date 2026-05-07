"""Launcher wrapper for environments that ignore the configured working directory.

Useful for Windows / Claude Desktop (MSIX) setups where the configured
``cwd`` is not reliably applied to the spawned Python process.
"""

from __future__ import annotations

import os
import runpy
import sys
from pathlib import Path


def resolve_repo_root() -> Path:
    """Return the repository root for launching ``src.main``.

    ``EWS_MCP_ROOT`` can override the auto-detected location when the file is
    copied elsewhere or launched through a wrapper.
    """
    override = os.environ.get("EWS_MCP_ROOT")
    if override:
        return Path(override).expanduser().resolve()
    return Path(__file__).resolve().parent


def main() -> None:
    repo_root = resolve_repo_root()
    root_str = str(repo_root)
    os.chdir(root_str)
    if root_str not in sys.path:
        sys.path.insert(0, root_str)
    runpy.run_module("src.main", run_name="__main__")


if __name__ == "__main__":
    main()
