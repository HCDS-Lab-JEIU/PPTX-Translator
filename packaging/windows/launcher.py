"""Startup for the frozen app; normal Python execution remains in app.py."""
import os
from pathlib import Path
import sys


def main():
    if getattr(sys, "frozen", False):
        local = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData/Local"))
        workspace = local / "PPTTranslator"
        workspace.mkdir(parents=True, exist_ok=True)
        os.chdir(workspace)
    from app import main as run_app
    return run_app()


if __name__ == "__main__":
    raise SystemExit(main())
