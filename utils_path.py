from pathlib import Path
import os, sys

APP_NAME = "FollowUpAutomation"

def resource_dir() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).parent

def app_data_dir() -> Path:
    base = os.getenv("APPDATA") or os.path.expanduser("~")
    d = Path(base) / APP_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d
