import os, sys
from pathlib import Path

APP_DIRNAME = "FollowUpApp"

def app_data_dir():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    p = Path(base) / APP_DIRNAME
    p.mkdir(parents=True, exist_ok=True)
    return p

def resource_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent
