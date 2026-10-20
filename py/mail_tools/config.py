import json
import os
import sys
import tempfile
from pathlib import Path

# Shipped files (templates, scripts, docs): the repository, or the bundle inside the app.
RESOURCES = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))


def _data_root():
    # Private working folder: settings, mail index, drafts and output.
    if os.environ.get("MAILTOOL_HOME"):
        return Path(os.environ["MAILTOOL_HOME"]).expanduser().resolve()
    if getattr(sys, "frozen", False):
        return Path.home() / "Documents" / "E-Mail-App"
    return RESOURCES


ROOT = _data_root()


def read_config():
    path = Path(os.environ.get("MAILTOOL_CONFIG", ROOT / ".local/settings.json"))
    if not path.exists():
        raise ValueError("Konfiguration fehlt: .local/settings.json (siehe README.md).")
    cfg = json.loads(path.read_text(encoding="utf-8"))
    cfg["root"] = ROOT
    return cfg


def path_for(cfg, key):
    if not cfg.get(key):
        raise ValueError("In .local/settings.json ist „" + key + "“ nicht eingetragen (siehe README.md).")
    path = Path(cfg[key]).expanduser()
    return path if path.is_absolute() else cfg["root"] / path


def atomic_write(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".mailtool-", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data if isinstance(data, bytes) else data.encode("utf-8"))
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def write_json(path, value):
    atomic_write(path, json.dumps(value, ensure_ascii=False, indent=2, default=str) + "\n")
