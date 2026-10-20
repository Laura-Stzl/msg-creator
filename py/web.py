"""Browser glue: the app's export code on Pyodide's in-memory filesystem.

Nothing leaves the browser; files exist only in this tab's memory. The page ships
no signature: each user loads their own (Signatur.json or a sent mail as .msg).
"""
import io
import json
import shutil
import zipfile
from pathlib import Path

import yaml

from mail_tools import batches
from mail_tools.drafts import create_draft, import_signature

WORK = Path("/work")
INBOX = WORK / "Eingabe"
SIGNATURE = Path("/signatur")


def _signature_json():
    marker = (SIGNATURE / "marker.txt").read_text(encoding="utf-8")
    return json.dumps({"format": "msg-signatur", "marker": marker,
                       "html": (SIGNATURE / "signature.html").read_text(encoding="utf-8"),
                       "text": (SIGNATURE / "signature.txt").read_text(encoding="utf-8")}, ensure_ascii=False)


def set_signature(data):
    """Load a Signatur.json (as created by this page or `mailtool draft signature --json`)."""
    try:
        data = json.loads(data)
    except ValueError as error:
        raise ValueError("Die Datei ist keine gültige Signatur.json.") from error
    if not isinstance(data, dict) or not all(isinstance(data.get(key), str) and data[key].strip() for key in ("marker", "html", "text")):
        raise ValueError("Signatur.json braucht die Felder marker, html und text.")
    if data["marker"] not in data["html"] or data["marker"] not in data["text"]:
        raise ValueError("Der Marker kommt in der Signatur nicht vor.")
    shutil.rmtree(SIGNATURE, ignore_errors=True)
    SIGNATURE.mkdir(parents=True)
    (SIGNATURE / "marker.txt").write_text(data["marker"], encoding="utf-8")
    (SIGNATURE / "signature.html").write_text(data["html"], encoding="utf-8")
    (SIGNATURE / "signature.txt").write_text(data["text"], encoding="utf-8")
    return _signature_json()


def signature_from_msg(path, marker):
    """Take the signature from a sent mail, starting at the paragraph containing the marker."""
    marker = marker.strip()
    if not marker:
        raise ValueError("Bitte die erste Zeile der Signatur angeben, z. B. den vollen Namen.")
    shutil.rmtree(SIGNATURE, ignore_errors=True)
    import_signature(path, SIGNATURE, marker)
    (SIGNATURE / "marker.txt").write_text(marker, encoding="utf-8")
    return _signature_json()


def config():
    if not (SIGNATURE / "marker.txt").exists():
        raise ValueError("Bitte zuerst eine Signatur laden.")
    return {"root": WORK, "signature_dir": str(SIGNATURE), "output_dir": str(WORK / "output"),
            "signature_marker": (SIGNATURE / "marker.txt").read_text(encoding="utf-8")}


def reset():
    shutil.rmtree(WORK, ignore_errors=True)
    INBOX.mkdir(parents=True)
    return str(INBOX)


def _zip(folder):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(folder.rglob("*")):
            if path.is_file() and path.name not in (".auftrag.json", ".lock") and not path.name.endswith(".pending.msg"):
                archive.write(path, path.relative_to(folder).as_posix().replace(".vorschau/", "Vorschau/"))
    return buffer.getvalue()


def batch(name="", progress=None):
    files = sorted(p for p in INBOX.iterdir() if p.is_file())
    sources = [p for p in files if p.suffix.lower() in (".md", ".markdown", ".csv", ".xlsx")]
    if not sources:
        raise ValueError("Keine Entwürfe (.md) und keine Liste (.csv/.xlsx) ausgewählt.")
    cfg = config()
    original = batches._persist

    def persist(state, folder, progress_file):
        original(state, folder, progress_file)
        if progress is not None:
            progress(json.dumps(batches._summary(state, folder), ensure_ascii=False, default=str))

    batches._persist = persist
    try:
        result = batches.create_batch(cfg, [str(p) for p in sources], name=name or None)
    finally:
        batches._persist = original
    (WORK / "ergebnis.zip").write_bytes(_zip(Path(result["folder"])))
    return json.dumps(result, ensure_ascii=False, default=str)


def single(fields):
    fields = json.loads(fields)
    meta = {key: fields.get(key, "") for key in ("subject", "to", "cc", "bcc")}
    meta["attachments"] = fields.get("attachments", [])
    if fields.get("images"):
        meta["images"] = fields["images"]
    path = INBOX / "Entwurf.md"
    path.write_text("---\n" + yaml.safe_dump(meta, allow_unicode=True, sort_keys=False) + "---\n"
                    + fields.get("body", "") + "\n", encoding="utf-8")
    result = create_draft(config(), path)
    return json.dumps({"msg": result["msg"], "preview": result["preview"]}, ensure_ascii=False)
