import email
import email.policy
import hashlib
import json
import mimetypes
import re
import tempfile
from datetime import datetime
from email.utils import getaddresses
from pathlib import Path

import extract_msg
import markdown
import yaml
from bs4 import BeautifulSoup
from msgforge import Message
from RTFDE.deencapsulate import DeEncapsulator
from extract_msg.structures.entry_id import OneOffRecipient

from .config import atomic_write, path_for, write_json
from .msg_body import outlook_bytes, plain_text


def signature_marker(cfg):
    marker = (cfg.get("signature_marker") or "").strip()
    if not marker:
        raise ValueError("In .local/settings.json fehlt „signature_marker“ (erste Zeile der Signatur, z. B. der volle Name).")
    return marker


def _mail_bodies(source):
    """HTML and plain text of a saved mail: .msg (Outlook for Windows) or .eml (Outlook for Mac, Apple Mail)."""
    source = Path(source)
    if source.suffix.lower() == ".eml":
        mail = email.message_from_bytes(source.read_bytes(), policy=email.policy.default)
        html_part, text_part = mail.get_body(("html",)), mail.get_body(("plain",))
        html = html_part.get_content() if html_part else ""
        text = text_part.get_content() if text_part else plain_text(html)
        return html, text
    with extract_msg.Message(str(source)) as msg:
        html = msg.htmlBody or b""
        return html.decode("utf-8", errors="replace") if isinstance(html, bytes) else html, msg.body or ""


def import_signature(source, destination, marker, end=None):
    """Keep the original signature's links and inline styling, not the invitation."""
    destination = Path(destination)
    html, body = _mail_bodies(source)
    soup = BeautifulSoup(html, "html.parser")
    paragraph = next((p for p in soup.find_all("p") if marker in p.get_text()), None)
    if paragraph is None:
        raise ValueError("„" + marker + "“ steht in dieser Mail nicht in der Signatur. Bitte die erste Zeile der Signatur genau so angeben.")
    # Only the paragraph containing the marker is taken. With signature_end set,
    # refuse a template whose signature continues elsewhere instead of cutting it.
    if end and end not in paragraph.get_text():
        raise ValueError("Signaturende nicht erkannt; Vorlage bitte gezielt prüfen.")
    body = body.replace("\r\n", "\n")
    if marker not in body:
        raise ValueError("„" + marker + "“ fehlt im Textteil der Mail.")
    start = body.index(marker)
    line = body[:start].rstrip().split("\n")[-1]
    if re.fullmatch(r"-{10,}", line):
        start = body.rfind(line, 0, start)
    atomic_write(destination / "signature.html", str(paragraph))
    atomic_write(destination / "signature.txt", body[start:].strip() + "\n")
    write_json(destination / "source.json", {
        "source": str(Path(source).resolve()),
        "sha256": hashlib.sha256(Path(source).read_bytes()).hexdigest(),
        "marker": marker,
    })
    return {"signature": str(destination), "links": len(paragraph.find_all("a")),
            "text": body[start:].strip()}


def export_signature(cfg, path):
    """Signatur.json for the web version: marker plus HTML and text of the configured signature."""
    signature_dir = path_for(cfg, "signature_dir")
    data = {"format": "msg-signatur", "marker": signature_marker(cfg),
            "html": (signature_dir / "signature.html").read_text(encoding="utf-8"),
            "text": (signature_dir / "signature.txt").read_text(encoding="utf-8")}
    if data["marker"] not in data["html"] or data["marker"] not in data["text"]:
        raise ValueError("Die konfigurierte Signatur passt nicht zu signature_marker.")
    write_json(path, data)
    return {"signature_json": str(Path(path).resolve()), "marker": data["marker"]}


def addresses(value):
    if value is None:
        return []
    values = value if isinstance(value, list) else [value]
    if not all(isinstance(x, str) for x in values):
        raise ValueError("Empfänger müssen E-Mail-Adressen als Text sein.")
    if any("\n" in x or "\r" in x for x in values):
        raise ValueError("Zeilenumbrüche sind in Empfängerfeldern nicht erlaubt.")
    # Outlook uses semicolons; do not split semicolons inside quoted names.
    tokens = []
    for value in values:
        quoted = escaped = False
        comments = 0
        chars = []
        for char in value:
            if char == '"' and not escaped and not comments:
                quoted = not quoted
            if not quoted and not escaped:
                if char == "(":
                    comments += 1
                elif char == ")":
                    comments -= 1
                    if comments < 0:
                        raise ValueError("Ungültige Klammer im Empfängerfeld.")
            if char in ",;" and not quoted and not comments and not escaped:
                if "".join(chars).strip():
                    tokens.append("".join(chars).strip())
                chars = []
            else:
                chars.append(char)
            escaped = char == "\\" and not escaped
        if quoted or comments:
            raise ValueError("Anführungszeichen oder Klammern im Empfängerfeld sind nicht geschlossen.")
        if "".join(chars).strip():
            tokens.append("".join(chars).strip())
    result = []
    seen = set()
    for token in tokens:
        parsed = getaddresses([token])
        if len(parsed) != 1 or not parsed[0][1]:
            raise ValueError("Ungültiger Empfänger: " + token)
        name, address = parsed[0]
        if not re.fullmatch(r'[^\s@<>,;:"]+@[^\s@<>,;:"]+\.[^\s@<>,;:"]+', address):
            raise ValueError("Ungültige E-Mail-Adresse: " + (address or name))
        if address.casefold() not in seen:
            result.append((address, name))
            seen.add(address.casefold())
    return result


def _filename_part(value, max_bytes):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f\x7f]', "-", value)
    value = re.sub(r"\s+", " ", value).strip(" .-") or "Entwurf"
    value = value.encode("utf-8")[:max_bytes].decode("utf-8", errors="ignore").rstrip(" .")
    if re.fullmatch(r"(?i:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\..*)?", value):
        value = "_" + value
    return value


def draft_filename(subject, recipients):
    # Only An recipients: forwarding the file must not disclose Bcc in its name.
    to = recipients["to"]
    person = (to[0][1] or to[0][0]) if to else "ohne An-Empfänger"
    extra = " (+%d weitere)" % (len(to) - 1) if len(to) > 1 else ""
    return _filename_part(subject, 90) + " – " + _filename_part(person, 55) + extra


def read_markdown(path, validate=True, marker=None):
    return parse_markdown(Path(path).read_bytes().decode("utf-8-sig"), validate, marker)


def parse_markdown(raw, validate=True, marker=None):
    # Files edited on Windows or received by mail may use CRLF line endings.
    raw = raw.replace("\r\n", "\n")
    match = re.match(r"\A---\s*\n(.*?)\n---[ \t]*\n(.*)\Z", raw, re.S)
    if not match:
        raise ValueError("Markdown braucht einen YAML-Kopf zwischen zwei --- Zeilen.")
    try:
        metadata = yaml.safe_load(match[1])
    except yaml.YAMLError as error:
        raise ValueError("Ungültiger YAML-Kopf: " + str(error)) from error
    if not isinstance(metadata, dict):
        raise ValueError("Der YAML-Kopf muss Felder wie subject und to enthalten.")
    unknown = set(metadata) - {"subject", "to", "cc", "bcc", "attachments", "images", "include_deck", "importance"}
    if unknown:
        raise ValueError("Unbekannte Markdown-Felder: " + ", ".join(sorted(unknown)))
    subject = metadata.get("subject")
    if not isinstance(subject, str) or "\n" in subject or "\r" in subject or (validate and not subject.strip()):
        raise ValueError("Ein nicht-leerer, einzeiliger subject ist erforderlich.")
    body = match[2].strip()
    if validate and not body:
        raise ValueError("Der Nachrichtentext fehlt.")
    if validate and marker and marker in body:
        raise ValueError("Die vollständige Signatur bitte aus dem Markdown entfernen; sie wird automatisch ergänzt.")
    if validate and re.search(r"\{\{[^{}]+\}\}", json.dumps(metadata, ensure_ascii=False, default=str) + body):
        raise ValueError("Nicht ausgefüllter Platzhalter {{…}} vorhanden.")
    return metadata, body


def read_editor(path):
    path = Path(path).expanduser().resolve()
    raw = path.read_bytes()
    metadata, body = parse_markdown(raw.decode("utf-8-sig"), validate=False)
    return {"metadata": metadata, "body": body, "path": str(path),
            "sha256": hashlib.sha256(raw).hexdigest()}


def save_editor(path, data):
    path = Path(path).expanduser().resolve()
    expected = data.get("sha256")
    actual = hashlib.sha256(path.read_bytes()).hexdigest() if path.exists() else None
    if actual != expected:
        raise ValueError("Die Markdown-Datei wurde außerhalb der App geändert. Bitte zuerst neu laden; dein Text im Fenster bleibt erhalten.")
    metadata, body = data.get("metadata"), data.get("body")
    if not isinstance(metadata, dict) or not isinstance(body, str):
        raise ValueError("Betreff und Nachrichtentext konnten nicht gespeichert werden.")
    text = "---\n" + yaml.safe_dump(metadata, allow_unicode=True, sort_keys=False, width=1000) + "---\n" + body + "\n"
    atomic_write(path, text)
    return read_editor(path)


def file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def prepare_draft(cfg, source, attachments=(), include_deck=False):
    """Validate and snapshot inputs before writing any MSG."""
    source = Path(source).expanduser().resolve()
    raw = source.read_bytes()
    if cfg.get("without_signature"):
        # Web version without a loaded signature: the mail ends with the typed text.
        marker, signature_html, signature_text = None, "", ""
        metadata, body = parse_markdown(raw.decode("utf-8-sig"))
    else:
        marker = signature_marker(cfg)
        metadata, body = parse_markdown(raw.decode("utf-8-sig"), marker=marker)
        signature_dir = path_for(cfg, "signature_dir")
        signature_html = (signature_dir / "signature.html").read_text(encoding="utf-8")
        signature_text = (signature_dir / "signature.txt").read_text(encoding="utf-8").strip()
        if marker not in signature_text or marker not in signature_html:
            raise ValueError("Die konfigurierte Signatur ist unvollständig oder passt nicht zu signature_marker.")
    configured = metadata.get("attachments") or []
    if not isinstance(configured, list) or not all(isinstance(x, str) for x in configured):
        raise ValueError("attachments muss eine Liste von Dateipfaden sein.")
    if not isinstance(metadata.get("include_deck", False), bool):
        raise ValueError("include_deck muss true oder false sein.")
    files = [Path(x).expanduser() for x in configured]
    files = [(x if x.is_absolute() else source.parent / x).resolve() for x in files]
    files += [Path(x).expanduser().resolve() for x in attachments]
    if include_deck or metadata.get("include_deck"):
        files.append(path_for(cfg, "slide_deck").resolve())
    files = list(dict.fromkeys(files))
    for path in files:
        if not path.is_file():
            raise ValueError("Anhang fehlt oder ist nicht lokal verfügbar: " + str(path))
    recipients = {key: addresses(metadata.get(key)) for key in ("to", "cc", "bcc")}
    importance = metadata.get("importance", "normal")
    if importance not in ("low", "normal", "high"):
        raise ValueError("importance: low, normal oder high.")
    hashes = [{"name": path.name, "bytes": path.stat().st_size, "sha256": file_hash(path)} for path in files]
    images = _inline_images(metadata.get("images"), source.parent, body)
    return {"source": source, "source_sha256": hashlib.sha256(raw).hexdigest(),
            "metadata": metadata, "body": body, "recipients": recipients, "files": files,
            "attachments": hashes, "images": images,
            "signature_html": signature_html, "signature_text": signature_text}


def _inline_images(configured, folder, body):
    """Pictures shown inside the text: `![](cid:name.png)` plus an entry in `images`."""
    configured = configured or []
    if not isinstance(configured, list):
        raise ValueError("images muss eine Liste sein.")
    images = []
    for entry in configured:
        entry = {"file": entry} if isinstance(entry, str) else entry
        if not isinstance(entry, dict) or not isinstance(entry.get("file"), str) or set(entry) - {"file", "width"}:
            raise ValueError("images: Einträge brauchen „file“ (und optional „width“).")
        width = entry.get("width")
        if width is not None and (not isinstance(width, int) or isinstance(width, bool) or not 1 <= width <= 2000):
            raise ValueError("images: width muss eine Pixelzahl zwischen 1 und 2000 sein.")
        path = Path(entry["file"]).expanduser()
        path = (path if path.is_absolute() else folder / path).resolve()
        if not path.is_file():
            raise ValueError("Bild fehlt: " + str(path))
        mime = mimetypes.guess_type(path.name)[0] or ""
        if not mime.startswith("image/"):
            raise ValueError("Kein Bild: " + path.name)
        if not re.fullmatch(r"[A-Za-z0-9._-]+", path.name):
            raise ValueError("Bilddateinamen nur mit Buchstaben, Ziffern, . _ -: " + path.name)
        if any(image["cid"] == path.name for image in images):
            raise ValueError("Bild doppelt angegeben: " + path.name)
        images.append({"path": path, "cid": path.name, "mime": mime, "width": width,
                       "sha256": file_hash(path), "bytes": path.stat().st_size})
    known = {image["cid"] for image in images}
    missing = sorted(set(re.findall(r"cid:([^\s)\"'>]+)", body)) - known)
    if missing:
        raise ValueError("Bild im Text nicht gefunden: " + ", ".join(missing) + ". Bitte das Bild erneut einfügen.")
    return images


def build_draft(prepared):
    metadata, body, recipients = prepared["metadata"], prepared["body"], prepared["recipients"]
    signature_html, signature_text = prepared["signature_html"], prepared["signature_text"]
    hashes = prepared["attachments"]
    rendered = markdown.markdown(body, extensions=["extra", "sane_lists", "nl2br"])
    # Preserve readable plain text without exposing Markdown syntax.
    plain = plain_text(rendered)
    full_text = plain + "\n\n" + signature_text if signature_text else plain
    content = BeautifulSoup(rendered + '<br>' + signature_html if signature_html else rendered, "html.parser")
    for paragraph in content.find_all("p"):
        default = "margin:0;font-size:11pt;" if "MsoNormal" in paragraph.get("class", []) else "margin:0 0 12pt;"
        paragraph["style"] = default + paragraph.get("style", "")
    for link in content.find_all("a"):
        link["style"] = "color:#0563c1;" + link.get("style", "")
    images = prepared.get("images", [])
    widths = {image["cid"]: image["width"] for image in images}
    for img in content.find_all("img"):
        cid = img.get("src", "")[4:] if img.get("src", "").startswith("cid:") else None
        if cid in widths:
            # Outlook (Word engine) ignores max-width; a fixed width keeps big screenshots readable.
            if widths[cid]:
                img["width"] = str(widths[cid])
                img["style"] = "width:%dpx;height:auto;" % widths[cid] + img.get("style", "")
            img["alt"] = img.get("alt") or "Bild"
    html = ('<!doctype html><html><head><meta charset="utf-8"></head>'
            '<body style="font:11pt Arial,sans-serif">' + str(content) + '</body></html>')
    # No sender or sent date: the sending account is chosen in Outlook.
    message = Message(subject=metadata["subject"], html_body=html, text_body=full_text,
                      importance=metadata.get("importance", "normal"), sent=False, **recipients)
    for path, expected in zip(prepared["files"], hashes):
        data = path.read_bytes()
        if hashlib.sha256(data).hexdigest() != expected["sha256"]:
            raise ValueError("Anhang wurde während des Exports geändert: " + str(path))
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        message.attach_bytes(path.name, data, mime_type=mime)
    for image in images:
        data = image["path"].read_bytes()
        if hashlib.sha256(data).hexdigest() != image["sha256"]:
            raise ValueError("Bild wurde während des Exports geändert: " + image["cid"])
        message.attach_bytes(image["cid"], data, mime_type=image["mime"], content_id=image["cid"])
    data = outlook_bytes(message)
    # Independently parse the generated OLE/MSG and verify draft flag and attachments.
    with extract_msg.Message(data) as check:
        flags = check.props.getValue("0E070003", 0)
        if not flags & 8 or flags & 4 or check.subject != metadata["subject"]:
            raise ValueError("MSG-Prüfung fehlgeschlagen: kein gültiger ungesendeter Entwurf.")
        actual = [hashlib.sha256(a.data).hexdigest() for a in check.attachments]
        if actual != [a["sha256"] for a in hashes] + [image["sha256"] for image in images] or signature_text not in (check.body or "").replace("\r\n", "\n"):
            raise ValueError("MSG-Prüfung fehlgeschlagen: Anhang oder Signatur stimmt nicht überein.")
        for attachment, image in zip(check.attachments[len(hashes):], images):
            if attachment.cid != image["cid"] or not attachment.hidden:
                raise ValueError("MSG-Prüfung fehlgeschlagen: Bild ist nicht im Text eingebettet: " + image["cid"])
        actual_recipients = [(int(r.type), r.email, r.name) for r in check.recipients]
        expected_recipients = [(kind, email, name or email) for kind, key in enumerate(("to", "cc", "bcc"), 1)
                               for email, name in recipients[key]]
        if actual_recipients != expected_recipients:
            raise ValueError("MSG-Prüfung fehlgeschlagen: Empfängerliste stimmt nicht überein.")
        for recipient in check.recipients:
            entry = OneOffRecipient(recipient.getStream("__substg1.0_0FFF0102"))
            if (entry.displayName, entry.addressType, entry.emailAddress) != (recipient.name, "SMTP", recipient.email):
                raise ValueError("MSG-Prüfung fehlgeschlagen: Outlook-Empfängerkennung stimmt nicht überein.")
        recovered = DeEncapsulator(check.rtfBody)
        recovered.deencapsulate()
        if recovered.content != html.encode("utf-8") or check.htmlBody != html.encode("utf-8"):
            raise ValueError("MSG-Prüfung fehlgeschlagen: HTML- und RTF-Inhalt stimmen nicht überein.")
    if file_hash(prepared["source"]) != prepared["source_sha256"]:
        raise ValueError("Markdown wurde während des Exports geändert: " + str(prepared["source"]))
    return data, html


def create_draft(cfg, source, attachments=(), include_deck=False, output=None):
    prepared = prepare_draft(cfg, source, attachments, include_deck)
    data, html = build_draft(prepared)
    out_root = Path(output).resolve() if output else path_for(cfg, "output_dir")
    out_root.mkdir(parents=True, exist_ok=True)
    subject, recipients = prepared["metadata"]["subject"], prepared["recipients"]
    slug = _filename_part(subject, 32)
    filename = draft_filename(subject, recipients)
    folder = Path(tempfile.mkdtemp(prefix=datetime.now().strftime("%Y%m%d-%H%M%S-") + slug + "-", dir=out_root))
    atomic_write(folder / (filename + ".msg"), data)
    atomic_write(folder / (filename + ".html"), html)
    manifest = {"folder": str(folder), "msg": str(folder / (filename + ".msg")),
                "preview": str(folder / (filename + ".html")), "subject": subject,
                "recipients": recipients, "attachments": prepared["attachments"],
                "source": str(prepared["source"]), "draft": True}
    write_json(folder / "manifest.json", manifest)
    return manifest
