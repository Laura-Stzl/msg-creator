"""Outlook HTML with a readable RTF alternative, including links and Unicode.

msgforge 1.0.0 exposes stylesheet contents in its RTF alternative and loses
paragraph breaks. Replace that stream through extract-msg's public OLE writer.
MS-OXRTFEX: HTMLTAG destinations hold HTML; HTMLRTF regions hold only RTF.
https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/ba056776-cf8e-4037-b864-15b72c8c582f
"""
import io
import re
import struct

import compressed_rtf
import extract_msg
from bs4 import BeautifulSoup, Comment, NavigableString
from extract_msg.ole_writer import OleWriter

_BLOCKS = {"p", "div", "h1", "h2", "h3", "h4", "blockquote", "pre", "ul", "ol", "li", "table", "tr"}


def _normalized_text(node):
    value = re.sub(r"[\t\r\n ]+", " ", str(node))
    previous, following = node.previous_sibling, node.next_sibling
    if getattr(previous, "name", None) in _BLOCKS | {"br"} or (previous is None and node.parent.name in _BLOCKS):
        value = value.lstrip(" ")
    if getattr(following, "name", None) in _BLOCKS | {"br"} or (following is None and node.parent.name in _BLOCKS):
        value = value.rstrip(" ")
    return value


def _body_soup(html):
    soup = BeautifulSoup(html, "html.parser")
    for node in soup.find_all(["head", "style", "script"]):
        node.decompose()
    for node in soup.find_all(string=lambda text: isinstance(text, Comment)):
        node.extract()
    return soup.body or soup


def plain_text(html):
    soup = _body_soup(html)
    for node in soup.find_all(string=True):
        if not node.find_parent("pre"):
            node.replace_with(_normalized_text(node))
    for node in soup.find_all("br"):
        node.replace_with("\n")
    for node in soup.find_all(["p", "div", "h1", "h2", "h3", "blockquote", "pre"]):
        node.insert_after("\n\n")
    for node in soup.find_all(["li", "tr"]):
        node.insert_after("\n")
    return re.sub(r"\n[ \t]+", "\n", soup.get_text()).strip()


def _escape(text):
    # RTF Unicode escapes are signed UTF-16 code units, including surrogate pairs.
    out = []
    for (unit,) in struct.iter_unpack("<H", text.encode("utf-16-le")):
        if unit in (92, 123, 125):
            out.append("\\" + chr(unit))
        elif 32 <= unit < 127:
            out.append(chr(unit))
        else:
            out.append("\\u%d?" % (unit if unit < 32768 else unit - 65536))
    return "".join(out)


def encapsulated_rtf(html):
    colors = ["000000", "0563c1"]

    def render(node):
        if isinstance(node, NavigableString):
            value = str(node)
            if node.find_parent("pre"):
                return _escape(value).replace("\\u10?", "\\line ").replace("\\u9?", "\\tab ")
            return _escape(_normalized_text(node))
        name = node.name
        if name == "br":
            return "\\line "
        if name == "hr":
            return "\\par " + _escape("————————") + "\\par "
        if name == "img":
            return _escape(node.get("alt", ""))
        content = "".join(render(child) for child in node.children)
        if name == "a" and node.get("href"):
            # Quotes/newlines must not become RTF field instructions.
            url = node["href"].replace('"', "%22").replace("\r", "%0D").replace("\n", "%0A")
            content = ('{\\field{\\*\\fldinst HYPERLINK "' + _escape(url) +
                       '"}{\\fldrslt{\\ul\\cf2 ' + content + '}}}')
        controls = {"b": "\\b", "strong": "\\b", "i": "\\i", "em": "\\i",
                    "u": "\\ul", "s": "\\strike", "del": "\\strike"}.get(name, "")
        color = re.search(r"(?:^|;)\s*color\s*:\s*(#[0-9a-fA-F]{6}|black)\s*(?:;|$)", node.get("style", ""))
        if color:
            value = "000000" if color[1] == "black" else color[1][1:].lower()
            if value not in colors:
                colors.append(value)
            controls += "\\cf%d" % (colors.index(value) + 1)
        if controls:
            content = "{" + controls + " " + content + "}"
        if name in ("p", "div", "h1", "h2", "h3", "h4", "blockquote", "pre"):
            spacing = 0 if "MsoNormal" in node.get("class", []) else 240
            return "{\\pard\\sa%d " % spacing + content.strip() + "\\par}"
        if name == "li":
            prefix = "• "
            if node.parent.name == "ol":
                prefix = str(len(node.find_previous_siblings("li")) + int(node.parent.get("start", 1))) + ". "
            return _escape(prefix) + content + "\\par "
        if name in ("td", "th"):
            return content + "\\tab "
        if name == "tr":
            return content + "\\par "
        return content

    fallback = render(_body_soup(html))
    color_table = "{\\colortbl;" + "".join(
        "\\red%d\\green%d\\blue%d;" % tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
        for color in colors) + "}"
    rtf = ("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1\\deff0"
           "{\\fonttbl{\\f0\\fswiss Arial;}}" + color_table + "\\uc1\n"
           "{\\*\\htmltag64 " + _escape(html) + "}\n"
           "\\htmlrtf1\\pard\\plain\\f0\\fs22\\cf0 " + fallback + "\\htmlrtf0}")
    return rtf.encode("ascii")


def _add_recipient_ids(writer, original):
    """Give SMTP recipients self-contained Outlook address-book identities.

    MS-OXCDATA 2.2.5.1 / MAPI CreateOneOff: Unicode strings, MIME delivery.
    https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/one-off-entry-identifiers
    """
    provider = bytes.fromhex("812b1fa4bea310199d6e00dd010f5402")
    for index, recipient in enumerate(original.recipients):
        storage = "__recip_version1.0_#%08X" % index
        entry_id = (b"\0" * 4 + provider + struct.pack("<HH", 0, 0x8001) +
                    "\0".join((recipient.name, "SMTP", recipient.email, "")).encode("utf-16-le"))
        values = {0x0FFF0102: entry_id,  # PidTagEntryId
                  0x0FF90102: entry_id,  # PidTagRecordKey
                  0x300B0102: ("SMTP:" + recipient.email.upper() + "\0").encode("utf-8")}
        props = bytearray(recipient.getStream("__properties_version1.0"))
        for tag, value in values.items():
            path = [storage, "__substg1.0_%08X" % tag]
            if original.exists(path):
                writer.editEntry(path, data=value)
            else:
                writer.addEntry(path, data=value)
            for pos in range(8, len(props), 16):
                if struct.unpack_from("<I", props, pos)[0] == tag:
                    struct.pack_into("<II", props, pos + 8, len(value), 0)
                    break
            else:
                props.extend(struct.pack("<IIII", tag, 6, len(value), 0))
        writer.editEntry([storage, "__properties_version1.0"], data=bytes(props))


def outlook_bytes(message):
    """Write Outlook body alternatives and resolvable SMTP recipient identities."""
    compressed = compressed_rtf.compress(encapsulated_rtf(message.html_body))
    writer = OleWriter()
    with extract_msg.Message(message.as_bytes()) as original:
        writer.fromMsg(original)
        _add_recipient_ids(writer, original)
        props = bytearray(original.getStream("__properties_version1.0"))
    found_rtf = False
    native_body = False
    for pos in range(32, len(props), 16):
        tag = struct.unpack_from("<I", props, pos)[0]
        if tag == 0x10090102:  # PidTagRtfCompressed: update binary stream length.
            struct.pack_into("<I", props, pos + 8, len(compressed))
            found_rtf = True
        elif tag == 0x10160003:  # PidTagNativeBody: HTML is the preferred body.
            struct.pack_into("<I", props, pos + 8, 3)
            native_body = True
    if not found_rtf:
        raise ValueError("MSG enthält keinen RTF-Nachrichtentext.")
    if not native_body:
        props.extend(struct.pack("<IIII", 0x10160003, 6, 3, 0))
    writer.editEntry("__properties_version1.0", data=bytes(props))
    writer.editEntry("__substg1.0_10090102", data=compressed)
    result = io.BytesIO()
    writer.write(result)
    return result.getvalue()
