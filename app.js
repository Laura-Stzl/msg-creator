const $ = (id) => document.getElementById(id);
const worker = new Worker("worker.js", { type: "module" });
const status = $("status");
let ready = false;
let busy = false;
let batchFiles = [];
let signature = null;
let signatureFile = null;
// Pictures pasted into the single mail: cid -> { blob, url, width }. Referenced in the text as ![Bild](cid:…).
const images = new Map();
let previewImages = new Map();
const STORE = "msg-signatur";
const store = {
  get() { try { return localStorage.getItem(STORE); } catch { return null; } },
  set(value) { try { value ? localStorage.setItem(STORE, value) : localStorage.removeItem(STORE); } catch { /* privater Modus */ } },
};

function setStatus(text, kind = "muted") { status.textContent = text; status.className = kind; }
function updateButtons() {
  $("single-run").disabled = !ready || busy || !signature;
  $("batch-run").disabled = !ready || busy || !signature || batchFiles.length === 0;
  $("sig-run").disabled = !ready || busy || !signatureFile || !$("sig-marker").value.trim();
}
function download(label, name, blob, primary = false) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob); link.download = name; link.textContent = "⬇ " + label;
  link.className = primary ? "dl primary" : "dl"; link.title = name;
  return link;
}
// Copy the finished mail (text + signature) for pasting into webmail: formatted and as plain text.
function copyButton(preview) {
  const button = Object.assign(document.createElement("button"), { className: "dl", textContent: "⧉ Mailtext kopieren" });
  button.addEventListener("click", async () => {
    const body = preview.contentDocument.body;
    // Webmail cannot resolve cid: pictures; copy them as embedded data.
    const clone = body.cloneNode(true);
    await Promise.all([...clone.querySelectorAll("img[data-cid]")].map(async (img) => {
      const image = previewImages.get(img.dataset.cid);
      if (image) img.src = await dataUrl(image.blob);
      img.removeAttribute("data-cid");
    }));
    const html = '<div style="' + (body.getAttribute("style") || "") + '">' + clone.innerHTML + "</div>";
    try {
      await navigator.clipboard.write([new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([body.innerText], { type: "text/plain" }),
      })]);
    } catch {
      // Older browsers: copy the selected preview instead.
      const range = preview.contentDocument.createRange(); range.selectNodeContents(body);
      const selection = preview.contentWindow.getSelection(); selection.removeAllRanges(); selection.addRange(range);
      preview.contentDocument.execCommand("copy"); selection.removeAllRanges();
    }
    button.textContent = "✓ Kopiert – jetzt in die Mail einfügen";
    setTimeout(() => (button.textContent = "⧉ Mailtext kopieren"), 3000);
  });
  return button;
}
function note(text, className = "ok") { return Object.assign(document.createElement("p"), { className, innerHTML: text }); }
// Show the result, start the main download right away and keep buttons for repeating it.
function finish(box, message, links, extra = []) {
  const actions = Object.assign(document.createElement("div"), { className: "actions" });
  actions.append(...links);
  box.replaceChildren(message, actions, ...extra);
  box.hidden = false;
  links[0].click();
  box.scrollIntoView({ behavior: "smooth", block: "nearest" });
}
async function read(files) {
  return Promise.all([...files].map(async (file) => ({ name: file.name, data: await file.arrayBuffer() })));
}
function dataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result); reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

// Pictures in the text (single mail only).
const IMAGES = "msg-bilder";
const IMAGE_TYPES = { "image/png": "png", "image/jpeg": "jpg", "image/gif": "gif" };
const MAX_WIDTH = 600;
function referenced() { return [...images.keys()].filter((cid) => $("body").value.includes("cid:" + cid)); }
async function saveImages() {
  try {
    const list = await Promise.all([...images].map(async ([cid, image]) => ({ cid, width: image.width, data: await dataUrl(image.blob) })));
    localStorage.setItem(IMAGES, JSON.stringify(list));
  } catch {
    try { localStorage.removeItem(IMAGES); } catch { /* ohne Speicher */ }
    setStatus("Hinweis: Die Bilder sind zu groß, um sie beim Neuladen zu behalten. Sie bleiben bis zum Neuladen erhalten.");
  }
}
function showImages() {
  $("image-list").replaceChildren(...[...images].map(([cid, image]) => {
    const item = document.createElement("li");
    const remove = Object.assign(document.createElement("button"), { className: "secondary", textContent: "Entfernen" });
    remove.addEventListener("click", () => {
      URL.revokeObjectURL(image.url); images.delete(cid);
      $("body").value = $("body").value.replace(new RegExp("\\n?!\\[[^\\]]*\\]\\(cid:" + cid.replace(/\./g, "\\.") + "\\)\\n?", "g"), "\n");
      saveDraft(); saveImages(); showImages();
    });
    const thumb = Object.assign(document.createElement("img"), { src: image.url, alt: cid });
    const label = document.createElement("div");
    label.textContent = referenced().includes(cid) ? cid : cid + " (nicht im Text)";
    item.append(thumb, label, remove);
    return item;
  }));
}
// Browsers paste PNG; Outlook cannot show WebP/HEIC etc., so convert those to PNG.
async function normalize(file) {
  const bitmap = await createImageBitmap(file);
  let blob = file;
  if (!IMAGE_TYPES[file.type]) {
    const canvas = Object.assign(document.createElement("canvas"), { width: bitmap.width, height: bitmap.height });
    canvas.getContext("2d").drawImage(bitmap, 0, 0);
    blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  }
  // Screenshots on high-resolution screens are twice as wide as shown.
  const natural = bitmap.width / (file.type === "image/png" && bitmap.width > MAX_WIDTH ? Math.max(1, Math.min(2, window.devicePixelRatio || 1)) : 1);
  return { blob, width: Math.max(1, Math.min(MAX_WIDTH, Math.round(natural))) };
}
async function addImages(files, at) {
  const textarea = $("body");
  let position = at ?? textarea.selectionStart ?? textarea.value.length;
  let added = 0;
  for (const file of files) {
    if (!file.type.startsWith("image/")) continue;
    let image;
    try { image = await normalize(file); } catch { setStatus("Bild konnte nicht gelesen werden: " + (file.name || "Zwischenablage"), "err"); continue; }
    const cid = "bild-" + Date.now().toString(36) + "-" + (images.size + 1) + "." + IMAGE_TYPES[image.blob.type];
    images.set(cid, { ...image, url: URL.createObjectURL(image.blob) });
    const before = textarea.value.slice(0, position), after = textarea.value.slice(position);
    const snippet = (before && !before.endsWith("\n") ? "\n\n" : "") + "![Bild](cid:" + cid + ")" + (after.startsWith("\n") ? "\n" : "\n\n");
    textarea.value = before + snippet + after;
    position += snippet.length; added++;
  }
  if (!added) return;
  textarea.setSelectionRange(position, position); textarea.focus();
  saveDraft(); showImages(); await saveImages();
}
$("body").addEventListener("paste", (event) => {
  const files = [...event.clipboardData.files].filter((file) => file.type.startsWith("image/"));
  // Copied from Word/Excel: text plus a picture of it. Then paste the text as usual.
  if (!files.length || event.clipboardData.getData("text/plain").trim()) return;
  event.preventDefault(); addImages(files);
});
$("body").addEventListener("dragover", (event) => {
  if ([...event.dataTransfer.items].some((item) => item.kind === "file")) { event.preventDefault(); $("body").classList.add("over"); }
});
$("body").addEventListener("dragleave", () => $("body").classList.remove("over"));
$("body").addEventListener("drop", (event) => {
  $("body").classList.remove("over");
  const files = [...event.dataTransfer.files].filter((file) => file.type.startsWith("image/"));
  if (!event.dataTransfer.files.length) return;
  event.preventDefault();
  if (!files.length) { setStatus("Nur Bilder können in den Text gezogen werden; andere Dateien bitte unter „Anhänge“ wählen.", "err"); return; }
  addImages(files);
});
$("image-files").addEventListener("change", async (event) => { await addImages(event.target.files); event.target.value = ""; });
$("body").addEventListener("input", () => { if (images.size) showImages(); });
try {
  for (const { cid, width, data } of JSON.parse(localStorage.getItem(IMAGES) || "[]")) {
    const bytes = Uint8Array.from(atob(data.split(",")[1]), (c) => c.charCodeAt(0));
    const blob = new Blob([bytes], { type: data.slice(5, data.indexOf(";")) });
    images.set(cid, { blob, width, url: URL.createObjectURL(blob) });
  }
} catch { /* ohne Speicher */ }
showImages();

// Signature: Signatur.json or a sent mail (.msg) plus the signature's first line.
function showSignature(value) {
  signature = value;
  $("sig-load").hidden = !!value; $("sig-done").hidden = !value;
  $("sig-state").textContent = value ? "Signatur geladen: " + value.marker : "Noch keine Signatur geladen.";
  if (value) {
    const json = JSON.stringify(value, null, 2);
    store.set(json);
    $("sig-preview").srcdoc = '<body style="font:11pt Arial,sans-serif">' + value.html + "</body>";
    $("sig-save").href = URL.createObjectURL(new Blob([json], { type: "application/json" }));
  }
  updateButtons();
}
$("sig-json").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  busy = true; updateButtons(); setStatus("Lade Signatur …");
  worker.postMessage({ type: "signature-json", text: await file.text() });
});
$("sig-mail").addEventListener("change", (event) => {
  signatureFile = event.target.files[0] || null;
  $("sig-mail-name").textContent = signatureFile ? signatureFile.name : "";
  $("sig-msg").hidden = !signatureFile;
  if (signatureFile) $("sig-marker").focus();
  updateButtons();
});
$("sig-marker").addEventListener("input", updateButtons);
$("sig-run").addEventListener("click", async () => {
  busy = true; updateButtons(); setStatus("Lese Signatur aus der Mail …");
  const data = await signatureFile.arrayBuffer();
  worker.postMessage({ type: "signature-msg", file: data, name: signatureFile.name, marker: $("sig-marker").value }, [data]);
});
$("sig-clear").addEventListener("click", () => {
  store.set(null); $("sig-json").value = ""; $("sig-mail").value = ""; $("sig-mail-name").textContent = "";
  signatureFile = null; $("sig-msg").hidden = true; showSignature(null);
});
const saved = store.get();
if (saved) worker.postMessage({ type: "signature-json", text: saved });

// Typed text survives a reload (this browser only); chosen files cannot be kept.
const FIELDS = ["subject", "to", "cc", "bcc", "body", "batch-name"];
const DEFAULT_BODY = [
  "Liebe Anna,", "",
  "vielen Dank für das schöne Gespräch letzte Woche. Wie besprochen möchte ich Dich herzlich zu unserer " +
  "Veranstaltung einladen. Alle Details findest Du in der angehängten Übersicht.", "",
  "Ich würde mich sehr freuen, wenn Du dabei bist. Gib mir gerne kurz Bescheid, ob es bei Dir passt.", "",
  "Herzliche Grüße", "Isabell",
].join("\n");
const OLD_DEFAULT = "\n\n\nHerzliche Grüße\nIsabell";
const DRAFT = "msg-entwurf";
try { Object.entries(JSON.parse(localStorage.getItem(DRAFT) || "{}")).forEach(([id, value]) => { if (FIELDS.includes(id)) $(id).value = value; }); } catch { /* ohne Speicher */ }
function saveDraft() {
  try { localStorage.setItem(DRAFT, JSON.stringify(Object.fromEntries(FIELDS.map((id) => [id, $(id).value])))); } catch { /* ohne Speicher */ }
}
if (!$("body").value || $("body").value === OLD_DEFAULT) $("body").value = DEFAULT_BODY;
FIELDS.forEach((id) => $(id).addEventListener("input", saveDraft));
$("single-reset").addEventListener("click", () => {
  if (["subject", "to", "cc", "bcc"].some((id) => $(id).value) || $("body").value !== DEFAULT_BODY) {
    if (!confirm("Betreff, Empfänger, Text und Anhänge leeren?")) return;
  }
  ["subject", "to", "cc", "bcc"].forEach((id) => ($(id).value = ""));
  $("body").value = DEFAULT_BODY;
  images.forEach((image) => URL.revokeObjectURL(image.url)); images.clear(); saveImages(); showImages();
  $("single-files").value = ""; $("single-result").hidden = true; saveDraft(); setStatus("");
});
$("batch-reset").addEventListener("click", () => {
  batchFiles = []; $("batch-files").value = ""; $("batch-list").replaceChildren(); $("batch-name").value = "";
  $("batch-result").hidden = true; saveDraft(); setStatus(""); updateButtons();
});

// Light/dark: follows the system until chosen; the choice is remembered in this browser.
function currentTheme() {
  return document.documentElement.dataset.theme || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
}
function showThemeButton() { $("theme").textContent = currentTheme() === "dark" ? "☀ Hell" : "☾ Dunkel"; }
try { const theme = localStorage.getItem("msg-theme"); if (theme) document.documentElement.dataset.theme = theme; } catch { /* ohne Speicher */ }
showThemeButton();
$("theme").addEventListener("click", () => {
  const theme = currentTheme() === "dark" ? "light" : "dark";
  document.documentElement.dataset.theme = theme;
  try { localStorage.setItem("msg-theme", theme); } catch { /* ohne Speicher */ }
  showThemeButton();
});

// Tabs
document.querySelectorAll("[data-tab]").forEach((tab) => tab.addEventListener("click", () => {
  document.querySelectorAll("[data-tab]").forEach((other) => {
    other.setAttribute("aria-selected", String(other === tab));
    $(other.dataset.tab).hidden = other !== tab;
  });
}));

// Batch file selection (drop zone or picker); same names replace earlier files.
function addBatchFiles(files) {
  const byName = new Map(batchFiles.map((file) => [file.name, file]));
  for (const file of files) byName.set(file.name, file);
  batchFiles = [...byName.values()];
  const kind = (name) => /\.(xlsx|csv)$/i.test(name) ? "Liste" : /\.(md|markdown)$/i.test(name) ? "Entwurf" : "Anhang";
  $("batch-list").replaceChildren(...batchFiles.map((file) =>
    Object.assign(document.createElement("li"), { textContent: kind(file.name) + ": " + file.name })));
  updateButtons();
}
$("batch-files").addEventListener("change", (event) => addBatchFiles(event.target.files));
const drop = $("drop");
drop.addEventListener("dragover", (event) => { event.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (event) => { event.preventDefault(); drop.classList.remove("over"); addBatchFiles(event.dataTransfer.files); });

$("single-run").addEventListener("click", async () => {
  const files = await read($("single-files").files);
  // Only pictures still referenced in the text go into the mail.
  const used = referenced();
  const clash = used.find((cid) => files.some((file) => file.name === cid));
  if (clash) { setStatus("Ein Anhang heißt genauso wie ein Bild im Text (" + clash + "). Bitte den Anhang umbenennen.", "err"); return; }
  const pictures = await Promise.all(used.map(async (cid) => ({ name: cid, data: await images.get(cid).blob.arrayBuffer() })));
  previewImages = new Map(used.map((cid) => [cid, images.get(cid)]));
  const fields = { subject: $("subject").value, to: $("to").value, cc: $("cc").value, bcc: $("bcc").value,
                   body: $("body").value, attachments: files.map((file) => file.name),
                   images: used.map((cid) => ({ file: cid, width: images.get(cid).width })) };
  files.push(...pictures);
  busy = true; updateButtons(); $("single-result").hidden = true;
  setStatus("Erstelle MSG …");
  worker.postMessage({ type: "single", fields, files }, files.map((file) => file.data));
});

$("batch-run").addEventListener("click", async () => {
  const files = await read(batchFiles);
  busy = true; updateButtons(); $("batch-result").hidden = true;
  setStatus("Prüfe alle Mails und Anhänge …");
  worker.postMessage({ type: "batch", name: $("batch-name").value, files }, files.map((file) => file.data));
});

worker.onerror = (event) => setStatus("Start fehlgeschlagen: " + (event.message || "Worker konnte nicht geladen werden."), "err");

worker.onmessage = ({ data }) => {
  if (data.type === "status") setStatus(data.text);
  else if (data.type === "ready") { ready = true; if (!busy) setStatus(signature || saved ? "Bereit." : "Bereit. Bitte zuerst die Signatur laden."); }
  else if (data.type === "signature") { busy = false; showSignature(data.signature); setStatus("Signatur geladen. Bereit.", "ok"); }
  else if (data.type === "fatal") setStatus("Start fehlgeschlagen: " + data.text, "err");
  else if (data.type === "progress") setStatus(`${data.phase}: ${data.created} von ${data.total} MSGs fertig`);
  else if (data.type === "error") {
    busy = false; setStatus(data.text, "err");
    if (data.signature && !signature) store.set(null);
  }
  else if (data.type === "single") {
    busy = false;
    // The preview cannot resolve cid: links; show the pasted pictures from this tab instead.
    const page = new DOMParser().parseFromString(data.preview, "text/html");
    page.querySelectorAll('img[src^="cid:"]').forEach((img) => {
      const cid = img.getAttribute("src").slice(4);
      if (previewImages.has(cid)) { img.dataset.cid = cid; img.src = previewImages.get(cid).url; }
    });
    const preview = Object.assign(document.createElement("iframe"), { srcdoc: "<!doctype html>" + page.documentElement.outerHTML, title: "Vorschau" });
    finish($("single-result"),
      note("<b>✓ Fertig: Die MSG liegt in deinem Ordner „Downloads“.</b><br>Windows: in Outlook öffnen. Mac: im Webmail verschicken (MSG als Anhang), nicht über Outlook für Mac."),
      [download("MSG nochmal herunterladen", data.name, new Blob([data.msg], { type: "application/vnd.ms-outlook" }), true),
       copyButton(preview)],
      [note("Vorschau:", "muted"), preview]);
    setStatus("");
  } else if (data.type === "batch") {
    busy = false;
    const result = data.result;
    const name = ($("batch-name").value.trim() || "MSG-Entwürfe") + ".zip";
    const extra = [];
    if (!result.complete) {
      const errors = result.errors.map((e) => `Nr. ${e.number} (${e.source.split("/").pop()}): ${e.error}`).join("\n");
      extra.push(Object.assign(document.createElement("p"), { className: "err", textContent: errors }));
    }
    finish($("batch-result"),
      note(result.complete
        ? `<b>✓ Fertig: ${result.created} MSGs als ZIP in deinem Ordner „Downloads“.</b><br>Die ZIP als Anhang per Mail weitergeben oder doppelklicken: darin alle MSGs, die Übersicht.csv und HTML-Vorschauen.`
        : `<b>${result.created} von ${result.total} MSGs erstellt und als ZIP geladen.</b> Fehler siehe unten und in Übersicht.csv.`,
        result.complete ? "ok" : "err"),
      [download("ZIP nochmal herunterladen", name, new Blob([data.zip], { type: "application/zip" }), true)], extra);
    setStatus("");
  }
  updateButtons();
};
