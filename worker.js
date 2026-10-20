// Runs Python (Pyodide) off the main thread so the page stays responsive during exports.
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/";
import { loadPyodide } from "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/pyodide.mjs";

let py, web;
const ready = (async () => {
  const say = (text) => postMessage({ type: "status", text });
  say("Lade Python …");
  py = await loadPyodide({ indexURL: PYODIDE });
  const cfg = await (await fetch("config.json", { cache: "no-cache" })).json();
  say("Lade Programmdateien …");
  for (const file of ["config.json", ...cfg.files]) {
    const data = new Uint8Array(await (await fetch(file, { cache: "no-cache" })).arrayBuffer());
    py.FS.mkdirTree("/site/" + file.split("/").slice(0, -1).join("/"));
    py.FS.writeFile("/site/" + file, data);
  }
  say("Installiere Pakete (beim ersten Mal ca. 20 Sekunden) …");
  await py.loadPackage("micropip");
  const micropip = py.pyimport("micropip");
  await micropip.install(cfg.wheels.map((wheel) => new URL(wheel, self.location.href).href));
  await micropip.install(cfg.packages);
  py.runPython("import sys; sys.path.insert(0, '/site/py')");
  web = py.pyimport("web");
  postMessage({ type: "ready" });
})().catch((error) => postMessage({ type: "fatal", text: String(error) }));

function pythonError(error) {
  // Show the Python message, not the traceback.
  const lines = String(error.message || error).trim().split("\n");
  return lines[lines.length - 1].replace(/^\w*(Error|Exception): /, "");
}

function writeInputs(files) {
  const inbox = web.reset();
  for (const file of files) py.FS.writeFile(inbox + "/" + file.name, new Uint8Array(file.data));
}

onmessage = async ({ data }) => {
  await ready;
  try {
    if (data.type === "signature-json") {
      postMessage({ type: "signature", signature: JSON.parse(web.set_signature(data.text)) });
      return;
    }
    if (data.type === "signature-msg") {
      const path = "/tmp/Beispiel" + (data.name.toLowerCase().endsWith(".eml") ? ".eml" : ".msg");
      py.FS.writeFile(path, new Uint8Array(data.file));
      postMessage({ type: "signature", signature: JSON.parse(web.signature_from_msg(path, data.marker)) });
      return;
    }
    writeInputs(data.files);
    if (data.type === "batch") {
      const progress = (json) => postMessage({ type: "progress", ...JSON.parse(json) });
      const result = JSON.parse(web.batch(data.name || "", progress));
      const zip = py.FS.readFile("/work/ergebnis.zip");
      postMessage({ type: "batch", result, zip }, [zip.buffer]);
    } else {
      const result = JSON.parse(web.single(JSON.stringify(data.fields)));
      const msg = py.FS.readFile(result.msg);
      const preview = py.FS.readFile(result.preview, { encoding: "utf8" });
      postMessage({ type: "single", name: result.msg.split("/").pop(), msg, preview }, [msg.buffer]);
    }
  } catch (error) {
    postMessage({ type: "error", text: pythonError(error), signature: data.type.startsWith("signature") });
  }
};
