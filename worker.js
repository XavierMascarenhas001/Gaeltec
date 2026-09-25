/* Gaeltec Tools - Python worker.
 *
 * Runs Pyodide (Python compiled for the browser) off the main page so the
 * page stays responsive while a tool runs. Messages from the page:
 *   init  {indexURL, sources, libzip}  load Python + packages, write the tool files
 *   mount {items:[{rel, file}]}        make dropped files readable as /in/<n>/<rel>
 *                                      (read on demand - nothing is copied into memory)
 *   unmount {dir}
 *   call  {fn, kwargs, needs}          run bridge.<fn>(**kwargs); logs stream back,
 *                                      output files come back as bytes
 */
let pyodide = null;
const loaded = new Set();
let mountSeq = 0;
let currentId = null;   // the call whose log receives print() output
const post = (m, t) => self.postMessage(m, t || []);

// pillow up front: openpyxl only checks for it once, on first import (needed for the report logos)
const BASE_PACKAGES = ["pandas", "numpy", "python-dateutil", "pytz", "six", "typing-extensions", "pillow"];

async function init(m) {
  post({type: "status", text: "Downloading Python (first time only, ~30 s)…"});
  importScripts(m.indexURL + "pyodide.js");
  pyodide = await loadPyodide({indexURL: m.indexURL});
  // print() from the tools (e.g. "Image load failed") goes to the running tool's log
  const toLog = (s) => { if (currentId !== null) post({type: "log", id: currentId, text: s}); else console.log(s); };
  pyodide.setStdout({batched: toLog});
  pyodide.setStderr({batched: toLog});
  post({type: "status", text: "Loading pandas…"});
  await pyodide.loadPackage(BASE_PACKAGES);
  BASE_PACKAGES.forEach(p => loaded.add(p));
  const FS = pyodide.FS;
  FS.mkdirTree("/app"); FS.mkdirTree("/in"); FS.mkdirTree("/out");
  for (const [name, code] of Object.entries(m.sources)) {
    const full = "/app/" + name;
    FS.mkdirTree(full.slice(0, full.lastIndexOf("/")));
    FS.writeFile(full, code);
  }
  const bin = Uint8Array.from(atob(m.libzip), c => c.charCodeAt(0));
  FS.writeFile("/tmp/pylibs.zip", bin);
  await pyodide.runPythonAsync(`
import sys, site, zipfile, warnings
zipfile.ZipFile("/tmp/pylibs.zip").extractall(site.getsitepackages()[0])
sys.path.insert(0, "/app")
warnings.simplefilter("ignore", FutureWarning)
import py_boot, bridge, json, traceback

def _gt_run(fn, kwargs_json, log):
    try:
        res = getattr(bridge, fn)(**json.loads(kwargs_json), log=log)
        return json.dumps({"ok": True, "result": res}, default=str)
    except bridge.UserError as e:
        return json.dumps({"ok": False, "error": str(e)})
    except Exception as e:
        tb = traceback.format_exc()
        log(tb)
        return json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"})
`);
  post({type: "status", text: "ready", ready: true});
}

function mount(m) {
  const dir = "/in/" + (++mountSeq);
  const FS = pyodide.FS;
  FS.mkdirTree(dir);
  FS.mount(FS.filesystems.WORKERFS, {blobs: m.items.map(i => ({name: i.rel, data: i.file}))}, dir);
  return {dir};
}

function unmount(m) {
  try { pyodide.FS.unmount(m.dir); } catch (e) { /* already gone */ }
  return {};
}

async function call(m) {
  const need = (m.needs || []).filter(p => !loaded.has(p));
  if (need.length) {
    post({type: "log", id: m.id, text: "Loading " + need.join(", ") + " (first use only)…"});
    await pyodide.loadPackage(need);
    need.forEach(p => loaded.add(p));
  }
  const log = (s) => post({type: "log", id: m.id, text: String(s)});
  const runner = pyodide.globals.get("_gt_run");
  let out;
  currentId = m.id;
  try { out = JSON.parse(runner(m.fn, m.kwargs, log)); }
  finally { runner.destroy(); currentId = null; }
  if (!out.ok) throw new Error(out.error);
  // A list of paths under /out = files the tool wrote -> send them back.
  const res = out.result;
  const files = [], transfer = [];
  if (Array.isArray(res) && res.every(p => typeof p === "string" && p.startsWith("/out/"))) {
    for (const p of res) {
      const data = pyodide.FS.readFile(p);
      files.push({name: p.split("/").pop(), data});
      transfer.push(data.buffer);
      try { pyodide.FS.unlink(p); } catch (e) {}
    }
    return [{result: null, files}, transfer];
  }
  return [{result: res, files}, transfer];
}

self.onmessage = async (ev) => {
  const m = ev.data;
  try {
    if (m.type === "init") { await init(m); post({type: "done", id: m.id}); return; }
    if (m.type === "mount") { post({type: "done", id: m.id, ...mount(m)}); return; }
    if (m.type === "unmount") { post({type: "done", id: m.id, ...unmount(m)}); return; }
    if (m.type === "call") { const [r, t] = await call(m); post({type: "done", id: m.id, ...r}, t); return; }
  } catch (e) {
    post({type: "error", id: m.id, error: e && e.message ? e.message : String(e)});
  }
};
