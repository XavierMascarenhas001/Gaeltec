"""
Gaeltec Tools - web server
==========================
Runs the five Gaeltec Python tools behind one web page (index.html).

 * The browser never touches files directly: every file/folder is picked
   from the SERVER's view of the network share, restricted to the folders
   listed under "roots" in config.json.
 * Only computers listed in "allowed_clients" (IP addresses / subnets) can
   open the page or call the API. An optional shared "access_key" can be
   added on top.
 * Long jobs run in background threads, so several people can run tools at
   the same time; the page polls for live progress.

Start it with:   python app.py        (or double-click start_server.bat)
Then open:       http://<this-computer-name>:8080
"""
import io
import ipaddress
import json
import os
import threading
import time
import traceback
import uuid
from collections import OrderedDict
from datetime import datetime

from flask import Flask, abort, jsonify, request, send_file, send_from_directory

import tool_aggregate_cf
import tool_build_master
import tool_materials_report
import tool_tp_to_cf
import tool_work_instructions

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.environ.get("GAELTEC_CONFIG", os.path.join(HERE, "config.json"))


def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg.setdefault("host", "0.0.0.0")
    cfg.setdefault("port", 8080)
    cfg.setdefault("allowed_clients", ["127.0.0.1", "::1"])
    cfg.setdefault("access_key", "")
    cfg.setdefault("roots", [])
    cfg.setdefault("cors_origins", [])
    return cfg


CONFIG = load_config()
ALLOWED_NETS = [ipaddress.ip_network(c, strict=False) for c in CONFIG["allowed_clients"]]

app = Flask(__name__, static_folder=None)


# ---------------------------------------------------------------------------
# Access control
# ---------------------------------------------------------------------------
def _client_allowed(addr):
    try:
        ip = ipaddress.ip_address(addr)
    except ValueError:
        return False
    if ip.version == 6 and ip.ipv4_mapped:
        ip = ip.ipv4_mapped
    return any(ip in net for net in ALLOWED_NETS if net.version == ip.version)


@app.before_request
def _guard():
    if request.method == "OPTIONS":
        return  # CORS pre-flight
    if not _client_allowed(request.remote_addr or ""):
        abort(403, description=f"This computer ({request.remote_addr}) is not on the allowed list.")
    if request.path.startswith("/api/") and request.path != "/api/ping" and CONFIG["access_key"]:
        key = request.headers.get("X-Access-Key") or request.args.get("key", "")
        if key != CONFIG["access_key"]:
            abort(401, description="Access key required.")


@app.after_request
def _cors(resp):
    origin = request.headers.get("Origin")
    if origin and (origin in CONFIG["cors_origins"] or "*" in CONFIG["cors_origins"]):
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type, X-Access-Key"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    return resp


@app.errorhandler(400)
@app.errorhandler(401)
@app.errorhandler(403)
@app.errorhandler(404)
@app.errorhandler(409)
def _json_error(e):
    return jsonify({"error": getattr(e, "description", str(e))}), e.code


class UserError(Exception):
    """A problem the user can fix (bad path, no rows, ...) - shown as-is."""


@app.errorhandler(UserError)
def _user_error(e):
    return jsonify({"error": str(e)}), 400


# ---------------------------------------------------------------------------
# Allowed folders (the server-side file browser)
# ---------------------------------------------------------------------------
def _norm(p):
    return os.path.normcase(os.path.realpath(os.path.abspath(p)))


def _roots():
    return [dict(r, _norm=_norm(r["path"])) for r in CONFIG["roots"]]


def _root_for(path):
    n = _norm(path)
    best = None
    for r in _roots():
        rn = r["_norm"]
        if n == rn or n.startswith(rn.rstrip("\\/") + os.sep):
            if best is None or len(rn) > len(best["_norm"]):
                best = r
    return best


def check_path(path, must_exist=True, want_dir=None, writable=False):
    if not path or not isinstance(path, str):
        raise UserError("No file or folder was chosen.")
    root = _root_for(path)
    if root is None:
        raise UserError(f"This location is outside the allowed folders:\n{path}")
    if writable and not root.get("writable", False):
        raise UserError(f"Saving is not allowed in '{root['name']}'.")
    if must_exist and not os.path.exists(path):
        raise UserError(f"Not found (or not reachable from the server):\n{path}")
    if want_dir is True and must_exist and not os.path.isdir(path):
        raise UserError(f"Not a folder:\n{path}")
    if want_dir is False and must_exist and not os.path.isfile(path):
        raise UserError(f"Not a file:\n{path}")
    return path


def resolve_output(out_dir, name, ext, overwrite=False):
    """Folder + file name chosen in the page -> full path (like asksaveasfilename)."""
    check_path(out_dir, want_dir=True, writable=True)
    name = os.path.basename((name or "").strip())
    if not name:
        raise UserError("Please type an output file name.")
    if ext and not name.lower().endswith(ext):
        name += ext
    full = os.path.join(out_dir, name)
    if os.path.exists(full) and not overwrite:
        abort(409, description=f"'{name}' already exists in that folder. Replace it?")
    return full


@app.get("/api/ping")
def ping():
    return jsonify({"ok": True, "key_required": bool(CONFIG["access_key"])})


@app.get("/api/roots")
def roots():
    return jsonify([
        {"name": r["name"], "path": r["path"], "writable": bool(r.get("writable")),
         "reachable": os.path.isdir(r["path"])}
        for r in CONFIG["roots"]
    ])


@app.get("/api/browse")
def browse():
    path = request.args.get("path", "")
    check_path(path, want_dir=True)
    exts = [e.strip().lower() for e in request.args.get("ext", "").split(",") if e.strip()]
    dirs, files = [], []
    try:
        with os.scandir(path) as it:
            for entry in it:
                try:
                    if entry.name.startswith("~$"):
                        continue  # Office lock files
                    st = entry.stat()
                    if entry.is_dir():
                        dirs.append({"name": entry.name, "path": entry.path, "mtime": st.st_mtime})
                    elif not exts or os.path.splitext(entry.name)[1].lower() in exts:
                        files.append({"name": entry.name, "path": entry.path,
                                      "size": st.st_size, "mtime": st.st_mtime})
                except OSError:
                    continue
    except OSError as e:
        raise UserError(f"Could not open folder: {e}")
    dirs.sort(key=lambda d: d["name"].lower())
    files.sort(key=lambda f: f["name"].lower())
    root = _root_for(path)
    parent = os.path.dirname(path.rstrip("\\/"))
    at_root = _norm(path) == root["_norm"]
    return jsonify({
        "path": path, "root": root["name"], "writable": bool(root.get("writable")),
        "parent": None if at_root else parent, "dirs": dirs, "files": files,
    })


@app.get("/api/download")
def download():
    path = request.args.get("path", "")
    check_path(path, want_dir=False)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.get("/api/logo")
def logo():
    p = CONFIG.get("logo_path") or ""
    if p and os.path.isfile(p):
        return send_file(p)
    abort(404)


# ---------------------------------------------------------------------------
# Background jobs
# ---------------------------------------------------------------------------
JOBS = OrderedDict()
JOBS_LOCK = threading.Lock()


def start_job(tool, fn, **kwargs):
    job_id = uuid.uuid4().hex[:12]
    job = {"id": job_id, "tool": tool, "status": "running", "log": [], "result": None,
           "outputs": [], "error": None, "started": time.time(), "client": request.remote_addr}

    def log(msg):
        for line in str(msg).split("\n"):
            job["log"].append(line)

    def worker():
        try:
            res = fn(log=log, **kwargs)
            if isinstance(res, list):
                job["outputs"] = [{"name": os.path.basename(p), "path": p} for p in res]
            else:
                job["result"] = res
            job["status"] = "done"
        except UserError as e:
            job["error"] = str(e)
            job["status"] = "error"
        except Exception as e:
            traceback.print_exc()
            job["error"] = f"{type(e).__name__}: {e}"
            log(traceback.format_exc())
            job["status"] = "error"
        job["finished"] = time.time()

    with JOBS_LOCK:
        JOBS[job_id] = job
        while len(JOBS) > 300:
            JOBS.popitem(last=False)
    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.get("/api/job/<job_id>")
def job_status(job_id):
    job = JOBS.get(job_id)
    if not job:
        abort(404, description="Job not found (the server may have restarted).")
    since = int(request.args.get("since", 0))
    return jsonify({
        "status": job["status"], "log": job["log"][since:], "next": len(job["log"]),
        "result": job["result"], "outputs": job["outputs"], "error": job["error"],
    })


# ---------------------------------------------------------------------------
# Shared dataframe cache (parquet files are slow to pull over the share)
# ---------------------------------------------------------------------------
_DF_CACHE = OrderedDict()
_DF_LOCK = threading.Lock()


def cached_df(kind, path, loader):
    check_path(path, want_dir=False)
    key = (kind, _norm(path), os.path.getmtime(path))
    with _DF_LOCK:
        if key in _DF_CACHE:
            _DF_CACHE.move_to_end(key)
            return _DF_CACHE[key]
    df = loader(path)
    with _DF_LOCK:
        _DF_CACHE[key] = df
        while len(_DF_CACHE) > 6:
            _DF_CACHE.popitem(last=False)
    return df


def body():
    return request.get_json(force=True, silent=True) or {}


# ---------------------------------------------------------------------------
# Tool 1 - Aggregate Control Files
# ---------------------------------------------------------------------------
@app.post("/api/aggregate/run")
def aggregate_run():
    b = body()
    files = b.get("files") or []
    if not files:
        raise UserError("Select at least one Control File.")
    for f in files:
        check_path(f, want_dir=False)
    out = resolve_output(b.get("output_dir"), b.get("output_name"), ".xlsx", b.get("overwrite"))
    return start_job("aggregate", tool_aggregate_cf.run, file_paths=files, output_file=out)


# ---------------------------------------------------------------------------
# Tool 2 - Build Master Parquet
# ---------------------------------------------------------------------------
@app.post("/api/master/run")
def master_run():
    b = body()
    for k, label in [("aggregated", "CF_aggregated.parquet"), ("tracker", "Project Tracker.parquet"),
                     ("misc", "miscelaneous.parquet")]:
        if not b.get(k):
            raise UserError(f"Select the {label} file.")
        check_path(b[k], want_dir=False)
    out = resolve_output(b.get("output_dir"), b.get("output_name"), ".parquet", b.get("overwrite"))
    return start_job("master", tool_build_master.run,
                     aggregated_file=b["aggregated"], tracker_file=b["tracker"], misc_file=b["misc"],
                     output_parquet_file=out, include_man_day=bool(b.get("include_man_day", True)))


# ---------------------------------------------------------------------------
# Tool 3 - Target Price -> Control File
# ---------------------------------------------------------------------------
@app.post("/api/tp/analyze")
def tp_analyze():
    b = body()
    check_path(b.get("tp"), want_dir=False)
    return start_job("tp-analyze", tool_tp_to_cf.analyze, tp_path=b["tp"])


@app.post("/api/tp/run")
def tp_run():
    b = body()
    check_path(b.get("control"), want_dir=False)
    check_path(b.get("tp"), want_dir=False)
    out_dir = b.get("output_dir") or os.path.dirname(b["control"])
    check_path(out_dir, want_dir=True, writable=True)
    return start_job("tp-run", tool_tp_to_cf.run, control_path=b["control"], tp_path=b["tp"],
                     choices=b.get("choices") or [], output_dir=out_dir)


# ---------------------------------------------------------------------------
# Tool 4 - Materials / CV Excel report
# ---------------------------------------------------------------------------
def _materials_df(path):
    return cached_df("materials", path, tool_materials_report.load_file)


def _materials_filtered(b):
    df = _materials_df(b.get("path"))
    try:
        return tool_materials_report.apply_filters(
            df, b.get("filters") or {}, b.get("dates") or [], b.get("start", ""), b.get("end", ""))
    except (ValueError, TypeError) as e:
        raise UserError(f"Could not apply the date filter: {e}")


@app.post("/api/materials/options")
def materials_options():
    b = body()
    return jsonify(tool_materials_report.filter_options(_materials_df(b.get("path"))))


@app.post("/api/materials/apply")
def materials_apply():
    df, _ = _materials_filtered(body())
    return jsonify({"rows": int(len(df))})


@app.post("/api/materials/export")
def materials_export():
    b = body()
    df, date_filters = _materials_filtered(b)
    out = resolve_output(b.get("output_dir"), b.get("output_name"), ".xlsx", b.get("overwrite"))
    return start_job("materials", tool_materials_report.export, df_filtered=df,
                     selected_date_filters=date_filters,
                     selected_cv_groups=b.get("cv_groups") or [], output_path=out)


# ---------------------------------------------------------------------------
# Tool 5 - Work Instructions & map check
# ---------------------------------------------------------------------------
WI = tool_work_instructions


def _wi_df(path):
    try:
        return cached_df("wi", path, WI.load_master_parquet)
    except ValueError as e:
        raise UserError(str(e))


def _wi_filtered(b):
    df = _wi_df(b.get("path"))
    try:
        return df, WI.filtered(df, b.get("filters") or {})
    except ValueError as e:
        raise UserError(str(e))


def _entries_out(entries):
    out = []
    for e in entries:
        e = dict(e)
        for k in ("outage_start", "outage_end"):
            if isinstance(e.get(k), datetime):
                e[k] = e[k].strftime("%Y-%m-%d")
        out.append(e)
    return out


def _entries_in(entries):
    out = []
    for e in entries or []:
        e = dict(e)
        check_path(e.get("path"), want_dir=False)
        for k in ("outage_start", "outage_end"):
            if e.get(k):
                e[k] = datetime.strptime(e[k][:10], "%Y-%m-%d")
            else:
                e.pop(k, None)
        out.append(e)
    return out


@app.post("/api/wi/options")
def wi_options():
    return jsonify(WI.filter_options(_wi_df(body().get("path"))))


@app.post("/api/wi/preview")
def wi_preview():
    _full, df = _wi_filtered(body())
    return jsonify({"count": int(len(df)), **WI.preview_rows(df)})


@app.post("/api/wi/add_folder")
def wi_add_folder():
    folder = body().get("folder")
    check_path(folder, want_dir=True)
    try:
        return jsonify({"entries": WI.entries_from_folder(folder)})
    except FileNotFoundError as e:
        raise UserError(str(e))


@app.post("/api/wi/add_files")
def wi_add_files():
    files = body().get("files") or []
    for f in files:
        check_path(f, want_dir=False)
    return jsonify({"entries": WI.entries_from_files(files)})


@app.post("/api/wi/scan")
def wi_scan():
    b = body()
    root = (b.get("root") or "").strip()
    check_path(root, want_dir=True)
    try:
        date_from = WI.parse_date(b.get("date_from"))
        date_to = WI.parse_date(b.get("date_to"))
    except ValueError as e:
        raise UserError(str(e))

    def scan(log):
        res = WI.scan_network(root, date_from=date_from, date_to=date_to, log=log)
        res["entries"] = _entries_out(res["entries"])
        return res

    return start_job("wi-scan", scan)


@app.post("/api/wi/word")
def wi_word():
    b = body()
    _full, df = _wi_filtered(b)
    if len(df) == 0:
        raise UserError("There are no rows to write - check the filters on step 1.")
    out = resolve_output(b.get("output_dir"), b.get("output_name") or WI.DEFAULT_DOCX_NAME,
                         ".docx", b.get("overwrite"))
    return start_job("wi-word", WI.generate_word, filtered_df=df, save_path=out,
                     pdf_entries=_entries_in(b.get("pdf_entries")),
                     use_map_order=bool(b.get("use_map_order", True)))


@app.post("/api/wi/report")
def wi_report():
    b = body()
    full, df = _wi_filtered(b)
    entries = _entries_in(b.get("pdf_entries"))
    if len(df) == 0:
        raise UserError("Load a parquet file and apply filters first (no rows match).")
    if not entries:
        raise UserError("Add at least one folder or PDF file first, or use the network scan.")
    try:
        threshold = float(b.get("threshold", WI.DEFAULT_SIMILARITY_THRESHOLD))
    except (TypeError, ValueError):
        raise UserError("Threshold must be a number, e.g. 85")
    out = resolve_output(b.get("output_dir"), b.get("output_name") or WI.DEFAULT_REPORT_NAME,
                         ".xlsx", b.get("overwrite"))
    return start_job("wi-report", WI.generate_report, full_df=full, filtered_df=df,
                     pdf_entries=entries, save_path=out, threshold=threshold,
                     network_root=(b.get("network_root") or "").strip(),
                     drive_letter=(b.get("drive_letter") or "").strip())


@app.get("/api/defaults")
def defaults():
    return jsonify({"network_root": WI.NETWORK_ROOT,
                    "similarity_threshold": WI.DEFAULT_SIMILARITY_THRESHOLD})


# ---------------------------------------------------------------------------
# The page itself
# ---------------------------------------------------------------------------
@app.get("/")
def index():
    return send_from_directory(HERE, "index.html")


if __name__ == "__main__":
    host, port = CONFIG["host"], int(CONFIG["port"])
    print(f"Gaeltec Tools running on http://{host}:{port}  (allowed: {', '.join(CONFIG['allowed_clients'])})")
    for r in CONFIG["roots"]:
        print(f"  root '{r['name']}': {r['path']}  {'OK' if os.path.isdir(r['path']) else 'NOT REACHABLE'}")
    try:
        from waitress import serve
        serve(app, host=host, port=port, threads=16)
    except ImportError:
        app.run(host=host, port=port, threaded=True)
