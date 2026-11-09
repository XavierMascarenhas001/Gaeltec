"""
patch_generators.py - adds the "Generators ⇄ High-level planning" tool to the page.

  * a tile on the home page, right after Work Instructions & Map Check
    (in the 3-across grid that puts it directly below Build Master File),
    with a generator icon
  * the tool page: drop the Generator Tracker + High-level planning, password
    if protected, task 1 / 2, dates, tick days -> outages, preview with
    keep / replace for anything already filled in, apply & download
  * bridge.py: the gs_* functions the page calls
  * pylibs.zip: msoffcrypto + olefile (to open password-protected workbooks;
    the browser's own 'cryptography' package is loaded only when needed)

Put this next to page_template.html, bridge.py, pylibs.zip and
tool_generator_sync.py and run it once:
    python patch_generators.py
then rebuild as usual:
    python build_page.py
Running it again does nothing (each part checks whether it's already there).
"""
import importlib
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))

ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 7V4h12v3"/><rect x="2.5" y="7" width="19" height="11" rx="2"/><path d="m12.8 9.2-2.6 3.3h3.6l-2.6 3.3"/><circle cx="6.5" cy="20.5" r="1.5"/><circle cx="17.5" cy="20.5" r="1.5"/><path d="M5.5 10.5v4"/><path d="M18.5 10.5v4"/></svg>'

TILE = """  {id:"gens", n:5, c:"c6", icon:"generator", tag:"Planning", noflow:true, title:"Generators ⇄ High-level planning",
   desc:"Pass outages from the High-level planning to the Generator Tracker, or send the generators (status, provider, phase, kVA, isolation point, cable) and battery pods back to the planning. Pick days and outages; nothing already filled in is overwritten without asking."},
"""

PAGE_JS = r'''/* ---------------- Generators ⇄ High-level planning ---------------- */
const NEED_GS_CRYPT = ["cryptography"];   // only loaded when a workbook is password protected
function pageGen(t){
  let loaded = false, needCrypt = false, preview = null, choiceEls = new Map(), days = [], run = null;

  /* 1 - files + password */
  const b1 = box(1, "Drop the two workbooks", "If a workbook is password protected, type its password and click <b>Read files</b>. The password isn't stored anywhere.");
  const xl = [["Excel files", ".xlsx,.xlsm"]];
  const onFiles = () => { loaded = false; resetDays(); readFiles(); };
  const trk = dropField({label: "Generator Tracker", kind: "file", types: xl, onChange: onFiles});
  const pln = dropField({label: "High-level planning", kind: "file", types: xl, onChange: onFiles});
  const r1 = h(`<div class="row"></div>`); r1.append(trk, pln); b1.append(r1);
  const pw = h(`<div class="row" style="margin-top:12px;align-items:flex-end"><div class="field"><label>Password (if protected)</label><input type="password" autocomplete="off" placeholder="file password" style="width:220px"></div><button class="btn">${I.folder} Read files</button></div>`);
  const pwIn = $("input", pw);
  const fStat = h(`<div class="status" style="margin-top:10px;white-space:pre-line"></div>`);
  b1.append(pw, fStat);
  $(".btn", pw).onclick = () => readFiles(true);
  pwIn.addEventListener("keydown", e => { if (e.key === "Enter") readFiles(true); });

  async function readFiles(explicit){
    if (!trk.get() || !pln.get()){ fStat.className = "status"; fStat.textContent = "Drop both workbooks."; return; }
    fStat.className = "status"; fStat.textContent = PY.ready ? "Reading the workbooks…" : "Starting Python and reading the workbooks (first time ~30 s)…";
    try {
      const p = (await PY.call("gs_probe", {tracker: trk.get(), planning: pln.get()})).result;
      needCrypt = p.tracker.encrypted || p.planning.encrypted;
      if (needCrypt && !pwIn.value){
        const which = [p.tracker.encrypted && "Generator Tracker", p.planning.encrypted && "High-level planning"].filter(Boolean).join(" and ");
        fStat.className = "status err"; fStat.textContent = `🔒 The ${which} is password protected – type the password and click Read files.`;
        pwIn.focus(); return;
      }
      const s = (await PY.call("gs_load", {tracker: trk.get(), planning: pln.get(), password: pwIn.value},
                               {needs: needCrypt ? NEED_GS_CRYPT : []})).result;
      loaded = true;
      fStat.className = "status ok"; fStat.textContent = `✔ ${s.tracker}\n✔ ${s.planning}`;
    } catch(e){
      loaded = false;
      fStat.className = "status err";
      fStat.textContent = /cryptography|loadPackage|Failed to fetch/i.test(e.message)
        ? "Couldn't load the decryption package (cryptography) – check the internet connection, or remove the password from the file and drop it again."
        : e.message;
    }
  }

  /* 2 - task + dates */
  const b2 = box(2, "Choose the task and the dates");
  const tasks = h(`<div class="seg" style="margin-bottom:14px">
      <label><input type="radio" name="gstask" value="1" checked><span>1 · Planning → Generator Tracker</span></label>
      <label><input type="radio" name="gstask" value="2"><span>2 · Generator Tracker → Planning</span></label></div>`);
  const taskHint = h(`<p class="hint" style="margin:0 0 12px"></p>`);
  const task = () => +$("input[name=gstask]:checked", tasks).value;
  const setHint = () => { taskHint.innerHTML = task() === 1
    ? "Adds <b>District, Date, Weekday, Outage, Circuit and PM</b> of the chosen outages to the first free rows of the Generators table."
    : "Writes the <b>Generator</b> rows into the planning <b>Generators</b> column (e.g. <i>No Access, (x1) Generator from GWF, 3 phase 10 kVa, Setubal, 30 m of cable needed.</i>) and the <b>Battery pod</b> multipliers into <b>Battery Pods</b>."; };
  setHint();
  tasks.querySelectorAll("input").forEach(i => i.onchange = () => { setHint(); resetDays(); });
  const today = new Date(), iso = d => new Date(d.getTime() - d.getTimezoneOffset() * 60e3).toISOString().slice(0, 10);
  const dr = h(`<div class="row" style="align-items:flex-end"><div class="field"><label>From</label><input type="date"></div><div class="field"><label>To</label><input type="date"></div><button class="btn primary">${I.cal} Show days &amp; outages</button><span class="status"></span></div>`);
  const [dFrom, dTo] = dr.querySelectorAll("input");
  dFrom.value = iso(today); dTo.value = iso(new Date(today.getTime() + 7 * 864e5));
  [dFrom, dTo].forEach(i => i.onchange = resetDays);
  const dStat = $(".status", dr);
  b2.append(tasks, taskHint, dr);

  /* 3 - days and outages */
  const b3 = box(3, "Tick the days, then the outages", "Tick a day to see its outages – they all start ticked, untick any you don't want.");
  const dBtns = h(`<div class="row" style="margin-bottom:10px"><button class="btn sm">Tick all days</button><button class="btn sm">Clear all</button><span class="status"></span></div>`);
  const [bAll, bNone] = dBtns.querySelectorAll(".btn"); const selStat = $(".status", dBtns);
  const list = h(`<div style="max-height:460px;overflow:auto;border:1px solid var(--line,#ddd);border-radius:10px;padding:8px 12px"></div>`);
  b3.append(dBtns, list);

  function resetDays(){
    days = []; list.innerHTML = `<p class="hint" style="margin:6px 0">Click <b>Show days &amp; outages</b>.</p>`;
    selStat.textContent = ""; clearPreview();
  }
  function countSel(){
    const n = ids().length; selStat.className = "status"; selStat.textContent = n ? `${n} outage(s) selected` : "";
    clearPreview();
  }
  function ids(){
    return days.flatMap(d => d.cb.checked ? d.items.filter(x => x.cb.checked).map(x => x.id) : []);
  }
  $(".btn", dr).onclick = async () => {
    clearPreview();
    if (!loaded){ dStat.className = "status err"; dStat.textContent = "Read both workbooks first (step 1)."; return; }
    dStat.className = "status"; dStat.textContent = "Listing outages…";
    try {
      const r = (await PY.call("gs_days", {task: task(), date_from: dFrom.value, date_to: dTo.value})).result;
      list.innerHTML = ""; days = [];
      if (!r.days.length){
        list.innerHTML = `<p class="hint" style="margin:6px 0">No outages in this date range${task() === 2 ? " with Type of supply = Generator / Battery pod" : ""}.</p>`;
        dStat.textContent = ""; return;
      }
      r.days.forEach(d => {
        const day = h(`<div style="padding:4px 0;border-bottom:1px solid var(--line,#eee)">
            <label style="display:flex;gap:8px;align-items:center;font-weight:700;cursor:pointer"><input type="checkbox"><span></span></label>
            <div class="hidden" style="padding:4px 0 6px 28px"></div></div>`);
        $("span", day).textContent = `${d.label}  (${d.items.length} outage${d.items.length > 1 ? "s" : ""})`;
        const dayCb = $("input", day), inner = $("div", day);
        const entry = {cb: dayCb, items: []};
        d.items.forEach(it => {
          const row = h(`<label style="display:flex;gap:8px;align-items:center;padding:2px 0;cursor:pointer;font-size:13px"><input type="checkbox" checked><span></span></label>`);
          $("span", row).textContent = it.label;
          const cb = $("input", row); cb.onchange = countSel;
          entry.items.push({id: it.id, cb}); inner.append(row);
        });
        dayCb.onchange = () => { inner.classList.toggle("hidden", !dayCb.checked); if (dayCb.checked) entry.items.forEach(x => x.cb.checked = true); countSel(); };
        days.push(entry); list.append(day);
      });
      dStat.className = "status ok"; dStat.textContent = `${r.days.length} day(s)`;
    } catch(e){ dStat.className = "status err"; dStat.textContent = e.message; }
  };
  bAll.onclick = () => { days.forEach(d => { if (!d.cb.checked){ d.cb.checked = true; d.cb.onchange(); } }); };
  bNone.onclick = () => { days.forEach(d => { if (d.cb.checked){ d.cb.checked = false; d.cb.onchange(); } }); };

  /* 4 - preview + decisions */
  const b4 = box(4, "Preview", "Shows exactly what will be written. Anything that already has information in the file is listed below the table – choose to keep it or replace it.");
  const pv = h(`<div class="runbar"><button class="btn">${I.search} Preview</button><span class="status"></span></div>`);
  const pvStat = $(".status", pv);
  const pvOut = h(`<div style="margin-top:12px"></div>`);
  const conf = h(`<div style="margin-top:12px"></div>`);
  b4.append(pv, pvOut, conf);
  function clearPreview(){ preview = null; choiceEls = new Map(); pvOut.innerHTML = ""; conf.innerHTML = ""; pvStat.textContent = ""; if (run) run.button.disabled = true; }
  $(".btn", pv).onclick = async () => {
    clearPreview();
    const sel = ids();
    if (!days.length){ pvStat.className = "status err"; pvStat.textContent = "Click Show days & outages first (step 2)."; return; }
    if (!sel.length){ pvStat.className = "status err"; pvStat.textContent = "Tick at least one day (and outage)."; return; }
    pvStat.className = "status"; pvStat.textContent = "Working out what will change…";
    try {
      const r = (await PY.call("gs_preview", {task: task(), ids: sel})).result;
      preview = r;
      pvStat.className = "status ok"; pvStat.textContent = r.summary;
      pvOut.append(tableEl(r.columns, r.rows));
      if (r.conflicts.length){
        const opts = `<option value="">— choose —</option>` + r.choices.map(c => `<option value="${esc(c.value)}">${esc(c.label)}</option>`).join("");
        const head = h(`<div class="row" style="align-items:center;margin:4px 0 10px"><b style="color:#b26a00">${r.conflicts.length} item(s) already have information in the file – what do you want to do?</b>
            <div class="field"><label>Set all to</label><select style="min-width:260px">${opts}</select></div></div>`);
        const all = $("select", head);
        conf.append(head);
        r.conflicts.forEach(c => {
          const card = h(`<div style="display:flex;gap:14px;align-items:center;border-left:3px solid #e69500;background:rgba(230,149,0,.08);border-radius:6px;padding:8px 10px;margin:6px 0">
              <div style="flex:1;font-size:13px;line-height:1.45"><b></b><div class="cur"></div><div class="new"></div></div><select style="min-width:260px">${opts}</select></div>`);
          $("b", card).textContent = c.where;
          $(".cur", card).innerHTML = "Current: " + esc(c.current).replace(/\n/g, "<br>");
          $(".new", card).innerHTML = "New: " + esc(c.new).replace(/\n/g, "<br>");
          const s = $("select", card); choiceEls.set(c.i, s);
          conf.append(card);
        });
        all.onchange = () => { if (all.value) choiceEls.forEach(s => s.value = all.value); };
      }
      run.button.disabled = !r.can_apply;
    } catch(e){ pvStat.className = "status err"; pvStat.textContent = e.message; }
  };

  /* 5 - apply */
  const b5 = box(5, "Apply & download", "Downloads an updated copy (<i>name</i>_updated_<i>date</i>.xlsx) – the original file isn't touched. Pivots, charts and links in the workbook are kept, and a protected file stays protected with the same password.");
  run = runArea("Apply & download", "save");
  run.button.disabled = true;
  b5.append(run);
  run.button.onclick = async () => {
    if (!preview){ run.fail("Click Preview first (step 4)."); return; }
    const choices = {};
    let missing = 0;
    choiceEls.forEach((s, i) => { if (s.value) choices[i] = s.value; else missing++; });
    if (missing){ run.fail(`Choose keep / replace for the ${missing} item(s) that already have information (or use 'Set all to').`); return; }
    const r = await run.run("gs_apply", {choices}, {needs: needCrypt ? NEED_GS_CRYPT : [], busy: "Writing the workbook…"});
    if (!r) return;
    preview = null; choiceEls = new Map(); conf.innerHTML = ""; pvOut.innerHTML = ""; run.button.disabled = true;
    pvStat.className = "status ok"; pvStat.textContent = "Applied – click Show days & outages / Preview again for another run (it continues from the updated copy).";
    try { const s = (await PY.call("gs_status", {})).result; fStat.textContent = `✔ ${s.tracker}\n✔ ${s.planning}`; } catch(e){}
  };

  resetDays();
  app.append(b1, b2, b3, b4, b5);
}'''

BRIDGE = r'''

# ---------------------------------------------------------------------------
# Generators <-> High-level planning (tool_generator_sync.py - the same file
# that runs in Jupyter). The page calls these one after the other:
#   gs_probe -> gs_load -> gs_days -> gs_preview -> gs_apply
# The two workbooks stay decrypted in memory between calls; after an Apply
# the updated copy is kept, so a second run continues from it (dropping a
# new file replaces it).
# ---------------------------------------------------------------------------
_GS = {}


def _gs_is_encrypted(path):
    with open(path, "rb") as f:
        return f.read(4) != b"PK\x03\x04"


def gs_probe(tracker="", planning="", log=print):
    """Which of the dropped workbooks are password protected (the page then
    loads the decryption package only when it's actually needed)."""
    out = {}
    for key, path in (("tracker", tracker), ("planning", planning)):
        out[key] = {"encrypted": bool(path and os.path.exists(path) and _gs_is_encrypted(path))}
    return out


def _gs_status():
    import tool_generator_sync as G
    s = {}
    t = _GS.get("trk")
    if t:
        s["tracker"] = (f"{_GS['trk_name']}{' (protected)' if _GS['trk_enc'] else ''} – {len(t['rows'])} rows, "
                        f"next free row {t['last_used_row'] + 1}")
    p = _GS.get("plan")
    if p:
        s["planning"] = (f"{_GS['plan_name']}{' (protected)' if _GS['plan_enc'] else ''} – {len(p['rows'])} outages "
                         f"on sheets {', '.join(p['sheets'])}")
    return s


def gs_load(tracker="", planning="", password="", log=print):
    import tool_generator_sync as G
    _need(tracker, "Generator Tracker")
    _need(planning, "High-level planning workbook")
    for key, path, reader in (("trk", tracker, G.read_tracker), ("plan", planning, G.read_planning)):
        stamp = (path, os.path.getmtime(path), password if _gs_is_encrypted(path) else "")
        if _GS.get(key + "_stamp") == stamp:
            continue  # same file (and password) as last time - keep it, incl. any applied changes
        label = "Generator Tracker" if key == "trk" else "High-level planning"
        raw = open(path, "rb").read()
        if raw[:4] != b"PK\x03\x04" and not password:
            raise UserError(f"The {label} ({os.path.basename(path)}) is password protected – type the password and click Read files.")
        try:
            plain, enc = G.decrypt_if_needed(raw, password)
            data = reader(plain)
        except RuntimeError as e:
            raise UserError(f"{label}: {e}")
        log(f"Read {label}: {os.path.basename(path)}")
        _GS.update({key: data, key + "_plain": plain, key + "_enc": enc, key + "_name": os.path.basename(path),
                    key + "_stamp": stamp})
    _GS["password"] = password
    _GS.pop("plan_res", None)
    return _gs_status()


def _gs_ready():
    if not (_GS.get("trk") and _GS.get("plan")):
        raise UserError("Drop both workbooks and click Read files first.")


def gs_days(task=1, date_from="", date_to="", log=print):
    import tool_generator_sync as G
    _gs_ready()
    try:
        a = datetime.strptime(date_from[:10], "%Y-%m-%d").date()
        b = datetime.strptime(date_to[:10], "%Y-%m-%d").date()
    except ValueError:
        raise UserError("Choose both dates (From and To).")
    a, b = min(a, b), max(a, b)
    task = int(task)
    cands = G.candidates_task1(_GS["plan"], a, b) if task == 1 else G.candidates_task2(_GS["trk"], a, b)
    _GS["cands"] = {c["_id"]: c for c in cands}
    _GS["task"] = task
    _GS.pop("plan_res", None)
    days = {}
    for c in cands:
        d = days.setdefault(c["_date"], {"date": c["_date"].isoformat(), "label": c["_date"].strftime("%a %d %b %Y"), "items": []})
        d["items"].append({"id": c["_id"], "label": c["_label"]})
    return {"days": list(days.values())}


def gs_preview(task=1, ids=None, log=print):
    import tool_generator_sync as G
    _gs_ready()
    if int(task) != _GS.get("task"):
        raise UserError("The task was changed – click Show days & outages again.")
    items = [_GS["cands"][i] for i in ids or [] if i in _GS.get("cands", {})]
    if not items:
        raise UserError("Tick at least one day (and outage).")
    res = G.plan_task1(items, _GS["trk"]) if _GS["task"] == 1 else G.plan_task2(items, _GS["plan"])
    _GS["plan_res"] = res
    df = G.preview_frame(res)
    df = df.astype(object).where(df.notna(), "")
    conflicts = []
    for i, it in enumerate(res["items"]):
        if not it["conflict"]:
            continue
        if res["task"] == 1:
            conflicts.append({"i": i, "where": it["where"], "current": it["current"], "new": it["new"]})
        else:
            conflicts.append({"i": i, "where": f"{it['sheet']}!{it['col']}{it['row']} ({it['field']}) · {it['date']} "
                                                f"{it['outage']} {it['circuit']}".strip(),
                              "current": it["current"], "new": it["new"]})
    n_conf = len(conflicts)
    if res["task"] == 1:
        n_new = len(res["items"]) - n_conf
        summary = (f"{n_new} outage(s) will be added to the Generators table starting at row "
                   f"{_GS['trk']['last_used_row'] + 1}; {n_conf} already in the tracker.")
        choices = G.T1_CHOICES
    else:
        n_w = sum(1 for it in res["items"] if not it["conflict"] and not it["same"])
        summary = (f"{n_w} empty cell(s) will be filled; {n_conf} cell(s) already have information; "
                   f"{len(res['unmatched'])} outage(s) could not be matched.")
        choices = G.T2_CHOICES
    return {
        "summary": summary,
        "columns": [str(c) for c in df.columns],
        "rows": [[str(v).replace("\n", "  ⏎  ") for v in r] for r in df.values.tolist()],
        "conflicts": conflicts,
        "choices": [{"label": l, "value": v} for l, v in choices],
        "can_apply": any(not it.get("same") for it in res["items"]),
    }


def gs_apply(choices=None, log=print):
    import tool_generator_sync as G
    res = _GS.get("plan_res")
    if not res:
        raise UserError("Click Preview first.")
    choices = {int(k): v for k, v in (choices or {}).items() if v}
    undecided = [it for i, it in enumerate(res["items"]) if it["conflict"] and i not in choices]
    if undecided:
        raise UserError(f"Choose what to do for the {len(undecided)} item(s) that already have information "
                        f"(or use 'Set all to').")
    pwd = _GS.get("password", "")
    if res["task"] == 1:
        edits, summary = G.build_edits_task1(res, choices, _GS["trk"])
        if not edits:
            raise UserError("Nothing to write – everything was kept as it is.")
        plain, wlog = G.apply_task1(_GS["trk_plain"], edits)
        final = G.encrypt(plain, pwd) if _GS["trk_enc"] else plain
        name = _GS["trk_name"]
        _GS.update(trk_plain=plain, trk=G.read_tracker(plain))
    else:
        edits, summary = G.build_edits_task2(res, choices)
        if not edits:
            raise UserError("Nothing to write – everything was kept as it is.")
        plain, wlog = G.apply_task2(_GS["plan_plain"], edits)
        final = G.encrypt(plain, pwd) if _GS["plan_enc"] else plain
        name = _GS["plan_name"]
        _GS.update(plan_plain=plain, plan=G.read_planning(plain))
    _GS.pop("plan_res", None)
    for line in wlog:
        log("• " + line)
    for s in summary:
        log(" · ".join(f"{v}" for v in s.values()).replace("\n", "  ⏎  "))
    stem, ext = os.path.splitext(name)
    out = os.path.join(_outdir(), f"{stem}_updated_{datetime.now():%Y-%m-%d_%H%M}{ext or '.xlsx'}")
    with open(out, "wb") as f:
        f.write(final)
    return [out]


def gs_status(log=print):
    return _gs_status()
'''

# (description, old text, new text) - each old text must be found exactly once
TEMPLATE_EDITS = [
    ("Tile colour",
     ".c5{background:linear-gradient(135deg,#B03A48,#E4606D)}",
     ".c5{background:linear-gradient(135deg,#B03A48,#E4606D)}\n.c6{background:linear-gradient(135deg,#1E6B52,#3DAA6E)}"),
    ("Generator icon",
     "\nfunction box(n, title, hint){",
     "\nI.generator = '" + ICON + "';\n\nfunction box(n, title, hint){"),
    ("Keep the tool out of the 'Usual order' line",
     "Usual order: ${TOOLS.map(t=>",
     "Usual order: ${TOOLS.filter(t=>!t.noflow).map(t=>"),
    ("Tile tag",
     '<span class="tag">Step ${t.n}</span>',
     '<span class="tag">${t.tag ? esc(t.tag) : "Step " + t.n}</span>'),
    ("Tool page",
     "/* ---------------- router ---------------- */",
     PAGE_JS + "\n\n/* ---------------- router ---------------- */"),
    ("Route",
     "const PAGES = {",
     "const PAGES = {gens: pageGen, "),
]


def patch_template(text):
    if "function pageGen(" in text:
        return text, False
    for desc, old, new in TEMPLATE_EDITS:
        if text.count(old) != 1:
            raise SystemExit(f"page_template.html: couldn't find '{desc}' exactly as expected - nothing was changed.")
        text = text.replace(old, new)
    # the tile goes last in TOOLS = right after Work Instructions & Map Check
    m = re.search(r"const TOOLS = \[\n(.*?)\n\];", text, re.S)
    if not m or 'id:"wi"' not in m.group(1):
        raise SystemExit("page_template.html: couldn't find the TOOLS list - nothing was changed.")
    text = text[:m.end(1)] + "\n" + TILE.rstrip("\n") + text[m.end(1):]
    return text, True


def patch_bridge(text):
    if "def gs_probe(" in text:
        return text, False
    return text.rstrip("\n") + "\n" + BRIDGE, True


def _package_files(modname):
    """(path on disk, name inside the zip) for an installed pure-Python package."""
    mod = importlib.import_module(modname)
    path = mod.__file__
    if os.path.basename(path) == "__init__.py":
        root = os.path.dirname(path)
        base = os.path.dirname(root)
        for d, dirs, files in os.walk(root):
            dirs[:] = [x for x in dirs if x != "__pycache__"]
            for f in files:
                if not f.endswith((".pyc", ".pyo")):
                    full = os.path.join(d, f)
                    yield full, os.path.relpath(full, base).replace(os.sep, "/")
    else:
        yield path, os.path.basename(path)


def patch_pylibs(path):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
    have = {n.split("/")[0] for n in names}
    missing = [m for m in ("msoffcrypto", "olefile") if m not in have and m + ".py" not in have]
    if not missing:
        return False
    try:
        files = [f for m in missing for f in _package_files(m)]
    except ImportError:
        raise SystemExit("pylibs.zip needs msoffcrypto and olefile, but they aren't installed in this Python.\n"
                         "Run:  pip install msoffcrypto-tool   and then run this script again.\n"
                         "(page_template.html and bridge.py were already patched - that's fine, they won't be patched twice.)")
    with zipfile.ZipFile(path, "a", zipfile.ZIP_DEFLATED) as z:
        for full, arc in files:
            if arc not in names:
                z.write(full, arc)
    return True


if __name__ == "__main__":
    template = os.path.join(HERE, "page_template.html")
    bridge = os.path.join(HERE, "bridge.py")
    pylibs = os.path.join(HERE, "pylibs.zip")
    tool = os.path.join(HERE, "tool_generator_sync.py")
    for p in (template, bridge, pylibs):
        if not os.path.exists(p):
            raise SystemExit(f"{os.path.basename(p)} not found - put this script next to it.")
    changed = []
    for p, fn in ((template, patch_template), (bridge, patch_bridge)):
        old = open(p, encoding="utf-8").read()
        new, did = fn(old)
        if did:
            open(p, "w", encoding="utf-8").write(new)
            changed.append(os.path.basename(p))
    if patch_pylibs(pylibs):
        changed.append("pylibs.zip (added msoffcrypto + olefile)")
    print("Patched: " + ", ".join(changed) if changed else "The Generators tool is already there - nothing to do.")
    if not os.path.exists(tool):
        print("!! tool_generator_sync.py is missing - put it in this folder too (build_page.py picks up every tool_*.py).")
    print("Now run: python build_page.py")
