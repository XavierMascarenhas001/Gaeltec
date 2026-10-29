"""
patch_dashboard.py - Planning Dashboard (Network Job Tracker) on the page:

  Totals tab
    * Total / Original / Variation for the filters
    * Total value by project and Difference by project - click a project
      to open it in the Variations pivot below
    * Variations - Project > Source file > Job with Orig / Total /
      Variation / Reason (same as the pivot on the Excel Variations sheet)
    * Variations by service code (same as the Variations_service codes sheet)
    * Variations - every item (the rows under the Excel pivot)
  Pole Position tab
    * the PID breakdown from Master Control underneath
  Files strip
    * a "Service codes" slot for miscelaneous.parquet, remembered in this
      browser (shared with the Outputs tool's box)

Put this next to page_template.html, dash_core.js, dash_pages.js and
dashboards.css and run it once:
    python patch_dashboard.py
then rebuild as usual:
    python build_page.py
Running it again does nothing.
"""
import os
import sys

FILES = ["page_template.html", "dash_core.js", "dash_pages.js", "dashboards.css"]

# (description, old text, new text) - each old text is found in exactly one file
EDITS = []


def edit(desc, old, new):
    EDITS.append((desc, old, new))


# ---- 1. a "Service codes" file slot ----
edit("Service codes slot",
     '''  ics:       {label: "Outage calendar", hint: "outage_calendar_2026 (vN).ics", types: [".ics"], icon: "file"},
''',
     '''  ics:       {label: "Outage calendar", hint: "outage_calendar_2026 (vN).ics", types: [".ics"], icon: "file"},
  misc:      {label: "Service codes", hint: "miscelaneous.parquet", types: [".parquet"], icon: "file"},
''')

# ---- 2. remember the miscelaneous file in this browser (same memory as the Outputs tool) ----
edit("Remember the miscelaneous file",
     '''  if (old) PY.unmount(old.dir);
  fileListeners.forEach(cb => cb(slot));
  return true;
}''',
     '''  if (old) PY.unmount(old.dir);
  fileListeners.forEach(cb => cb(slot));
  if (slot === "misc" && file.size < 3e6){   /* remembered, like the logos - shared with the Outputs tool */
    const rd = new FileReader(); rd.onload = () => store.set("outputs_misc", {name: file.name, data: rd.result}); rd.readAsDataURL(file);
  }
  return true;
}''')

# ---- 3. a dropped file that only changes the figures re-runs the view ----
edit("Refresh the view when the Service codes file changes",
     '''    if (def.loadSlots.includes(slot)) P.start(); else if (P.v) P.renderTab();''',
     '''    if (def.loadSlots.includes(slot)) P.start(); else if (P.v && (def.viewSlots || []).includes(slot)) P.refresh(); else if (P.v) P.renderTab();''')

# ---- 4. tracker: slot, about text, defaults, view arguments ----
edit("Tracker slots",
     '''  slots: [["master", true], ["outages", false], ["forecast", false], ["ics", false]], loadSlots: ["master"], needs: NEED_NT,''',
     '''  slots: [["master", true], ["outages", false], ["forecast", false], ["ics", false], ["misc", false]], loadSlots: ["master"], viewSlots: ["misc"], needs: NEED_NT,''')
edit("Tracker about text",
     '''          "Service Partner Workbank → Pole Position.  Outage calendar (.ics) → the events list."],''',
     '''          "Service Partner Workbank → Pole Position.  Outage calendar (.ics) → the events list.",
          "Service codes (miscelaneous.parquet) → the service codes in Totals – remembered in this browser once dropped."],''')
edit("Tracker defaults",
     '''    f.preset = f.preset || "All time"; st.gran = st.gran || "Month"; st.o = st.o || {view: "calendar"}; st.fc = st.fc || {}; },''',
     '''    f.preset = f.preset || "All time"; st.gran = st.gran || "Month"; st.o = st.o || {view: "calendar"}; st.fc = st.fc || {};
    f.pid_mode = f.pid_mode || "All"; restoreMisc(); },''')
edit("Tracker view arguments",
     '''  viewArgs(P){ return {path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P), granularity: P.st.gran, detail: P.st.detail || null}; },''',
     '''  viewArgs(P){ return {path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P), granularity: P.st.gran, detail: P.st.detail || null,
    misc: FILES.misc ? FILES.misc.path : "", pid_mode: P.st.f.pid_mode || "All"}; },''')
edit("Pole Position + PID breakdown",
     '''    {id: "forecast", label: "Pole Position", icon: "chart", render: ntForecast},''',
     '''    {id: "forecast", label: "Pole Position", icon: "chart", render: ntForecastAndPid},''')

# ---- 5. the Totals tab ----
OLD_TOTALS_START = "function ntTotals(P, panel){"
OLD_TOTALS_END = "/* ---- Jobs: Outages Programme calendar + .ics ---- */"
NEW_TOTALS = r'''/* the miscelaneous file dropped earlier (here or in the Outputs tool) is put back in its slot */
function restoreMisc(){
  if (FILES.misc || restoreMisc.tried) return;
  restoreMisc.tried = true;
  const saved = store.get("outputs_misc", null);
  if (!saved || !saved.data) return;
  fetch(saved.data).then(r => r.blob()).then(b => setSlotFile("misc", new File([b], saved.name))).catch(() => {});
}
/* Pole Position, with Master Control's PID breakdown underneath */
function ntForecastAndPid(P, panel){
  const top = h(`<div></div>`), bottom = h(`<div style="margin-top:22px"></div>`);
  panel.append(top, bottom);
  ntForecast(P, top);
  bottom.append(h(`<div class="vsec"><h2>PID breakdown</h2><p>Same as Master Control › PID Breakdown, for the rows under the filters above</p></div>`));
  if (P.v.pid_error){ bottom.append(note(P.v.pid_error, "warn", "warn")); return; }
  P.st.f.pid_mode = P.st.f.pid_mode || "All";
  const pid = P.v.pid || {rows: [], kpis: {}};
  mcPid({v: {pid}, st: P.st, refresh: P.refresh}, bottom);
}
function ntTotals(P, panel){
  const t = P.v.totals;
  if (t.total == null){ panel.append(note("The mapped Total column isn't in the file – pick it under Column mapping.", "warn", "warn")); return; }
  const items = [{label: "Total", value: money(t.total), hero: true}];
  if (t.orig != null){ const d = t.total - t.orig;
    items.push({label: "Original", value: money(t.orig)},
               {label: "Variation", value: signed(d), tone: d >= 0 ? "up" : "down", note: d >= 0 ? "▲ total above the original" : "▼ total below the original"}); }
  const kk = kpis(items); $(".kpi.hero", kk).append(imageEl("Pound", "", "heroimg")); panel.append(kk);

  const vr = t.variations || null;
  const pivot = vr && !vr.error ? ntVarPivot(P, vr) : null;
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  /* click a project's bar -> open it in the Variations pivot */
  const clickable = (card, rows) => {
    if (!pivot) return card;
    card.classList.add("clickrows");
    card.body.addEventListener("click", e => { const hit = e.target.closest(".viz [data-i]"); if (hit) pivot.openProject(String(rows[+hit.dataset.i][0])); });
    return card;
  };
  if (t.by_project){ const r = t.by_project.rows;
    g.append(clickable(vizCard({title: "Total value by project", sub: `${r.length} project(s) · click a project to open its source files below`, span: 6,
      draw: host => Viz.bars(host, {labels: r.map(x => String(x[0])), series: [{name: "Total value", values: r.map(x => x[1]), color: "var(--s1)"}], fmt: money, axis: moneyS, valueLabel: i => moneyS(r[i][1])}),
      table: () => tbl(t.by_project, "Total value by project")}), r)); }
  if (t.variance_by_project){ const r = t.variance_by_project.rows;
    g.append(clickable(vizCard({title: "Difference by project", sub: "Total − original · click a project to open it below", span: 6,
      draw: host => Viz.diverging(host, {labels: r.map(x => String(x[0])), values: r.map(x => x[1]), fmt: money, axis: moneyS, valueLabel: i => signed(r[i][1], moneyS),
        tip: i => ({title: String(r[i][0]), rows: [{color: r[i][1] >= 0 ? "var(--s1)" : "var(--s8)", name: "Difference", value: signed(r[i][1])}]})}),
      table: () => tbl(t.variance_by_project, "Difference by project")}), r)); }

  if (!vr) return;
  if (vr.error){ g.append(note(vr.error, "warn", "warn")); return; }
  g.append(pivot);

  /* Variations by service code - the Variations_service codes sheet */
  const sc = dcard("Variations by service code", "Same as the Excel 'Variations_service codes' sheet: every service code with qty, qsub, Orig, Total and Variation, plus the Traffic Management, Sand / Chips / Miscellaneous Works and Non Free Issue material lines", 12);
  if (vr.service_codes){
    const s = vr.service_codes, sum = i => s.rows.reduce((a, r) => a + (r[i] || 0), 0);
    sc.acts.append(h(`<span class="muted" style="font-size:12.5px">Codes from ${esc(vr.misc || "miscelaneous.parquet")}</span>`));
    sc.body.append(dtable({name: "Variations by service code", columns: s.columns, rows: s.rows, signed: ["Variation (£)"], height: 520}));
    const v = sum(6);
    sc.body.append(h(`<div class="vgt"><b>GRAND TOTAL</b><span>qty ${esc(qty(sum(2)))}</span><span>qsub ${esc(qty(sum(3)))}</span><span>Orig ${esc(money(sum(4)))}</span><span>Total ${esc(money(sum(5)))}</span><span class="${v >= 0 ? "pos" : "neg"}">Variation ${esc(signed(v))}</span></div>`));
  } else {
    sc.body.append(note("Drop miscelaneous.parquet below (or in the Service codes slot at the top) to see the service codes – it's remembered in this browser, so you only drop it once."),
                   slotZone("misc", false, "miscelaneous.parquet – Column_1 = description, Column_3 = service code"));
  }
  g.append(sc);

  /* every item - the rows under the Excel pivot */
  g.append(tableCard("Variations — every item", `${intf(vr.rows)} row(s) · the same rows and columns as under the pivot on the Excel Variations sheet`,
    tbl(vr.detail, "Variations - every item", {height: 460})));
}
/* Project › Source file › Job, like the pivot on the Excel Variations sheet */
function ntVarPivot(P, vr){
  const st = P.st.vp || (P.st.vp = {open: {}});
  const c = dcard("Variations — Project › Source file › Job",
    `Same as the pivot on the Excel Variations sheet (rows where variations ≠ 0) · ${intf(vr.rows)} row(s) · click a project to see its source files, a source file to see its jobs and the reason`, 12);
  const box = h(`<div class="vpiv" role="tree"></div>`);
  const allBtn = h(`<button class="btn sm">Expand all</button>`), noneBtn = h(`<button class="btn sm">Collapse all</button>`);
  c.acts.append(allBtn, noneBtn);
  c.body.append(box);
  const keyP = p => "p\u0001" + p.project, keyS = (p, s) => "s\u0001" + p.project + "\u0001" + s.source;
  const cls = v => (v || 0) > 0 ? "pos" : (v || 0) < 0 ? "neg" : "";
  const row = (lvl, key, label, o, reason, open, kids, extra) =>
    `<div class="vr l${lvl}${open ? " open" : ""}${extra || ""}"${key ? ` data-k="${esc(key)}" role="treeitem" tabindex="0" aria-expanded="${open}"` : ""}>` +
    `<span class="c0">${kids ? `<i class="tw">▸</i>` : ""}<span title="${esc(label)}">${esc(label)}</span></span>` +
    `<span class="vn">${esc(money(o.orig))}</span><span class="vn">${esc(money(o.total))}</span>` +
    `<span class="vn ${cls(o.variation)}">${esc(signed(o.variation))}</span><span class="reason" title="${esc(reason || "")}">${esc(reason || "")}</span></div>`;
  function draw(){
    const parts = [`<div class="vr vh"><span>Job</span><span class="vn">Orig</span><span class="vn">Total</span><span class="vn">Variation</span><span>Reason</span></div>`];
    vr.projects.forEach(p => {
      const po = !!st.open[keyP(p)];
      parts.push(row(0, keyP(p), p.project, p, `${p.sources.length} source file${p.sources.length === 1 ? "" : "s"}`, po, true));
      if (!po) return;
      p.sources.forEach(s => {
        const so = !!st.open[keyS(p, s)];
        parts.push(row(1, keyS(p, s), s.source, s, `${s.jobs.length} job${s.jobs.length === 1 ? "" : "s"}`, so, true));
        if (so) s.jobs.forEach(j => parts.push(row(2, "", j.job, j, j.reason, false, false)));
      });
    });
    parts.push(row(9, "", "GRAND TOTAL", vr.totals, "", false, false, " vt"));
    box.innerHTML = vr.projects.length ? parts.join("") : `<div class="empty">No rows with variations under the current filters.</div>`;
  }
  const toggle = k => { st.open[k] = !st.open[k]; draw(); const el = [...box.querySelectorAll("[data-k]")].find(x => x.dataset.k === k); if (el) el.focus(); };
  box.addEventListener("click", e => { const r = e.target.closest("[data-k]"); if (r) toggle(r.dataset.k); });
  box.addEventListener("keydown", e => { const r = e.target.closest("[data-k]"); if (r && (e.key === "Enter" || e.key === " ")){ e.preventDefault(); toggle(r.dataset.k); } });
  allBtn.onclick = () => { vr.projects.forEach(p => { st.open[keyP(p)] = true; p.sources.forEach(s => { st.open[keyS(p, s)] = true; }); }); draw(); };
  noneBtn.onclick = () => { st.open = {}; draw(); };
  c.openProject = name => {
    const p = vr.projects.find(x => x.project === name);
    if (!p){ box.querySelectorAll(".vmsg").forEach(n => n.remove()); const m = h(`<div class="vmsg"></div>`); m.textContent = `“${name}” has no rows with variations under the current filters.`; box.prepend(m); c.scrollIntoView({behavior: "smooth", block: "start"}); return; }
    st.open[keyP(p)] = true; draw();
    const el = [...box.querySelectorAll("[data-k]")].find(x => x.dataset.k === keyP(p));
    if (el){ el.classList.add("flash"); el.scrollIntoView({behavior: "smooth", block: "center"}); setTimeout(() => el.classList.remove("flash"), 1600); }
  };
  draw();
  return c;
}
'''

# ---- 6. styles ----
edit("Variations styles",
     '''.dt .foot{display:flex;justify-content:center}''',
     '''.dt .foot{display:flex;justify-content:center}

/* Totals: Variations pivot (Project > Source file > Job) */
.clickrows .viz .hit{cursor:pointer}
.vpiv{border:1px solid var(--line);border-radius:10px;overflow:auto;max-height:720px}
.vpiv .vr{display:grid;grid-template-columns:minmax(220px,1.4fr) 118px 118px 118px minmax(160px,1.2fr);gap:12px;align-items:center;padding:7px 12px;border-bottom:1px solid var(--line);font-size:13px;min-width:720px}
.vpiv .vr:last-child{border-bottom:0}
.vpiv .vh{position:sticky;top:0;z-index:1;background:var(--card);font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
.vpiv .c0{display:flex;gap:6px;align-items:center;min-width:0}
.vpiv .c0 span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.vpiv .tw{font-style:normal;display:inline-block;width:12px;color:var(--muted);transition:transform .12s}
.vpiv .open > .c0 .tw{transform:rotate(90deg)}
.vpiv .l0{background:var(--acc-soft);font-weight:700;cursor:pointer}
.vpiv .l1{font-weight:650;cursor:pointer}
.vpiv .l1 .c0{padding-left:20px}
.vpiv .l2 .c0{padding-left:46px}
.vpiv .l2{color:var(--ink2)}
.vpiv .l0:hover,.vpiv .l1:hover{background:color-mix(in srgb,var(--acc2) 14%,var(--card))}
.vpiv [data-k]:focus-visible{outline:2px solid var(--acc2);outline-offset:-2px}
.vpiv .vt{background:var(--acc);color:#fff;font-weight:700}
.vpiv .vt .vn.pos,.vpiv .vt .vn.neg{color:#fff}
.vpiv .vn{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.vpiv .vn.pos,.vgt .pos{color:var(--up)}
.vpiv .vn.neg,.vgt .neg{color:var(--down)}
.vpiv .reason{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--muted)}
.vpiv .vr.flash{box-shadow:inset 4px 0 0 var(--acc2);background:color-mix(in srgb,var(--acc2) 22%,var(--card))}
.vpiv .vmsg{padding:10px 12px;background:color-mix(in srgb,var(--warning) 16%,var(--card));font-size:13px}
.vgt{display:flex;flex-wrap:wrap;gap:6px 18px;align-items:center;margin-top:10px;padding:9px 12px;border-radius:10px;background:var(--acc-soft);font-size:13.5px;font-variant-numeric:tabular-nums}
.vsec h2{margin:0;font-size:20px}
.vsec p{margin:2px 0 14px;color:var(--muted);font-size:13.5px}''')


def apply(texts):
    """texts: {name: content}. Returns ({name: new content}, changed?)."""
    texts = dict(texts)
    joined = "\n".join(texts.values())
    if "function ntVarPivot(P, vr){" in joined:
        return texts, False
    for desc, old, new in EDITS:
        hits = [n for n, t in texts.items() if t.count(old) == 1]
        if len(hits) != 1:
            raise SystemExit(f"Couldn't find '{desc}' exactly as expected - nothing was changed.")
        texts[hits[0]] = texts[hits[0]].replace(old, new)
    hits = [n for n, t in texts.items() if OLD_TOTALS_START in t and OLD_TOTALS_END in t]
    if len(hits) != 1:
        raise SystemExit("Couldn't find the Totals tab (ntTotals) - nothing was changed.")
    t = texts[hits[0]]
    a, b = t.index(OLD_TOTALS_START), t.index(OLD_TOTALS_END)
    texts[hits[0]] = t[:a] + NEW_TOTALS + t[b:]
    return texts, True


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    names = sys.argv[1:] or FILES
    paths = {n: os.path.join(here, n) if not os.path.isabs(n) else n for n in names}
    texts = {n: open(p, encoding="utf-8").read() for n, p in paths.items() if os.path.exists(p)}
    new, changed = apply(texts)
    if not changed:
        print("The dashboard changes are already there - nothing to do.")
    else:
        for n, t in new.items():
            if t != texts[n]:
                open(paths[n], "w", encoding="utf-8").write(t)
                print(f"Patched {n}")
        print("Now run: python build_page.py")
