/* ======================================================================
   Dashboards - each one gets its own space: coloured app bar, switcher,
   file strip, one filter bar, tabs and a grid of cards.
   ====================================================================== */
const DASHES = [
  {key: "tracker", c: "c2", icon: "grid", title: "Network Job Tracker",
   desc: "Network jobs from the Master file, the Outages Programme calendar, mapped items, pole position and totals."},
  {key: "master", c: "c1", icon: "pound", title: "Master Control",
   desc: "Job costing: totals by date and team leader, PID breakdown and financial health."},
  {key: "materials", c: "c3", icon: "box3", title: "Materials Breakdown",
   desc: "Poles, Free Issue and GUK materials for the filtered Control File, with the maps check and Excel export."},
];
const DST = {};            // per-dashboard state, kept while you switch between them
let PYSTAT = ["Starting Python…", "busy"];

function pyPill(){ return h(`<span class="pill pystat"><span class="dot ${PYSTAT[1] === "on" ? "on" : PYSTAT[1] === "off" ? "off" : "busy"}"></span><span>${esc(PYSTAT[0])}</span></span>`); }
function dashBar(key){
  const d = DASHES.find(x => x.key === key);
  const el = h(`<div class="dbar"><div class="nav"><button class="btn sm">${I.back} All tools</button><div class="dsw"></div><span class="spacer"></span></div>
    <div class="dtitle"><div class="ico">${I[d.icon]}</div><div><h1></h1><p></p></div><span class="spacer"></span><span class="pill" title="Page version">v 25-Sep-2026</span></div></div>`);
  $("h1", el).textContent = d.title; $(".dtitle p", el).textContent = d.desc;
  $(".nav .btn", el).onclick = () => { location.hash = ""; };
  const sw = $(".dsw", el);
  DASHES.forEach(x => { const b = h(`<button class="${x.key === key ? "on" : ""}">${I[x.icon]}<span></span></button>`); $("span", b).textContent = x.title;
    b.onclick = () => { location.hash = "dash/" + x.key; }; sw.append(b); });
  const nav = $(".nav", el);
  const tb = h(`<button class="btn sm themebtn" title="Colour theme"></button>`); tb.textContent = Theme.label();
  tb.onclick = () => { Theme.next(); tb.textContent = Theme.label(); };
  nav.append(pyPill(), tb);
  return el;
}
const spinner = text => { const s = h(`<div class="dspin"><i></i><span></span></div>`); $("span", s).textContent = text; return s; };
const errNote = e => note(String(e && e.message || e), "err", "err");
function tabBtn(icon, label){ const b = h(`<button>${ICON2[icon] || I[icon] || ""}<span></span></button>`); $("span", b).textContent = label; return b; }

/* ---------------- generic driver ---------------- */
function pageDashboard(key){
  const def = DASHDEF[key];
  document.body.classList.add("dashmode"); document.body.dataset.dash = key;
  const st = DST[key] || (DST[key] = {tab: def.tabs[0].id, f: {}, cols: null});
  const P = {key, def, st, token: dashToken, seq: 0, tseq: 0, ck: st.ck || (st.ck = colorKeeper())};
  P.alive = () => P.token === dashToken;
  const strip = h(`<div class="dfiles"></div>`);
  def.slots.forEach(([s, req]) => strip.append(slotEl(s, req)));
  P.body = h(`<div></div>`);
  app.append(dashBar(key), strip, P.body);
  const onFile = slot => {
    if (!P.alive()) { fileListeners.delete(onFile); return; }
    if (!def.slots.some(x => x[0] === slot)) return;
    if (def.loadSlots.includes(slot)) P.start(); else if (P.v) P.renderTab();
  };
  fileListeners.add(onFile);

  P.start = async () => {
    if (!P.alive()) return;
    const missing = def.slots.filter(([s, req]) => req && !FILES[s]);
    if (missing.length){ P.body.textContent = ""; P.body.append(landing(P)); return; }
    const sig = def.loadSlots.map(s => FILES[s].v).join("|") + JSON.stringify(st.cols || {});
    if (st.sig !== sig || !st.L){
      P.body.textContent = ""; P.body.append(h(`<div class="dpanel loading"></div>`)); $(".dpanel", P.body).append(spinner(PY.ready ? "Reading the files…" : "Starting Python (first time ~30 s)…"));
      try { st.L = (await PY.call(def.loadFn, def.loadArgs(P), {needs: def.needs})).result; st.sig = sig; st.fresh = true; }
      catch(e){ if (!P.alive()) return; P.body.textContent = ""; P.body.append(errNote(e)); st.L = null; return; }
      if (!P.alive()) return;
      def.init && def.init(P);
    }
    P.layout();
  };
  P.layout = () => {
    P.body.textContent = "";
    if (def.pre) { const x = def.pre(P); if (x) P.body.append(x); }
    if (def.ready && !def.ready(P)) return;
    P.fbar = filterBar(P, def.filters(P));
    P.body.append(P.fbar);
    const tabs = h(`<div class="dtabs" role="tablist"></div>`);
    def.tabs.forEach(t => { const b = tabBtn(t.icon, t.label); b.classList.toggle("on", t.id === st.tab);
      b.onclick = () => { st.tab = t.id; tabs.querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b)); P.renderTab(); }; tabs.append(b); });
    P.panel = h(`<div class="dpanel"></div>`);
    P.body.append(tabs, P.panel);
    P.refresh();
  };
  P.refresh = async () => {
    const my = ++P.seq;
    P.panel.classList.add("loading"); P.panel.querySelectorAll(".dspin").forEach(n => n.remove()); P.panel.prepend(spinner("Updating…"));
    try {
      const v = (await PY.call(def.viewFn, def.viewArgs(P), {needs: def.needs})).result;
      if (my !== P.seq || !P.alive()) return;
      P.v = v;
      /* cascading filters: drop picks that the other filters have made impossible */
      let dropped = false;
      (def.cascade || []).forEach(([k, optKey]) => { const ok = new Set(v[optKey] || []); const keep = (st.f[k] || []).filter(x => ok.has(x)); if (keep.length !== (st.f[k] || []).length){ st.f[k] = keep; dropped = true; } });
      P.fbar.update(v);
      if (dropped) return P.refresh();
      P.renderTab();
    } catch(e){
      if (my !== P.seq || !P.alive()) return;
      P.panel.textContent = ""; P.panel.append(errNote(e));
    } finally { if (my === P.seq){ P.panel.classList.remove("loading"); P.panel.querySelectorAll(".dspin").forEach(n => n.remove()); } }
  };
  P.renderTab = () => {
    if (!P.v) return;
    P.tseq++;
    Tip.hide();
    P.panel.textContent = "";
    if (def.header) def.header(P, P.panel);
    const t = def.tabs.find(x => x.id === st.tab) || def.tabs[0];
    try { t.render(P, P.panel); } catch(e){ console.error(e); P.panel.append(errNote(e)); }
  };
  /* async work inside a tab (outages, maps...) - ignored if the tab was re-drawn meanwhile */
  P.tabCall = async (fn, kwargs, needs) => { const my = P.tseq; const r = await PY.call(fn, kwargs, {needs: needs || def.needs}); if (my !== P.tseq || !P.alive()) throw {stale: true}; return r.result; };
  P.start();
}

function landing(P){
  const d = DASHES.find(x => x.key === P.key), def = P.def;
  const el = h(`<div class="dland"><div><h2></h2><p>Drop the files below. They're read here in your browser – nothing is uploaded anywhere, and the files you drop are shared by all three dashboards, so a Master file dropped once works everywhere.</p><ul></ul></div><div class="zones"></div></div>`);
  $("h2", el).textContent = "Load the data for " + d.title;
  const ul = $("ul", el); (def.about || []).forEach(t => { const li = document.createElement("li"); li.textContent = t; ul.append(li); });
  const z = $(".zones", el);
  def.slots.forEach(([s, req]) => { if (req || def.landingOptional) z.append(slotZone(s, req, def.slotNotes && def.slotNotes[s])); });
  return el;
}

/* ---------------- the one filter bar ---------------- */
function filterBar(P, spec){
  const f = P.st.f;
  const el = h(`<div class="fbar"><div class="fr r1"></div><div class="fr r2"></div><div class="achips"></div></div>`);
  const r1 = $(".r1", el), r2 = $(".r2", el), chips = $(".achips", el);
  const lbl = t => { const s = h(`<span class="lbl"></span>`); s.textContent = t; return s; };
  const changed = () => { drawChips(); P.refresh(); };
  if (spec.dateFields && spec.dateFields.length){
    r1.append(lbl("Date field"), segc(spec.dateFields, f[spec.dateKey], v => { f[spec.dateKey] = v; changed(); }));
    r1.append(h(`<span class="sep"></span>`));
  }
  const dateIn = (key, title) => { const i = h(`<input type="date">`); i.title = title; i.value = f[key] || ""; i.onchange = () => { f[key] = i.value || null; changed(); }; return i; };
  let custom = null;
  if (spec.presets){
    const sel = h(`<select title="Quick range"></select>`);
    spec.presets.forEach(p => { const o = document.createElement("option"); o.textContent = p; sel.append(o); });
    sel.value = f.preset || "All time";
    custom = h(`<span class="fr" style="gap:6px"></span>`);
    const cf = dateIn("custom_from", "From"), ct = dateIn("custom_to", "To");
    custom.append(cf, h(`<span class="muted">→</span>`), ct);
    custom.from = cf; custom.to = ct;
    custom.classList.toggle("hidden", sel.value !== "Custom range");
    sel.onchange = () => { f.preset = sel.value; custom.classList.toggle("hidden", sel.value !== "Custom range"); changed(); };
    r1.append(lbl("Range"), sel, custom);
  } else if (spec.fromTo){
    const cf = dateIn(spec.fromTo[0], "From (blank = no limit)"), ct = dateIn(spec.fromTo[1], "To (blank = no limit)");
    r1.append(lbl("From"), cf, lbl("To"), ct);
  }
  if (spec.gran){
    const sel = h(`<select title="Date grouping"></select>`);
    spec.gran.forEach(g => { const o = document.createElement("option"); o.textContent = g; sel.append(o); });
    sel.value = f.granularity || spec.gran[0];
    sel.onchange = () => { f.granularity = sel.value; changed(); };
    r1.append(h(`<span class="sep"></span>`), lbl("Group by"), sel);
  }
  const cap = h(`<span class="range"></span>`); r1.append(cap);

  r2.append(lbl("Filters"));
  const ms = {};
  spec.multis.forEach(m => {
    const w = msel({label: m.label, options: m.options || [], value: f[m.key] || [], cascade: m.cascade, hint: m.hint,
      onChange: v => { f[m.key] = v; changed(); }});
    ms[m.key] = w; r2.append(w);
  });
  const reset = h(`<button class="btn sm ghost" title="Clear every filter">Reset filters</button>`);
  reset.onclick = () => { spec.multis.forEach(m => f[m.key] = []); if (spec.presets){ f.preset = "All time"; f.custom_from = f.custom_to = null; }
    if (spec.fromTo) spec.fromTo.forEach(k => f[k] = null); if (spec.gran) f.granularity = spec.gran[0]; spec.onReset && spec.onReset(); P.layout(); };
  r2.append(reset);

  function drawChips(){
    chips.textContent = "";
    spec.multis.forEach(m => {
      const vals = f[m.key] || [];
      vals.slice(0, 6).forEach(v => {
        const c = h(`<span class="achip"><b></b><span></span><button title="Remove">×</button></span>`);
        $("b", c).textContent = m.label; $("span", c).textContent = v; c.title = `${m.label}: ${v}`;
        $("button", c).onclick = () => { f[m.key] = (f[m.key] || []).filter(x => x !== v); ms[m.key].set(f[m.key]); changed(); };
        chips.append(c);
      });
      if (vals.length > 6){ const c = h(`<span class="achip"><b></b><span></span></span>`); $("b", c).textContent = m.label; $("span", c).textContent = `+${vals.length - 6} more`; chips.append(c); }
    });
  }
  el.update = v => {
    spec.multis.forEach(m => { if (m.cascade) ms[m.key].setOptions(v[m.cascade] || []); ms[m.key].set(f[m.key] || []); });
    cap.innerHTML = spec.caption ? spec.caption(v) : "";
    if (custom){ if (!custom.from.value && v.date_from) custom.from.value = v.date_from; if (!custom.to.value && v.date_to) custom.to.value = v.date_to; }
    drawChips();
  };
  drawChips();
  return el;
}
const capHTML = parts => parts.filter(Boolean).map(p => Array.isArray(p) ? `<b>${esc(p[0])}</b> ${esc(p[1] || "")}` : esc(p)).join(" · ");

/* ---------- shared renderers ---------- */
function tableCard(title, sub, table, span = 12){ const c = dcard(title, sub, span); c.body.append(table.rows && table.rows.length ? dtable(table) : emptyMsg("No rows for the current filters.")); return c; }
const tbl = (t, name, extra) => ({columns: t.columns, rows: t.rows, name, ...(extra || {})});

/* ======================================================================
   Master Control
   ====================================================================== */
const DASHDEF = {};
const STAGES = [["invoiced", "Invoiced", "var(--s7)"], ["done", "Done", "var(--s3)"], ["planned", "Planned", "var(--s4)"], ["remaining", "Remaining", "var(--neutral)"]];
DASHDEF.master = {
  slots: [["master", true]], loadSlots: ["master"], needs: NEED_MC,
  about: ["Master_DD-MM-YYYY (vN).parquet from 28.Project Tracker – the same file the Streamlit app picked automatically.",
          "Filters sit in one bar above the charts; Circuit and Pole narrow to what the other filters allow."],
  loadFn: "mc_load", loadArgs: () => ({path: FILES.master.path}),
  init(P){ const f = P.st.f, L = P.st.L;
    if (!L.date_fields.includes(f.date_field)) f.date_field = L.default_date_field;
    f.preset = f.preset || "All time"; f.granularity = f.granularity || "Auto"; f.pid_mode = f.pid_mode || "All"; },
  viewFn: "mc_view",
  viewArgs(P){ const f = P.st.f; return {path: FILES.master.path, date_field: f.date_field, preset: f.preset, custom_from: f.custom_from || null, custom_to: f.custom_to || null,
    districts: f.districts || [], projects: f.projects || [], pids: f.pids || [], pms: f.pms || [], sourcefiles: f.sourcefiles || [],
    circuits: f.circuits || [], poles: f.poles || [], granularity: f.granularity, pid_mode: f.pid_mode}; },
  cascade: [["circuits", "circuit_options"], ["poles", "pole_options"]],
  pre(P){
    const L = P.st.L; if (!L.other_count) return null;
    const d = h(`<details class="dmap"><summary></summary><p class="muted" style="margin:0 0 6px;font-size:13px">These jobs don't start with 'M -' / 'M-' (material) or 'C -' / 'C-' (construction), so they're excluded from every chart. Their dates still count toward the date range. Sample job codes:</p><div class="chips"></div></details>`);
    $("summary", d).textContent = `⚠ ${intf(L.other_count)} row(s) not classified`;
    L.other_sample.forEach(s => { const c = h(`<span class="chip"><span></span></span>`); $("span", c).textContent = s; $(".chips", d).append(c); });
    return d;
  },
  filters(P){ const L = P.st.L, o = L.options;
    return {dateFields: L.date_fields, dateKey: "date_field", presets: L.presets, gran: ["Auto", "Day", "Week", "Month", "Year"],
      multis: [{key: "districts", label: "District", options: o.district}, {key: "projects", label: "Project", options: o.project},
               {key: "pids", label: "PID", options: o.pid}, {key: "pms", label: "Project Manager", options: o.pm},
               {key: "sourcefiles", label: "Source file", options: o.sourcefile},
               {key: "circuits", label: "Circuit", cascade: "circuit_options", hint: "Narrowed to circuits that exist given the other filters."},
               {key: "poles", label: "Pole (enid)", cascade: "pole_options", hint: "Narrowed to poles that exist given the other filters (including Circuit)."}],
      caption: v => capHTML([[fmtDate(v.date_from) + " → " + fmtDate(v.date_to)], [intf(v.records), "records"], `by ${v.granularity}${v.granularity_auto ? " (auto)" : ""}`])}; },
  tabs: [
    {id: "trends", label: "Trends", icon: "trend", render: mcTrends},
    {id: "pid", label: "PID Breakdown", icon: "table", render: mcPid},
    {id: "finance", label: "Finance", icon: "pound", render: mcFinance},
  ],
};
function mcTrends(P, panel){
  const v = P.v, p1 = v.trends.panel1, k = p1.kpis;
  if (!p1.labels.length){ panel.append(note("No records under these filters. Widen the date range or clear a filter.")); return; }
  panel.append(kpis([
    {label: "Total (construction + materials)", value: money(k.grand_total), hero: true},
    {label: "Construction total", value: money(k.construction), swatch: "var(--s3)"},
    {label: "Original", value: money(k.original)},
    {label: "Variance", value: signed(k.variance), tone: k.variance >= 0 ? "up" : "down", note: k.variance >= 0 ? "▲ over the original" : "▼ under the original"},
    {label: "Materials", value: money(k.materials), swatch: "var(--s4)"},
  ]));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const varCol = i => p1.variance[i] >= 0 ? "var(--s1)" : "var(--s8)";
  g.append(vizCard({title: "Panel 01 — Total by date (construction + materials)", sub: "Each column = materials + construction base + variance vs. original, stacked",
    legend: [{name: "Materials", color: "var(--s4)"}, {name: "Construction base (original / the rest)", color: "var(--s3)"}, {name: "Positive variation", color: "var(--s1)"}, {name: "Negative variation", color: "var(--s8)"}],
    draw: host => Viz.columns(host, {labels: p1.labels, height: 340, axis: v => moneyS(v),
      series: [{name: "Materials", values: p1.material, color: "var(--s4)"}, {name: "Base", values: p1.base, color: "var(--s3)"}, {name: "Variance", values: p1.cap, colorAt: varCol}],
      tip: i => ({title: p1.labels[i], rows: [{color: "var(--s3)", name: "Construction total", value: money(p1.total[i])}, {name: "Original", value: money(p1.orig[i])},
        {color: varCol(i), name: p1.variance[i] >= 0 ? "Positive variation" : "Negative variation", value: signed(p1.variance[i])}, {color: "var(--s4)", name: "Materials", value: money(p1.material[i])}],
        total: {name: "Total", value: money(p1.total[i] + p1.material[i])}})}),
    table: () => ({name: "Total by date", columns: ["Period", "Construction total (£)", "Original (£)", "Variance (£)", "Materials (£)", "Total (£)"],
      rows: p1.labels.map((l, i) => [l, p1.total[i], p1.orig[i], p1.variance[i], p1.material[i], p1.total[i] + p1.material[i]])})}));
  const p2 = v.trends.panel2;
  const people = p2.leaders.filter(l => l !== "Other"), col = P.ck.assign(people);
  const cOf = l => l === "Other" ? "var(--neutral)" : col(l);
  const c2 = vizCard({title: "Panel 02 — Total by team leader (construction)", sub: "Stacked by team leader · same filters, date grouping and periods as Panel 01" + (p2.has_other ? ` · top ${p2.top_n} shown individually, the rest grouped as Other` : ""),
    legend: p2.leaders.map(l => ({name: l, color: cOf(l)})),
    draw: host => Viz.columns(host, {labels: p2.labels, height: 320, axis: v => moneyS(v), fmt: money,
      series: p2.leaders.map(l => ({name: l, values: p2.values[l], color: cOf(l)})),
      tip: i => { const rows = p2.leaders.map(l => ({color: cOf(l), name: l, value: money(p2.values[l][i] || 0)})).filter((r, j) => (p2.values[p2.leaders[j]][i] || 0) !== 0);
        return {title: p2.labels[i], rows, total: {name: "Total", value: money(p2.leaders.reduce((a, l) => a + (p2.values[l][i] || 0), 0))}}; }}),
    table: () => ({name: "Total by team leader", columns: ["Period", ...p2.leaders.map(l => l + " (£)")], rows: p2.labels.map((lab, i) => [lab, ...p2.leaders.map(l => p2.values[l][i] || 0)])})});
  const k2 = p2.kpis;
  c2.body.prepend(kpis([{label: "Total (construction)", value: money(k2.total)}, {label: "Original", value: money(k2.original)},
    {label: "Variance", value: signed(k2.variance), tone: k2.variance >= 0 ? "up" : "down", note: k2.variance >= 0 ? "▲ over" : "▼ under"}]));
  g.append(c2);
}
function mcPid(P, panel){
  const v = P.v, f = P.st.f, pid = v.pid;
  const bar = h(`<div class="fr" style="display:flex;gap:10px;align-items:center;margin-bottom:12px"><span class="lbl" style="font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase">Show</span></div>`);
  bar.append(segc(["All", "Construction", "Material"], f.pid_mode, x => { f.pid_mode = x; P.refresh(); }));
  panel.append(bar);
  if (!pid.rows.length){ panel.append(note("No PIDs under these filters.")); return; }
  const k = pid.kpis;
  panel.append(kpis([{label: "Total", value: money(k.total), hero: true}, {label: "Remaining", value: money(k.remaining), swatch: "var(--neutral)"},
    {label: "Planned", value: money(k.planned), swatch: "var(--s4)"}, {label: "Done", value: money(k.done), swatch: "var(--s3)"},
    {label: "Invoiced", value: money(k.invoiced), swatch: "var(--s7)"}, {label: "Variance", value: signed(k.variance), tone: k.variance >= 0 ? "up" : "down", note: k.variance >= 0 ? "▲ over value" : "▼ under value"}]));
  const c = dcard("Panel 03 — PID breakdown", `District → Project → Project Manager → PID · ${pid.rows.length} PID(s) · each bar is split by that PID's own total, so proportions compare across PID sizes`);
  let mode = "chart";
  const holder = h(`<div></div>`);
  c.acts.append(segc([["chart", "Bars"], ["table", "Table"]], mode, m => { mode = m; draw(); }));
  c.body.append(holder); panel.append(c);
  function draw(){
    holder.textContent = "";
    if (mode === "table"){ holder.append(dtable({name: "PID breakdown", columns: ["District", "Project", "Project Manager", "PID", "Job", "Total (£)", "Invoiced (£)", "Done (£)", "Planned (£)", "Remaining (£)", "Variance (£)"],
      rows: pid.rows.map(r => [r.district, r.project, r.pm, r.pid, r.job, r.total, r.invoiced, r.done, r.planned, r.remaining, r.variance])})); return; }
    holder.append(legendEl(STAGES.map(s => ({name: s[1], color: s[2]}))));
    const list = h(`<div class="pidlist scrollv" style="max-height:900px"></div>`);
    const lab = x => x == null || x === "" ? "Unassigned" : x;
    const sum = pred => pid.rows.filter(pred).reduce((a, r) => a + (r.total || 0), 0);
    let ld, lp, lm; const parts = [];
    pid.rows.forEach((r, i) => {
      const d = lab(r.district), p = lab(r.project), m = lab(r.pm);
      if (d !== ld){ parts.push(`<div class="pidg1"><span>${esc(d)}</span><span>${esc(money(sum(x => lab(x.district) === d)))}</span></div>`); ld = d; lp = lm = null; }
      if (p !== lp){ parts.push(`<div class="pidg2"><span>${esc(p)}</span><span>${esc(money(sum(x => lab(x.district) === d && lab(x.project) === p)))}</span></div>`); lp = p; lm = null; }
      if (m !== lm){ parts.push(`<div class="pidg3"><span>${esc(m)}</span><span>${esc(moneyS(sum(x => lab(x.district) === d && lab(x.project) === p && lab(x.pm) === m)))}</span></div>`); lm = m; }
      const tot = r.total || 0;
      const segs = STAGES.map(([key, , color]) => { const w = tot > 0 ? (r[key] || 0) / tot * 100 : 0; return w > 0 ? `<i data-i="${i}" data-s="${key}" style="width:calc(${w.toFixed(3)}% - 2px);background:${color}"></i>` : ""; }).join("");
      const ip = tot > 0 ? (r.invoiced || 0) / tot * 100 : 0, pos = (r.variance || 0) >= 0;
      parts.push(`<div class="pidr" data-row="${i}"><div class="pn" title="${esc(r.job)}"><b>${esc(r.pid)}</b><small>${esc(r.job)}</small></div><div class="bar">${segs}</div>
        <div class="tv">${esc(moneyS(tot))}<small style="display:block;color:var(--muted);font-weight:400;font-size:11.5px">Inv ${esc(moneyS(r.invoiced))} (${ip.toFixed(0)}%)</small></div>
        <div class="vv ${pos ? "pos" : "neg"}" data-i="${i}" data-s="variance">${pos ? ICON2.up : ICON2.down}${esc(signed(r.variance, moneyS))}</div></div>`);
    });
    list.innerHTML = parts.join("");
    list.addEventListener("mousemove", e => { const t = e.target.closest("[data-s]"); if (!t){ Tip.hide(); return; } const r = pid.rows[+t.dataset.i];
      if (t.dataset.s === "variance"){ Tip.show(e, {title: `${r.pid} — ${(r.variance || 0) >= 0 ? "Over value" : "Under value"}`, rows: [{name: "Variance vs. original (construction)", value: signed(r.variance)}]}); return; }
      const st = STAGES.find(s => s[0] === t.dataset.s); Tip.show(e, {title: r.pid, rows: STAGES.map(s => ({color: s[2], name: s[1] + (s[0] === st[0] ? " ◂" : ""), value: money(r[s[0]] || 0)})), total: {name: "Total", value: money(r.total)}}); });
    list.addEventListener("mouseleave", () => Tip.hide());
    holder.append(list);
  }
  draw();
}
function mcFinance(P, panel){
  const fin = P.v.finance, k = fin.kpis;
  if (!k || k.total == null){ panel.append(note("No records under these filters. Widen the date range or clear a filter.")); return; }
  panel.append(kpis([{label: "Total value", value: money(k.total), hero: true}, {label: "Original budget", value: money(k.original)},
    {label: "Variance", value: signed(k.variance), tone: k.variance >= 0 ? "up" : "down", note: `${k.variance_pct >= 0 ? "▲" : "▼"} ${(k.variance_pct >= 0 ? "+" : "") + k.variance_pct.toFixed(2)}% vs. original`}]));
  panel.append(kpis([{label: "Invoiced", value: money(k.invoiced), swatch: "var(--s7)", note: `${k.invoiced_pct.toFixed(1)}% of total`},
    {label: "WIP — done, not invoiced", value: money(k.wip), swatch: "var(--s3)"}, {label: "Backlog — not yet done", value: money(k.backlog), swatch: "var(--neutral)"},
    {label: "Material share", value: k.material_pct.toFixed(1) + "%", swatch: "var(--s4)"}]));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const pc = {"Remaining (not started)": "var(--neutral)", "Planned (not done)": "var(--s4)", "Done (not invoiced)": "var(--s3)", "Invoiced": "var(--s7)"};
  const pipe = fin.pipeline, ptot = pipe.reduce((a, s) => a + s.value, 0);
  g.append(vizCard({title: "Billing pipeline", sub: "Where the total value currently sits, end to end", span: 6,
    legend: pipe.map(s => ({name: s.stage, color: pc[s.stage]})),
    draw: host => { const bh = h(`<div></div>`); host.append(bh);
      const tiles = h(`<div class="pipe"></div>`); pipe.forEach(s => { const t = h(`<div><small><i></i><span></span></small><b></b><span></span></div>`);
        $("i", t).style.background = pc[s.stage]; $("small span", t).textContent = s.stage; $("b", t).textContent = money(s.value);
        t.lastChild.textContent = `${ptot ? (s.value / ptot * 100).toFixed(1) : 0}% of the total`; tiles.append(t); }); host.append(tiles);
      Viz.bars(bh, {labels: [""], rowH: 56, fmt: money, valueLabel: () => moneyS(ptot),
      series: pipe.map(s => ({name: s.stage, values: [s.value], color: pc[s.stage]})),
      tip: () => ({title: "Billing pipeline", rows: pipe.map(s => ({color: pc[s.stage], name: s.stage, value: `${money(s.value)} (${ptot ? (s.value / ptot * 100).toFixed(1) : 0}%)`})), total: {name: "Total", value: money(ptot)}})}); },
    table: () => ({name: "Billing pipeline", columns: ["Stage", "Value (£)"], rows: pipe.map(s => [s.stage, s.value])})}));
  const lag = fin.lag;
  if (!lag || !lag.n){ const c = dcard("Invoicing speed", "Days between job marked done and invoiced", 6); c.body.append(emptyMsg("Not enough done + invoiced pairs under these filters to measure lag.")); g.append(c); }
  else {
    const e = lag.edges, labs = lag.counts.map((_, i) => `${Math.round(e[i])}–${Math.round(e[i + 1])}`);
    const c = vizCard({title: "Invoicing speed", sub: "Days between job marked done and invoiced · jobs per bucket", span: 6,
      draw: host => Viz.columns(host, {labels: labs, height: 230, fill: true, axis: intf, series: [{name: "Jobs", values: lag.counts, color: "var(--s1)"}],
        marker: {x: (lag.median - e[0]) / ((e[e.length - 1] - e[0]) || 1) * lag.counts.length, label: `median ${Math.round(lag.median)} d`},
        tip: i => ({title: `${e[i].toFixed(0)}–${e[i + 1].toFixed(0)} days`, rows: [{color: "var(--s1)", name: "Jobs", value: intf(lag.counts[i])}]})}),
      table: () => ({name: "Invoicing speed", columns: ["From (days)", "To (days)", "Jobs"], rows: lag.counts.map((n, i) => [+e[i].toFixed(1), +e[i + 1].toFixed(1), n])})});
    const s = h(`<p style="margin:0 0 8px;font-size:13.5px"></p>`);
    s.innerHTML = `Median: <b>${Math.round(lag.median)} days</b> · Average: <b>${Math.round(lag.mean)} days</b> · n=${intf(lag.n)}`;
    c.body.prepend(s); g.append(c);
  }
  const pr = fin.projects;
  if (!pr.length){ const c = dcard("Variance by project", "Construction total vs. original budget", 12); c.body.append(emptyMsg("No projects with a non-zero original budget under these filters.")); g.append(c); }
  else g.append(vizCard({title: "Variance by project", sub: "Construction total vs. original budget, worst to best", span: 12,
    draw: host => Viz.diverging(host, {labels: pr.map(p => p.project || "Unassigned"), values: pr.map(p => p.pct), axis: t => t + "%", valueLabel: i => (pr[i].pct >= 0 ? "+" : "") + pr[i].pct.toFixed(1) + "%",
      tip: i => ({title: pr[i].project || "Unassigned", rows: [{name: "Total", value: money(pr[i].total)}, {name: "Original", value: money(pr[i].orig)}, {color: pr[i].variance >= 0 ? "var(--s1)" : "var(--s8)", name: "Variance", value: `${signed(pr[i].variance)} (${(pr[i].pct >= 0 ? "+" : "") + pr[i].pct.toFixed(1)}%)`}]})}),
    table: () => ({name: "Variance by project", pct: ["Variance %"], columns: ["Project", "Total (£)", "Original (£)", "Variance (£)", "Variance %"], rows: pr.map(p => [p.project, p.total, p.orig, p.variance, p.pct])})}));
  const lb = (rows, name) => ({name, pct: ["Variance %"], columns: ["PID", "District", "Project", "Variance (£)", "Variance %"], rows: rows.map(r => [r.pid, r.district, r.project, r.variance, r.pct]), search: false});
  if (fin.worst.length || fin.best.length){
    const a = dcard("Top overruns", "Largest £ variance below the original (construction only)", 6); a.body.append(dtable(lb(fin.worst, "Top overruns")));
    const b = dcard("Top underspends", "Largest £ variance above the original (construction only)", 6); b.body.append(dtable(lb(fin.best, "Top underspends")));
    $("h3", a).prepend(h(`<span style="color:var(--down);margin-right:6px">${ICON2.down}</span>`)); $("h3", b).prepend(h(`<span style="color:var(--up);margin-right:6px">${ICON2.up}</span>`));
    g.append(a, b);
  } else { const c = dcard("PID variance leaderboard", "", 12); c.body.append(emptyMsg("No PIDs with a non-zero original budget under these filters.")); g.append(c); }
}

/* ======================================================================
   Network Job Tracker
   ====================================================================== */
DASHDEF.tracker = {
  slots: [["master", true], ["outages", false], ["forecast", false], ["ics", false]], loadSlots: ["master"], needs: NEED_NT,
  about: ["Master file (.parquet or .csv) – required. The other three are optional and can be dropped any time in the strip at the top:",
          "Outages Programme (High-level_planning_2026.xlsx) → the Jobs tab calendar with the poles & tasks for each day.",
          "Service Partner Workbank → Pole Position.  Outage calendar (.ics) → the events list."],
  loadFn: "nt_load", loadArgs: P => ({path: FILES.master.path, cols: P.st.cols}),
  init(P){ const st = P.st, L = st.L, f = st.f; st.cols = L.cols;
    if (L.date_fields && !L.date_fields.includes(f.date_field)) f.date_field = L.default_date_field;
    f.preset = f.preset || "All time"; st.gran = st.gran || "Month"; st.o = st.o || {view: "calendar"}; st.fc = st.fc || {}; },
  pre(P){
    const L = P.st.L, wrap = h(`<div></div>`);
    const d = h(`<details class="dmap"><summary>⚙ Column mapping – confirm each field maps to the right column in your file</summary><div class="frm"></div></details>`);
    if (L.missing_required.length) d.open = true;
    L.col_fields.forEach(fl => {
      const miss = !L.cols[fl.key];
      const fe = h(`<div class="field${miss ? " miss" : ""}"><label></label><select></select></div>`);
      $("label", fe).textContent = fl.label + (miss ? " – pick one" : "");
      const sel = $("select", fe); [["", "(none)"], ...L.columns.map(c => [c, c])].forEach(([v, t]) => { const o = document.createElement("option"); o.value = v; o.textContent = t; sel.append(o); });
      sel.value = L.cols[fl.key] || "";
      sel.onchange = () => { P.st.cols = {...P.st.cols, [fl.key]: sel.value || null}; P.start(); };
      $(".frm", d).append(fe);
    });
    const cols = h(`<p class="muted" style="font-size:12.5px;margin:10px 0 0"></p>`); cols.textContent = `Detected columns (${L.columns.length}): ${L.columns.join(", ")}`; d.append(cols);
    wrap.append(d);
    if (L.missing_required.length) wrap.append(note(`Couldn't guess a column for: ${L.missing_required.map(k => (L.col_fields.find(x => x.key === k) || {}).label || k).join(", ")}. Pick them under Column mapping above.`, "warn", "warn"));
    else if (L.warn_no_pole) wrap.append(note("No pole column mapped – pole de-duplication and the pole filter are off until you pick one under Column mapping.", "warn", "warn"));
    return wrap;
  },
  ready: P => !P.st.L.missing_required.length,
  viewFn: "nt_view",
  ntFilters(P){ const f = P.st.f; return {date_field: f.date_field, preset: f.preset, custom_from: f.custom_from || null, custom_to: f.custom_to || null,
    districts: f.districts || [], projects: f.projects || [], pids: f.pids || [], sourcefiles: f.sourcefiles || [], circuits: f.circuits || [], poles: f.poles || []}; },
  viewArgs(P){ return {path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P), granularity: P.st.gran, detail: P.st.detail || null}; },
  cascade: [["circuits", "circuit_options"], ["poles", "pole_options"]],
  filters(P){ const L = P.st.L, o = L.options;
    return {dateFields: L.date_fields, dateKey: "date_field", presets: L.presets,
      multis: [{key: "districts", label: "District", options: o.district}, {key: "projects", label: "Project", options: o.project},
               {key: "pids", label: "PID", options: o.pid}, {key: "sourcefiles", label: "Source file", options: o.sourcefile},
               {key: "circuits", label: "Circuit", cascade: "circuit_options", hint: "Narrowed to circuits that exist given the filters above."},
               {key: "poles", label: "Pole", cascade: "pole_options", hint: "Narrowed to poles that exist given the filters above (including Circuit)."}],
      caption: v => capHTML([v.date_from ? [fmtDate(v.date_from) + " → " + fmtDate(v.date_to)] : "", [intf(v.rows), `of ${intf(v.rows_total)} rows`]])}; },
  tabs: [
    {id: "overview", label: "Overview", icon: "trend", render: ntOverview},
    {id: "jobs", label: "Jobs", icon: "cal", render: ntJobs},
    {id: "items", label: "Mapped Items", icon: "grid", render: ntItems},
    {id: "forecast", label: "Pole Position", icon: "chart", render: ntForecast},
    {id: "totals", label: "Totals", icon: "pound", render: ntTotals},
  ],
};
const periodLabel = (iso, gran) => { const [y, m, d] = iso.split("-"); return gran === "Month" ? `${MONTHS3[+m - 1]} ${y}` : `${+d} ${MONTHS3[+m - 1]}${gran === "Week" ? "" : ""} ${gran === "Day" || gran === "Week" ? y.slice(2) : ""}`.trim(); };
function ntOverview(P, panel){
  const o = P.v.overview;
  panel.append(kpis([{label: "Total CV7_recover count", value: intf(o.recover_total), hero: true}, {label: "Total poles (all categories)", value: intf(o.pole_total)},
    {label: "Rows after filters", value: intf(P.v.rows), note: `of ${intf(P.v.rows_total)}`}]));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const gsel = segc(["Day", "Week", "Month"], P.st.gran, x => { P.st.gran = x; P.refresh(); });
  if (!o.trend){ const c = dcard("CV7_recover — count over time", "", 7); c.acts.append(gsel); c.body.append(emptyMsg("No CV7_recover records (with a date) for the current filters.")); g.append(c); }
  else { const labs = o.trend.x.map(x => periodLabel(x, P.st.gran));
    g.append(vizCard({title: "CV7_recover — count over time", sub: `Grouped by ${P.st.gran.toLowerCase()}`, span: 7, extra: gsel,
      draw: host => Viz.line(host, {labels: labs, values: o.trend.y, name: "CV7_recover", height: 300}),
      table: () => ({name: "CV7_recover over time", columns: ["Date", "Count"], rows: o.trend.x.map((x, i) => [x, o.trend.y[i]])})})); }
  if (!o.poles.length){ const c = dcard("All pole categories (erect only)", "", 5); c.body.append(emptyMsg("No pole records for the current filters.")); g.append(c); }
  else g.append(vizCard({title: "All pole categories (erect only)", sub: "Count by pole type", span: 5,
    draw: host => Viz.bars(host, {labels: o.poles.map(p => p.type), series: [{name: "Count", values: o.poles.map(p => p.count), color: "var(--s1)"}], fmt: intf, labelMax: 200}),
    table: () => ({name: "Pole categories", columns: ["Pole type", "Count"], rows: o.poles.map(p => [p.type, p.count])})}));
}
const GROUP_ICON = {"Poles": "pole", "Transformers": "tx", "Conductor": "cable", "Switch gear": "switch"};
function ntItems(P, panel){
  const it = P.v.items;
  if (!it.detail_names.length){ panel.append(note("No mapped items for the current filters.")); return; }
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const chosen = it.detail && it.detail.chosen;
  const cardEl = c => { const b = h(`<button class="icard${c.name === chosen ? " on" : ""}"><small></small><b></b></button>`); $("small", b).textContent = c.label; $("b", b).textContent = c.value;
    b.title = "Show the rows behind " + c.label; b.onclick = () => { P.st.detail = c.name; P.refresh(); }; return b; };
  it.groups.forEach(gr => {
    const c = dcard(gr.title, "", 6); c.querySelector(".ch").remove();
    const w = h(`<div class="igrp"><div class="ih"><div class="ico">${ICON2[GROUP_ICON[gr.title]] || I.box3}</div><div><h3></h3><p></p></div></div><div class="icards"></div></div>`);
    $("h3", w).textContent = gr.title; $("p", w).textContent = gr.cards.length ? `${gr.cards.length} item type(s)` : "No records for this group under the current filters.";
    gr.cards.forEach(x => $(".icards", w).append(cardEl(x)));
    c.body.append(w); g.append(c);
  });
  if (it.other.length){ const c = dcard("Other items", "", 12); const w = h(`<div class="icards"></div>`); it.other.forEach(x => w.append(cardEl(x))); c.body.append(w); g.append(c); }
  if (it.detail){
    const c = dcard("Details", "The rows behind the selected item – click any card above, or pick here", 12);
    const sel = h(`<select></select>`); it.detail_names.forEach(n => { const o = document.createElement("option"); o.textContent = n; sel.append(o); }); sel.value = chosen;
    sel.onchange = () => { P.st.detail = sel.value; P.refresh(); };
    c.acts.append(sel); c.body.append(dtable(tbl(it.detail, "Details - " + chosen, {height: 380}))); g.append(c);
  }
}
function ntTotals(P, panel){
  const t = P.v.totals;
  if (t.total == null){ panel.append(note("The mapped Total column isn't in the file – pick it under Column mapping.", "warn", "warn")); return; }
  const items = [{label: "Total value", value: money(t.total), hero: true}];
  if (t.orig != null){ const d = t.total - t.orig; items.push({label: "Difference vs original", value: signed(d), tone: d >= 0 ? "up" : "down", note: d >= 0 ? "▲ above the original" : "▼ below the original"}, {label: "Original value", value: money(t.orig)}); }
  panel.append(kpis(items));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  if (t.by_project){ const r = t.by_project.rows;
    g.append(vizCard({title: "Total value by project", sub: `${r.length} project(s) under the current filters`, span: 6,
      draw: host => Viz.bars(host, {labels: r.map(x => String(x[0])), series: [{name: "Total value", values: r.map(x => x[1]), color: "var(--s1)"}], fmt: money, axis: moneyS, valueLabel: i => moneyS(r[i][1])}),
      table: () => tbl(t.by_project, "Total value by project")})); }
  if (t.variance_by_project){ const r = t.variance_by_project.rows;
    g.append(vizCard({title: "Difference by project", sub: "Total − original", span: 6,
      draw: host => Viz.diverging(host, {labels: r.map(x => String(x[0])), values: r.map(x => x[1]), fmt: money, axis: moneyS, valueLabel: i => signed(r[i][1], moneyS),
        tip: i => ({title: String(r[i][0]), rows: [{color: r[i][1] >= 0 ? "var(--s1)" : "var(--s8)", name: "Difference", value: signed(r[i][1])}]})}),
      table: () => tbl(t.variance_by_project, "Difference by project")})); }
  if (t.variance) g.append(tableCard("Jobs where total ≠ original", `${intf(t.variance.rows.length)} rows with a variance under the current filters`, tbl(t.variance, "Jobs where total differs from original")));
  if (t.by_job){ const r = t.by_job.rows;
    if (r.length && r.length <= 60) g.append(vizCard({title: "Difference by job", sub: `${r.length} job(s), largest difference first`, span: 12,
      draw: host => { const sc = h(`<div class="scrollv"></div>`); host.append(sc); Viz.diverging(sc, {labels: r.map(x => String(x[0])), values: r.map(x => x[1]), fmt: money, axis: moneyS, labelMax: 360, valueLabel: i => signed(r[i][1], moneyS),
        tip: i => ({title: String(r[i][0]), rows: [{color: r[i][1] >= 0 ? "var(--s1)" : "var(--s8)", name: "Difference", value: signed(r[i][1])}]})}); },
      table: () => tbl(t.by_job, "Difference by job")}));
    else g.append(tableCard("Difference by job", `${intf(r.length)} job(s)`, tbl(t.by_job, "Difference by job")));
  }
}
/* ---- Jobs: Outages Programme calendar + .ics ---- */
function ntJobs(P, panel){
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const oc = dcard("Outages Programme 2026", "High-level_planning_2026.xlsx, sheet '2026' · click a day to see its outages and the poles & MD Poling tasks from the Master file", 12);
  g.append(oc);
  if (!FILES.outages){ oc.body.append(slotZone("outages", false, "Drop High-level_planning_2026.xlsx – the calendar shows here")); }
  else ntOutages(P, oc);
  const ic = dcard("Outage calendar (.ics)", "Events from the outage calendar file", 12); g.append(ic);
  if (!FILES.ics){ ic.body.append(slotZone("ics", false, "Drop outage_calendar_2026 (vN).ics")); }
  else { ic.body.append(spinner("Reading the calendar…"));
    P.tabCall("nt_ics", {ics_path: FILES.ics.path}).then(r => { ic.body.textContent = ""; $("p", ic).textContent = `📅 Using: ${r.name} · ${intf(r.count)} event(s)`;
      ic.body.append(r.count ? dtable({columns: r.columns, rows: r.rows, name: "Outage calendar"}) : emptyMsg("No VEVENT entries found in this calendar file.")); })
      .catch(e => { if (!e.stale){ ic.body.textContent = ""; ic.body.append(errNote(e)); } }); }
}
function ntOutages(P, card){
  const o = P.st.o, body = card.body;
  body.append(spinner("Reading the Outages Programme…"));
  P.tabCall("nt_outages", {outage_path: FILES.outages.path, master_path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P),
    o_districts: o.districts || [], o_pms: o.pms || [], o_from: o.from || null, o_to: o.to || null, selected_date: o.date || null}, NEED_NT)
  .then(r => {
    body.textContent = "";
    const redo = () => { body.textContent = ""; ntOutages(P, card); };
    const bar = h(`<div class="fr" style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px"></div>`);
    bar.append(msel({label: "District", options: r.districts, value: o.districts || [], onChange: v => { o.districts = v; redo(); }}),
               msel({label: "SPEN PM", options: r.pms, value: o.pms || [], onChange: v => { o.pms = v; redo(); }}));
    const di = (k, def) => { const i = h(`<input type="date">`); i.value = o[k] || def || ""; i.onchange = () => { o[k] = i.value || null; redo(); }; return i; };
    bar.append(h(`<span class="lbl" style="font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;margin-left:6px">Outage dates</span>`), di("from", r.min), h(`<span class="muted">→</span>`), di("to", r.max));
    bar.append(segc([["calendar", "Calendar"], ["table", "Table"]], o.view || "calendar", v => { o.view = v; redo(); }));
    const cnt = h(`<span class="muted" style="margin-left:auto;font-size:13px"></span>`); cnt.textContent = `${intf(r.rows)} of ${intf(r.rows_total)} outage rows`; bar.append(cnt);
    body.append(bar);
    if ((o.view || "calendar") === "table"){ body.append(dtable({columns: r.table.columns, rows: r.table.rows, name: "Outages Programme", height: 460})); return; }
    const districts = [...new Set(r.events.map(e => e.district).filter(Boolean))].sort();
    const col = P.ck.assign(districts);
    if (districts.length > 1) body.append(legendEl(districts.map(d => ({name: d, color: col(d)}))));
    body.append(calendarEl(r, o, col, iso => { o.date = iso; redo(); }, m => { o.month = m; redo(); }));
    if (r.day) body.append(dayPanel(r.day));
    else body.append(h(`<p class="muted" style="margin:12px 0 0">Click a date or an outage above to see its full breakdown and the poles & tasks here.</p>`));
  })
  .catch(e => { if (!e.stale){ body.textContent = ""; body.append(errNote(e)); } });
}
function calendarEl(r, o, col, onDay, onMonth){
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const today = iso(new Date());
  let month = o.month;
  if (!month){ const t = today.slice(0, 7); month = (r.min && r.max && today >= r.min && today <= r.max) ? t : (r.events[0] ? [...r.events].sort((a, b) => a.start < b.start ? -1 : 1)[0].start.slice(0, 7) : t); }
  const [yy, mm] = month.split("-").map(Number);
  const first = new Date(yy, mm - 1, 1), start = new Date(first); start.setDate(1 - ((first.getDay() + 6) % 7));   // weeks start Monday
  const byDay = new Map(); r.events.forEach(e => { if (!byDay.has(e.start)) byDay.set(e.start, []); byDay.get(e.start).push(e); });
  const el = h(`<div class="cal"><div class="calh"><button class="btn sm" title="Previous month">‹</button><h4></h4><button class="btn sm">Today</button><button class="btn sm" title="Next month">›</button></div><div class="calg"></div></div>`);
  $("h4", el).textContent = `${["January","February","March","April","May","June","July","August","September","October","November","December"][mm - 1]} ${yy}`;
  const shift = n => { const d = new Date(yy, mm - 1 + n, 1); onMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`); };
  const [bp, bt, bn] = el.querySelectorAll(".calh .btn");
  bp.onclick = () => shift(-1); bn.onclick = () => shift(1); bt.onclick = () => onMonth(today.slice(0, 7));
  const grid = $(".calg", el);
  ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].forEach(d => { const x = h(`<div class="dow"></div>`); x.textContent = d; grid.append(x); });
  const weeks = Math.ceil((((first.getDay() + 6) % 7) + new Date(yy, mm, 0).getDate()) / 7);
  for (let i = 0; i < weeks * 7; i++){
    const d = new Date(start); d.setDate(start.getDate() + i); const k = iso(d), evs = byDay.get(k) || [];
    const cell = h(`<div class="day${d.getMonth() !== mm - 1 ? " oth" : ""}${k === today ? " today" : ""}${k === o.date ? " sel" : ""}"><div class="dn"><span></span><span class="muted"></span></div></div>`);
    $(".dn span", cell).textContent = d.getDate(); if (evs.length) $(".dn .muted", cell).textContent = evs.length;
    evs.slice(0, 3).forEach(e => { const ev = h(`<div class="ev"><i></i><span></span></div>`); $("i", ev).style.background = col(e.district); $("span", ev).textContent = e.title.replace(/^[^—]*—\s*/, ""); ev.title = e.title; cell.append(ev); });
    if (evs.length > 3){ const m = h(`<div class="more"></div>`); m.textContent = `+${evs.length - 3} more`; cell.append(m); }
    cell.title = evs.length ? evs.map(e => e.title).join("\n") : "No outages";
    cell.onclick = () => onDay(k);
    grid.append(cell);
  }
  return el;
}
function dayPanel(day){
  const w = h(`<div class="dgrid" style="margin-top:14px"></div>`);
  const a = dcard(`Outages on ${fmtDate(day.date)}`, "", 12);
  a.body.append(day.outages.rows.length ? dtable({columns: day.outages.columns, rows: day.outages.rows, name: "Outages " + day.date, search: false}) : emptyMsg("No outages on this date under the current filters."));
  const b = dcard(`Poles & MD Poling tasks on ${fmtDate(day.date)}`, "From the Master file, under the filters above", 12);
  if (day.note) b.body.append(note(day.note));
  else if (!day.tasks || !day.tasks.rows.length) b.body.append(emptyMsg("No pole/task records for this date in the main dataset under the current filters."));
  else { $("p", b).textContent = `${intf(day.pole_count)} pole(s) · ${intf(day.tasks.rows.length)} task row(s)`; b.body.append(poleTaskTable(day.tasks)); }
  w.append(a, b); return w;
}
/* pole cell merged across that pole's task rows (same as the app's HTML table) */
function poleTaskTable(t){
  const el = h(`<div class="dt"><div class="dth"><span class="cnt"></span><button class="btn sm">${ICON2.csv} CSV</button></div><div class="tw ptt"><table><thead><tr><th>Pole</th><th>Task</th><th class="nr">Qty</th><th>Erect</th></tr></thead><tbody></tbody></table></div></div>`);
  $(".cnt", el).textContent = `${intf(t.rows.length)} rows`;
  $(".dth .btn", el).onclick = () => downloadCSV("Pole tasks", t.columns, t.rows);
  const rows = t.rows, out = [];
  for (let i = 0; i < rows.length;){
    let j = i; while (j + 1 < rows.length && rows[j + 1][0] === rows[i][0]) j++;
    for (let k = i; k <= j; k++){ const r = rows[k];
      out.push(`<tr>${k === i ? `<td class="pole" rowspan="${j - i + 1}">${esc(r[0])}</td>` : ""}<td title="${esc(r[1])}">${esc(r[1])}</td><td class="nr">${esc(qty(r[2]))}</td><td class="${r[3] === "Yes" ? "erect" : ""}">${r[3] === "Yes" ? "✔ Yes" : ""}</td></tr>`); }
    i = j + 1;
  }
  $("tbody", el).innerHTML = out.join("");
  return el;
}
/* ---- Pole Position (Service Partner Workbank): Gantt by District → Voltage → Project → Circuit / PID ---- */
function ntForecast(P, panel){
  if (!FILES.forecast){ const c = dcard("Pole Position", "Poles disposed vs. forecasted, from the Service Partner Workbank", 12); c.body.append(slotZone("forecast", false, "Drop the Service Partner Workbank (.xlsx)")); panel.append(c); return; }
  const fc = P.st.fc;
  panel.append(spinner("Reading the workbook…"));
  P.tabCall("nt_forecast", {path: FILES.forecast.path, sheet: fc.sheet || null, f_cols: fc.f_cols || null, districts: fc.districts || [], voltages: fc.voltages || [], years: fc.years || [], statuses: fc.statuses || []}, NEED_NT)
  .then(r => {
    panel.querySelectorAll(".dspin").forEach(n => n.remove());
    const redo = () => P.renderTab();
    fc.sheet = r.sheet;
    const bar = h(`<div class="fr" style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px"><span class="lbl" style="font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase">Sheet</span></div>`);
    const ss = h(`<select></select>`); r.sheets.forEach(s => { const o = document.createElement("option"); o.textContent = s; ss.append(o); }); ss.value = r.sheet;
    ss.onchange = () => { fc.sheet = ss.value; fc.f_cols = null; fc.districts = fc.voltages = fc.years = fc.statuses = []; fc.pick = null; redo(); };
    bar.append(ss);
    if (r.options){
      if (r.options.district.length) bar.append(msel({label: "District", options: r.options.district, value: fc.districts || [], onChange: v => { fc.districts = v; redo(); }}));
      if (r.options.voltage.length) bar.append(msel({label: "Voltage", options: r.options.voltage, value: fc.voltages || [], onChange: v => { fc.voltages = v; redo(); }}));
      if (r.options.status.length) bar.append(msel({label: "Status", options: r.options.status, value: fc.statuses || [], onChange: v => { fc.statuses = v; redo(); }}));
      if (r.options.year.length) bar.append(msel({label: "Start year", options: r.options.year.map(String), value: (fc.years || []).map(String), onChange: v => { fc.years = v.map(Number); redo(); }}));
    }
    panel.append(bar);
    const d = h(`<details class="dmap"><summary>⚙ Column mapping</summary><div class="frm"></div></details>`);
    if (r.missing.length) d.open = true;
    r.fields.forEach(fl => { const fe = h(`<div class="field${r.f_cols[fl.key] ? "" : " miss"}"><label></label><select></select></div>`); $("label", fe).textContent = fl.label;
      const sel = $("select", fe); [["", "(none)"], ...r.columns.map(c => [c, c])].forEach(([v, t]) => { const o = document.createElement("option"); o.value = v; o.textContent = t; sel.append(o); });
      sel.value = r.f_cols[fl.key] || ""; sel.onchange = () => { fc.f_cols = {...r.f_cols, [fl.key]: sel.value || null}; redo(); }; $(".frm", d).append(fe); });
    panel.append(d);
    if (r.missing.length){ panel.append(note(`Couldn't guess a column for: ${r.missing.map(k => (r.fields.find(x => x.key === k) || {}).label).join(", ")} – pick them under Column mapping.`, "warn", "warn")); return; }
    const b = r.banner, jobs = r.jobs;
    const k = kpis([{label: "Poles disposed vs forecasted", value: `${intf(b.disposed)} / ${intf(b.forecast)}`, hero: true, note: b.pct != null ? `${Math.round(b.pct * 100)}% disposed` : ""},
                    {label: "Remaining to dispose", value: intf(Math.max(0, b.forecast - b.disposed))}, {label: "Jobs (project / circuit rows)", value: intf(jobs.length)}]);
    if (b.pct != null){ const m = h(`<div class="meter" style="margin-top:6px"><i></i></div>`); $("i", m).style.width = Math.min(100, b.pct * 100) + "%"; $(".kpi.hero", k).append(m); }
    panel.append(k);
    const missingCols = ["status", "start_date", "finish_date", "comment", "control_file"].filter(x => !r.f_cols[x]);
    if (missingCols.length) panel.append(note(`Not found in this sheet: ${missingCols.map(x => (r.fields.find(f => f.key === x) || {}).label).join(", ")}. Pick them under Column mapping if they're there under another name.`, "warn", "warn"));
    if (!jobs.length){ panel.append(note("No rows to chart for the current filters.")); return; }
    const g = h(`<div class="dgrid"></div>`); panel.append(g);
    const main = dcard("Poles disposed vs. forecasted", "District → Voltage → Project → Circuit / PID · each bar runs from Start Date to Finish Date, coloured by Status · click a job for its details", 8);
    const side = h(`<section class="dcard span4 jobpanel"></section>`);
    g.append(main, side);
    let mode = "gantt"; const holder = h(`<div></div>`);
    main.acts.append(segc([["gantt", "Gantt"], ["table", "Table"]], mode, m => { mode = m; draw(); }));
    main.body.append(holder);
    if (fc.pick == null || !jobs.some(j => j.row === fc.pick)) fc.pick = jobs[0].row;
    const pick = row => { fc.pick = row; holder.querySelectorAll(".gjob").forEach(x => x.classList.toggle("on", +x.dataset.row === row)); jobPanel(side, jobs.find(j => j.row === row)); };
    function draw(){
      holder.textContent = "";
      if (mode === "table"){ holder.append(dtable({name: "Pole position", columns: ["District", "Voltage", "Project", "Circuit", "PID", "Status", "Start Date", "Finish Date", "Forecast", "Disposed", "Remaining", "Comment", "Control File", "Project link"],
        rows: jobs.map(j => [j.district, j.voltage, j.project, j.circuit, j.pid, j.status, j.start, j.finish, j.forecast, j.disposed, j.remaining, j.comment, j.control_file, j.project_link])})); return; }
      const counts = new Map(); jobs.forEach(j => { const key = statusKey(r.statuses, j); counts.set(key.label, (counts.get(key.label) || 0) + 1); });
      holder.append(legendEl(r.statuses.filter(s => counts.has(s.label)).map(s => ({name: `${s.label} (${counts.get(s.label)})`, color: s.colour}))));
      holder.append(ganttEl(jobs, fc.pick, pick));
    }
    draw(); jobPanel(side, jobs.find(j => j.row === fc.pick));
  })
  .catch(e => { if (!e.stale){ panel.querySelectorAll(".dspin").forEach(n => n.remove()); panel.append(errNote(e)); } });
}
function statusKey(statuses, j){ const s = (j.status || "").toLowerCase();
  const map = [["complete", 0], ["planned", 1], ["awaiting outage plan", 2], ["land access", 3], ["still to be handed over", 4], ["to be priced", 5]];
  const hit = map.find(([k]) => s.includes(k)); return statuses[hit ? hit[1] : statuses.length - 1]; }
const inkOn = hex => { const m = /^#?([0-9a-f]{6})$/i.exec(hex || ""); if (!m) return "#fff"; const n = parseInt(m[1], 16), r = n >> 16 & 255, g = n >> 8 & 255, b = n & 255;
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 > 0.55 ? "#0b0b0b" : "#ffffff"; };
const dayMs = 864e5, toDay = s => s ? Date.parse(s.slice(0, 10) + "T00:00:00Z") : NaN;
function ganttEl(jobs, picked, onPick){
  const starts = jobs.map(j => toDay(j.start)).filter(isFinite), ends = jobs.map(j => toDay(j.finish || j.start)).filter(isFinite);
  const el = h(`<div class="gantt"><div class="gh"><div class="glab"></div><div class="gaxis"></div></div><div class="gbody"></div></div>`);
  if (!starts.length){ el.querySelector(".gbody").append(emptyMsg("No Start Dates in these rows – map the Start Date column under Column mapping.")); return el; }
  let t0 = new Date(Math.min(...starts)), t1 = new Date(Math.max(...ends, ...starts));
  t0 = Date.UTC(t0.getUTCFullYear(), t0.getUTCMonth(), 1); const e1 = new Date(t1); t1 = Date.UTC(e1.getUTCFullYear(), e1.getUTCMonth() + 1, 1);
  const pct = t => ((t - t0) / (t1 - t0) * 100);
  const months = []; for (let d = new Date(t0); d.getTime() < t1; d = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1))) months.push(d.getTime());
  const step = months.length > 24 ? 3 : months.length > 14 ? 2 : 1;
  const axis = $(".gaxis", el), grid = months.map(m => `<i style="left:${pct(m).toFixed(3)}%"></i>`).join("");
  months.forEach((m, i) => { if (i % step) return; const d = new Date(m), s = h(`<span></span>`); s.style.left = pct(m) + "%"; s.textContent = `${MONTHS3[d.getUTCMonth()]}${d.getUTCMonth() === 0 || i === 0 ? " " + d.getUTCFullYear() : ""}`; axis.append(s); });
  const now = Date.now(), todayX = now >= t0 && now <= t1 ? `<b class="gtoday" style="left:${pct(now).toFixed(3)}%"></b>` : "";
  if (todayX){ const s = h(`<span class="today">Today</span>`); s.style.left = pct(now) + "%"; axis.append(s); }
  const sum = (arr, k) => arr.reduce((a, j) => a + (j[k] || 0), 0);
  const tot = arr => `${intf(sum(arr, "disposed"))} / ${intf(sum(arr, "forecast"))} poles`;
  const body = $(".gbody", el), parts = [];
  const lab = x => x || "(blank)";
  const groupBy = (arr, key) => { const m = new Map(); arr.forEach(j => { const k = lab(j[key]); if (!m.has(k)) m.set(k, []); m.get(k).push(j); }); return [...m]; };
  groupBy(jobs, "district").forEach(([dn, dj]) => {
    parts.push(`<div class="g1"><span>${esc(dn)}</span><span>${esc(tot(dj))}</span></div>`);
    groupBy(dj, "voltage").forEach(([vn, vj]) => {
      parts.push(`<div class="g2"><span>${esc(vn)}</span><span>${esc(tot(vj))}</span></div>`);
      groupBy(vj, "project").forEach(([pn, pj]) => {
        const link = pj.find(j => j.project_link);
        parts.push(`<div class="g3"><span>${link ? `<a href="${esc(link.project_link)}" target="_blank" rel="noopener" title="${esc(link.project_link)}">${esc(pn)} ${I.ext.replace("<svg", "<svg width=13 height=13")}</a>` : esc(pn)}</span><span>${esc(tot(pj))}</span></div>`);
        pj.forEach(j => {
          const s = toDay(j.start), f = toDay(j.finish);
          let barH = "";
          if (isFinite(s)){
            const e = isFinite(f) && f >= s ? f + dayMs : s + 21 * dayMs;   // finish day included; no finish date -> short open-ended bar
            const left = pct(s), w = Math.max(0.6, pct(e) - left), ink = inkOn(j.colour), prog = j.forecast ? Math.min(100, j.disposed / j.forecast * 100) : 0;
            barH = `<div class="gbar${isFinite(f) ? "" : " open"}" title="${isFinite(f) ? "" : "No Finish Date in the workbook"}" style="left:${left.toFixed(3)}%;width:${w.toFixed(3)}%;background:${j.colour};color:${ink}"><em style="width:${prog.toFixed(1)}%;background:${ink}"></em><span data-long="Disposed ${intf(j.disposed)} · Forecast ${intf(j.forecast)}" data-short="${intf(j.disposed)} / ${intf(j.forecast)}">Disposed ${intf(j.disposed)} · Forecast ${intf(j.forecast)}</span></div>`;
          } else barH = `<div class="gnodate">No Start Date · ${intf(j.disposed)} / ${intf(j.forecast)} poles</div>`;
          parts.push(`<div class="gjob${j.row === picked ? " on" : ""}" data-row="${j.row}"><div class="glab"><b>${esc(j.circuit || "—")}</b><small>${esc(j.pid ? "PID " + j.pid : "")}</small></div><div class="gtrack">${grid}${todayX}${barH}</div></div>`);
        });
      });
    });
  });
  body.innerHTML = parts.join("");
  const fit = () => body.querySelectorAll(".gbar").forEach(bar => { const sp = $("span", bar); sp.textContent = sp.dataset.long; bar.classList.remove("outside");
    if (sp.scrollWidth > bar.clientWidth - 10){ sp.textContent = sp.dataset.short; if (sp.scrollWidth > bar.clientWidth - 10) bar.classList.add("outside"); } });
  new ResizeObserver(() => requestAnimationFrame(fit)).observe(body);
  body.addEventListener("click", e => { if (e.target.closest("a")) return; const r = e.target.closest(".gjob"); if (r) onPick(+r.dataset.row); });
  body.addEventListener("mousemove", e => { const r = e.target.closest(".gjob"); if (!r){ Tip.hide(); return; } const j = jobs.find(x => x.row === +r.dataset.row);
    Tip.show(e, {title: `${j.project} — ${j.circuit}`, rows: [{color: j.colour, name: "Status", value: j.status || "—"}, {name: "Start → Finish", value: `${fmtDate(j.start) || "—"} → ${fmtDate(j.finish) || "—"}`},
      {name: "Disposed / Forecast", value: `${intf(j.disposed)} / ${intf(j.forecast)}`}], note: "Click for the comment, control file and link"}); });
  body.addEventListener("mouseleave", () => Tip.hide());
  return el;
}
function linkish(v){ return /^(https?:|mailto:|file:)/i.test(v || "") ? v : /^\\\\/.test(v || "") ? "file:" + v.replace(/\\/g, "/") : ""; }
function jobPanel(side, j){
  side.textContent = "";
  if (!j){ side.append(emptyMsg("Click a job to see its details.")); return; }
  const w = h(`<div class="jp"><div class="jph"><small>Job details</small><h3></h3><div class="jst"><i></i><span></span></div></div><dl></dl><div class="jpc"><h4>Comment</h4><p></p></div><div class="jpc jcf"><h4>Control File</h4><div></div></div></div>`);
  const hh = $("h3", w);
  if (j.project_link){ const a = h(`<a target="_blank" rel="noopener"></a>`); a.href = j.project_link; a.title = j.project_link; a.textContent = j.project + " ↗"; hh.append(a); } else hh.textContent = j.project;
  $(".jst i", w).style.background = j.colour; $(".jst span", w).textContent = j.status || "No status";
  const dl = $("dl", w);
  [["District", j.district], ["Voltage", j.voltage], ["Circuit", j.circuit], ["PID", j.pid], ["Start Date", fmtDate(j.start) || "—"], ["Finish Date", fmtDate(j.finish) || "—"],
   ["Forecasted Total poles", intf(j.forecast)], ["Poles Disposed", intf(j.disposed)], ["Remaining", intf(j.remaining)], ["Workbook row", String(j.row)]].forEach(([k, v]) => {
    const dt = document.createElement("dt"), dd = document.createElement("dd"); dt.textContent = k; dd.textContent = v || "—"; dl.append(dt, dd); });
  const pm = h(`<div class="meter" style="margin:4px 0 12px"><i></i></div>`); $("i", pm).style.width = (j.forecast ? Math.min(100, j.disposed / j.forecast * 100) : 0) + "%"; dl.after(pm);
  $(".jpc p", w).textContent = j.comment || "No comment in column L for this job.";
  const cf = $(".jcf div", w);
  if (!j.control_file && !j.control_link) cf.textContent = "No control file in column O.";
  else {
    const target = j.control_link || j.control_file, href = linkish(target);
    const line = h(`<div class="jcfl"></div>`);
    if (href){ const a = h(`<a target="_blank" rel="noopener"></a>`); a.href = href; a.textContent = j.control_file || target; a.title = target; line.append(a); }
    else { const s = h(`<span></span>`); s.textContent = j.control_file; line.append(s); }
    const cp = h(`<button class="btn sm" title="Copy to the clipboard">Copy</button>`);
    cp.onclick = () => { navigator.clipboard && navigator.clipboard.writeText(target).then(() => { cp.textContent = "Copied ✔"; setTimeout(() => cp.textContent = "Copy", 1500); }); };
    line.append(cp); cf.append(line);
    if (/^\\\\/.test(target)) cf.append(h(`<p class="muted" style="font-size:12px;margin:6px 0 0">Network paths may not open straight from the browser – use Copy and paste it into File Explorer.</p>`));
  }
  if (j.project_link){ const b = h(`<a class="btn sm" target="_blank" rel="noopener" style="margin-top:12px">${I.ext} Open project link</a>`); b.href = j.project_link; w.append(b); }
  side.append(w);
}

/* ======================================================================
   Materials Breakdown
   ====================================================================== */
DASHDEF.materials = {
  slots: [["materials", true], ["master", true]], loadSlots: ["materials", "master"], needs: NEED_MB,
  about: ["materials_all.parquet (Programs\\Program files) and the Master_DD-MM-YYYY.parquet control file.",
          "Maps are optional – drop Outages Programme folders or PDF maps on the Maps tab."],
  slotNotes: {master: "The control file: Master_DD-MM-YYYY (vN).parquet"},
  loadFn: "mb_load", loadArgs: () => ({master_path: FILES.materials.path, control_path: FILES.master.path}),
  init(P){ const f = P.st.f, L = P.st.L; if (!L.date_fields.includes(f.date_field_label)) f.date_field_label = L.default_date_field || "DateToUse";
    P.st.scope = P.st.scope || "All poles matching filters"; P.st.pdf = P.st.pdf || []; P.st.mapsMode = P.st.mapsMode || "Manual folder(s)"; },
  viewFn: "mb_view",
  mbFilters(P){ const f = P.st.f; return {date_field_label: f.date_field_label, date_from: f.date_from || null, date_to: f.date_to || null, districts: f.districts || [], projects: f.projects || [],
    pids: f.pids || [], sourcefiles: f.sourcefiles || [], circuits: f.circuits || [], enids: f.enids || []}; },
  viewArgs(P){ const s = P.st; return {master_path: FILES.materials.path, control_path: FILES.master.path, filters: DASHDEF.materials.mbFilters(P),
    fi_cats: s.fi_cats || [], fi_mats: s.fi_mats || [], guk_cats: s.guk_cats || [], guk_mats: s.guk_mats || [], guk_submats: s.guk_submats || []}; },
  cascade: [["circuits", "circuit_options"], ["enids", "pole_options"]],
  filters(P){ const L = P.st.L, o = L.options;
    return {dateFields: L.date_fields, dateKey: "date_field_label", fromTo: ["date_from", "date_to"],
      multis: [{key: "districts", label: "District", options: o.district}, {key: "projects", label: "Project", options: o.project},
               {key: "pids", label: "PID", options: o.pid}, {key: "sourcefiles", label: "Source file", options: o.sourcefile},
               {key: "circuits", label: "Circuit", cascade: "circuit_options", hint: "Narrowed to circuits that exist given the District/Project/Date selected."},
               {key: "enids", label: "Pole (enid)", cascade: "pole_options", hint: "Narrowed to poles that exist given District/Project/Circuit/Date. Leave empty to include all."}],
      onReset: () => { ["fi_cats", "fi_mats", "guk_cats", "guk_mats", "guk_submats"].forEach(k => P.st[k] = []); },
      caption: v => capHTML([[intf(v.rows), "control-file rows matched"], [intf(v.mat_rows), "material rows"]])}; },
  header(P, panel){ const k = P.v.kpis;
    panel.append(kpis([{label: "Poles", value: intf(k.poles), hero: true}, {label: "Free Issue line items", value: intf(k.fi)}, {label: "GUK assemblies", value: intf(k.guk)},
      {label: "GUK components", value: intf(k.guk_sub)}, {label: "Control-file rows matched", value: intf(P.v.rows), note: `${intf(P.v.mat_rows)} material rows`}])); },
  tabs: [
    {id: "poles", label: "Poles", icon: "pole", render: mbPoles},
    {id: "fi", label: "Free Issue", icon: "box3", render: mbFI},
    {id: "guk", label: "GUK", icon: "layers", render: mbGUK},
    {id: "maps", label: "Maps (optional)", icon: "map", render: mbMaps},
    {id: "export", label: "Export", icon: "csv", render: mbExport},
  ],
};
function mbPoles(P, panel){
  const p = P.v.poles;
  if (!p || !p.rows.length){ panel.append(note("No poles matched for the current filters.")); return; }
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  g.append(vizCard({title: "Pole quantities by size", sub: "Poles needed for the filtered Control File", span: 5,
    draw: host => Viz.columns(host, {labels: p.chart.x.map(String), height: 300, axis: intf, series: [{name: "Quantity", values: p.chart.y, color: "var(--s1)"}], fmt: qty,
      tip: i => ({title: "Pole size " + p.chart.x[i], rows: [{color: "var(--s1)", name: "Quantity", value: qty(p.chart.y[i])}]})}),
    table: () => ({name: "Pole quantities", columns: ["Pole size", "Quantity"], rows: p.chart.x.map((x, i) => [x, p.chart.y[i]])})}));
  g.append(tableCard("Poles needed", "", tbl(p, "Poles"), 7));
}
function localFilters(items){
  const bar = h(`<div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px"><span style="font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase">Narrow</span></div>`);
  items.forEach(x => bar.append(msel(x)));
  return bar;
}
/* chart of items coloured by their Type, legend by Type */
function typedColumns(P, {title, sub, rows, xKey, desc, span = 12, name, extra}){
  const cats = [...new Set(rows.map(r => r.category ?? "(Uncategorized)"))];
  const col = P.ck.assign(cats);
  return vizCard({title, sub, span, extra, legend: cats.length > 1 ? cats.map(c => ({name: c, color: col(c)})) : null,
    draw: host => Viz.columns(host, {labels: rows.map(r => String(r[xKey] ?? "")), height: 320, axis: compact, fmt: qty,
      series: [{name: "Quantity", values: rows.map(r => r.qty), colorAt: i => col(rows[i].category ?? "(Uncategorized)")}],
      tip: i => ({title: String(rows[i][xKey] ?? ""), rows: [{color: col(rows[i].category ?? "(Uncategorized)"), name: "Type", value: rows[i].category ?? "(Uncategorized)"},
        ...(desc ? [{name: "Description", value: rows[i].description ?? ""}] : []), {name: "Quantity", value: qty(rows[i].qty)}]})}),
    table: () => ({name, columns: ["Type", xKey === "code" ? "Code" : "Item", ...(desc ? ["Description"] : []), "Quantity"], rows: rows.map(r => [r.category, r[xKey], ...(desc ? [r.description] : []), r.qty])})});
}
function mbFI(P, panel){
  const fi = P.v.fi, s = P.st;
  if (!P.v.kpis.fi && !(s.fi_cats || []).length && !(s.fi_mats || []).length){ panel.append(note("No Free Issue materials matched for the current filters.")); return; }
  panel.append(localFilters([{label: "Free Issue Type", options: fi.categories, value: s.fi_cats || [], onChange: v => { s.fi_cats = v; s.fi_mats = []; P.refresh(); }},
    {label: "Free Issue material", options: fi.materials, value: s.fi_mats || [], onChange: v => { s.fi_mats = v; P.refresh(); }}]));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  if (!fi.rows.length){ panel.append(note("Nothing matches the Type / material picked above.")); return; }
  g.append(typedColumns(P, {title: "Free Issue quantities by commodity code", sub: "Coloured by Free Issue Type", rows: fi.chart, xKey: "code", desc: true, name: "Free Issue by code"}));
  g.append(tableCard("Free Issue materials", "", tbl(fi, "Free Issue")));
}
function mbGUK(P, panel){
  const gk = P.v.guk, s = P.st;
  if (!P.v.kpis.guk && !(s.guk_cats || []).length && !(s.guk_mats || []).length){ panel.append(note("No GUK assemblies matched for the current filters.")); return; }
  panel.append(localFilters([{label: "GUK Type", options: gk.categories, value: s.guk_cats || [], onChange: v => { s.guk_cats = v; s.guk_mats = []; s.guk_submats = []; P.refresh(); }},
    {label: "GUK material (assembly)", options: gk.materials, value: s.guk_mats || [], onChange: v => { s.guk_mats = v; s.guk_submats = []; P.refresh(); }}]));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  if (gk.rows.length){
    g.append(typedColumns(P, {title: "GUK assembly quantities", sub: "Coloured by GUK Type", rows: gk.chart, xKey: "item", span: 5, name: "GUK assemblies"}));
    g.append(tableCard("GUK assemblies", "", tbl(gk, "GUK assemblies"), 7));
  } else { const e = dcard("GUK assemblies", "", 12); e.body.append(emptyMsg("Nothing matches the Type / material picked above.")); g.append(e); }
  const sub = gk.sub;
  const subSel = msel({label: "Sub GUK material", options: gk.sub_materials, value: s.guk_submats || [], right: true, onChange: v => { s.guk_submats = v; P.refresh(); }});
  if (!sub.rows.length){ const c = dcard("GUK sub-materials (component breakdown)", "Components of the assemblies in scope, with the qty-per-unit multiplier applied", 12);
    c.acts.append(subSel); c.body.append(emptyMsg("No GUK sub-materials in scope.")); g.append(c); return; }
  g.append(typedColumns(P, {title: "GUK sub-materials (component breakdown)", sub: "Quantities by sub code, with the qty-per-unit multiplier applied · coloured by GUK Type", rows: sub.chart, xKey: "code", desc: true, name: "GUK sub-materials", extra: subSel}));
  g.append(tableCard("GUK sub-materials", "", tbl(sub, "GUK sub-materials")));
}
function mbMaps(P, panel){
  const s = P.st;
  panel.append(note("Maps only ADD context (which poles were found on which map) – the Poles / Free Issue / GUK tabs always include every pole matching your filters, whether or not it's on a scanned map."));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const src = dcard("Connect maps", "", 12); g.append(src);
  src.acts.append(segc(["Manual folder(s)", "Network path scan"], s.mapsMode, v => { s.mapsMode = v; P.renderTab(); }));
  const stat = h(`<p class="status" style="margin:10px 0 0"></p>`);
  if (s.mapStat){ stat.className = s.mapStat[0]; stat.textContent = s.mapStat[1]; }
  new MutationObserver(() => { s.mapStat = [stat.className, stat.textContent]; }).observe(stat, {childList: true, characterData: true, subtree: true, attributes: true});
  const addEntries = (list, keyPrefix) => { const seen = new Set(s.pdf.map(e => e.path)); list.forEach(e => { if (!seen.has(e.path)){ seen.add(e.path); s.pdf.push(e); } }); };
  const busyRun = async (label, fn) => { stat.className = "status"; stat.textContent = label; try { await fn(); } catch(e){ if (!e.stale){ stat.className = "status err"; stat.textContent = e.message; } } };
  if (s.mapsMode === "Manual folder(s)"){
    $("p", src) || $(".tt", src).append(h(`<p></p>`));
    $(".tt p", src).textContent = "Each folder is scanned for PDFs (sub-folders too). Drop folders of maps or individual PDF files.";
    const row = h(`<div class="row"></div>`);
    row.append(folderDrop({label: "Folder(s) of PDF maps", hint: "drop one or more folders – or loose PDFs",
      onFolders: async groups => busyRun("Collecting PDFs…", async () => {
        const all = groups.flatMap(gr => gr.items.map(x => ({rel: "Manual/" + (gr.top ? x.rel : "Dropped PDFs/" + x.rel), file: x.file})));
        const tops = [...new Set(groups.map(gr => gr.top || "Dropped PDFs"))];
        const dir = await PY.mount(all);
        const r = (await PY.call("mb_folders", {folders: tops.map(t => dir + "/Manual/" + t)}, {needs: NEED_MB})).result;
        addEntries(r.entries); stat.className = "status ok"; stat.textContent = r.status; P.renderTab(); })}));
    const pdfBtn = h(`<button class="btn">${I.file} Choose PDF files…</button>`);
    pdfBtn.onclick = () => { const inp = document.createElement("input"); inp.type = "file"; inp.multiple = true; inp.accept = ".pdf";
      inp.onchange = () => busyRun("Collecting PDFs…", async () => { const fs = [...inp.files].filter(f => extOf(f.name) === ".pdf"); if (!fs.length) return;
        const dir = await PY.mount(fs.map(f => ({rel: "Manual/Dropped PDFs/" + f.name, file: f})));
        const r = (await PY.call("mb_folders", {folders: [dir + "/Manual/Dropped PDFs"]}, {needs: NEED_MB})).result; addEntries(r.entries); P.renderTab(); });
      inp.click(); };
    row.append(h(`<div class="field" style="justify-content:flex-end"></div>`)); row.lastChild.append(pdfBtn);
    src.body.append(row);
  } else {
    $("p", src) || $(".tt", src).append(h(`<p></p>`));
    $(".tt p", src).textContent = "Drop the “1 - Outages Programme” folder, a year (2026), a month (09 - September) or single outage folders. Uses the From / To dates in the filter bar – every outage folder in range is searched for “Workpack zones” folders.";
    src.body.append(folderDrop({label: "Outages Programme folder(s)", hint: "tip: drop the month folder – the whole programme works too but reads more",
      onFolders: async groups => {
        const f = s.f;
        if (!f.date_from && !f.date_to && !confirm("No From / To dates are set in the filter bar. This will scan EVERY outage folder in what you dropped.\n\nContinue anyway?")) return;
        for (const gr of groups){
          const lay = outageLayout(gr.top, gr.items);
          if (!lay){ alert(`“${gr.top || "(files)"}” doesn't look like a year, month or dated outage folder from the Outages Programme.\nUse “Manual folder(s)” for other folders.`); continue; }
          await busyRun(`Scanning ${gr.top}…`, async () => {
            const dir = await PY.mount(gr.items.map(x => ({rel: lay.prefix + x.rel, file: x.file})));
            const r = (await PY.call("mb_scan", {root: dir + lay.rootSuffix, date_from: f.date_from || "", date_to: f.date_to || ""}, {needs: NEED_MB, onLog: t => { stat.textContent = t; }})).result;
            addEntries(r.entries); stat.className = "status ok"; stat.textContent = r.status + (r.no_zone.length ? ` ${r.no_zone.length} outage folder(s) in range had no Workpack zones folder.` : "");
            if (r.no_zone.length) s.noZone = r.no_zone;
          });
        }
        P.renderTab();
      }}));
  }
  const cur = h(`<div class="row" style="margin-top:12px;align-items:center"></div>`);
  const lab = h(`<span class="chip"></span>`); lab.textContent = s.pdf.length ? `${intf(s.pdf.length)} PDF map(s) loaded` : "No maps loaded yet";
  cur.append(lab);
  if (s.pdf.length){ const clr = h(`<button class="btn sm">Clear maps</button>`); clr.onclick = () => { s.pdf = []; s.noZone = null; s.mapStat = null; P.renderTab(); }; cur.append(clr); }
  src.body.append(cur, stat);
  if (s.noZone && s.noZone.length && s.mapsMode === "Network path scan"){ const d = h(`<details style="margin-top:8px"><summary class="muted" style="cursor:pointer;font-size:13px"></summary><div class="chips" style="margin-top:6px"></div></details>`);
    $("summary", d).textContent = `${s.noZone.length} outage folder(s) in range had no Workpack zones folder`; s.noZone.forEach(n => { const c = h(`<span class="chip"><span></span></span>`); $("span", c).textContent = n; $(".chips", d).append(c); }); src.body.append(d); }
  if (!s.pdf.length) return;
  const res = h(`<div class="span12"></div>`); g.append(res);
  res.append(spinner(`Reading ${s.pdf.length} PDF(s) for pole references…`));
  P.tabCall("mb_maps", {master_path: FILES.materials.path, control_path: FILES.master.path, filters: DASHDEF.materials.mbFilters(P), pdf_entries: s.pdf.map(({path, folder, day_outage}) => ({path, folder, day_outage}))}, NEED_MB)
    .then(r => {
      res.textContent = ""; if (!r) return;
      s.mapsMatched = r.matched;
      res.append(kpis([{label: "Maps scanned", value: intf(r.maps_scanned)}, {label: "Filtered poles found on a map", value: intf(r.matched), swatch: "var(--s1)"},
        {label: "Filtered poles NOT on any map", value: intf(r.not_on_maps), swatch: "var(--neutral)", note: "their materials are still included"}]));
      const gg = h(`<div class="dgrid"></div>`); res.append(gg);
      gg.append(tableCard("Maps", "Poles found on each map", {columns: r.map_columns, rows: r.map_rows, name: "Maps"}, 7));
      gg.append(tableCard("Pole summary for your current filter", "Every pole matching the filters, flagged by whether it was found on a scanned map", {columns: r.summary_columns, rows: r.summary_rows, name: "Pole summary"}, 5));
    }).catch(e => { if (!e.stale){ res.textContent = ""; res.append(errNote(e)); } });
}
function mbExport(P, panel){
  const s = P.st;
  const c = dcard("Export current filters to Excel", "Summary (filters + PIDs), Poles, Free Issue, GUK (assemblies + subdivisions) and, when maps were scanned, Pole Summary – the same workbook as before", 12);
  panel.append(c);
  const row = h(`<div class="row" style="align-items:center;margin-bottom:14px"><span class="lbl" style="font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase">Export scope</span></div>`);
  row.append(segc(["All poles matching filters", "Only poles found on scanned maps"], s.scope, v => { s.scope = v; P.renderTab(); }));
  c.body.append(row);
  if (s.scope === "Only poles found on scanned maps" && !s.pdf.length)
    c.body.append(note('No maps have been scanned yet – go to the Maps tab first, or switch back to "All poles matching filters".', "warn", "warn"));
  const run = runArea("Download Materials_Breakdown.xlsx", "down");
  run.button.onclick = () => run.run("mb_export", {master_path: FILES.materials.path, control_path: FILES.master.path, filters: DASHDEF.materials.mbFilters(P), scope: s.scope,
    pdf_entries: s.pdf.map(({path, folder, day_outage}) => ({path, folder, day_outage}))}, {needs: NEED_MB, busy: "Building the workbook…"});
  c.body.append(run);
  const u = P.v.unmatched;
  if (u && u.rows.length){ const d = dcard(`Unmatched descriptions (${u.rows.length})`, "Not exported above – shown here for QA", 12); d.body.append(dtable(tbl(u, "Unmatched descriptions"))); panel.append(h(`<div style="height:16px"></div>`), d); }
}
