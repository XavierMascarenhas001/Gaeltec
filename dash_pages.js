/* ======================================================================
   Dashboards - each one gets its own space: coloured app bar, switcher,
   file strip, one filter bar, tabs and a grid of cards.
   ====================================================================== */
const DASHES = [
  {key: "tracker", c: "c2", icon: "grid", title: "Planning Dashboard",
   desc: "Jobs, High level planning calendar, mapped items, pole position and totals."},
  {key: "master", c: "c1", icon: "pound", title: "Financial Dashboard",
   desc: "Revenue, teams, PID breakdown and finance."},
  {key: "materials", c: "c3", icon: "box3", title: "Materials Breakdown",
   desc: "Poles, Free Issue and GUK for the filtered Control File."},
];
const DST = {};            // per-dashboard state, kept while you switch between them
let PYSTAT = ["Starting Python…", "busy"];

function pyPill(){ return h(`<span class="pill pystat"><span class="dot ${PYSTAT[1] === "on" ? "on" : PYSTAT[1] === "off" ? "off" : "busy"}"></span><span>${esc(PYSTAT[0])}</span></span>`); }
function dashBar(key){
  const d = DASHES.find(x => x.key === key);
  const el = h(`<div class="dbar"><div class="nav"><button class="btn sm">${I.back} All tools</button><div class="dsw"></div><span class="spacer"></span></div>
    <div class="dtitle"><div class="ico">${I[d.icon]}</div><div><h1></h1><p></p></div><span class="spacer"></span><span class="pill" title="Page version">v 12-Nov-2026</span></div></div>`);
  $("h1", el).textContent = d.title; $(".dtitle p", el).textContent = d.desc;
  $(".nav .btn", el).onclick = () => { location.hash = ""; };
  const sw = $(".dsw", el);
  DASHES.forEach(x => { const b = h(`<button class="${x.key === key ? "on" : ""}">${I[x.icon]}<span></span></button>`); $("span", b).textContent = x.title;
    b.onclick = () => { location.hash = "dash/" + x.key; }; sw.append(b); });
  const nav = $(".nav", el);
  const tb = h(`<button class="btn sm themebtn" title="Colour theme"></button>`); tb.textContent = Theme.label();
  tb.onclick = () => { Theme.next(); tb.textContent = Theme.label(); };
  nav.append(guideBtn("dash/" + key, true), pyPill(), tb);
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
    if (def.loadSlots.includes(slot)) P.start(); else if (P.v && (def.viewSlots || []).includes(slot)) P.refresh(); else if (P.v) P.renderTab();
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
    const tabs = h(`<div class="dtabs" role="tablist"></div>`);
    def.tabs.forEach(t => { const b = tabBtn(t.icon, t.label); b.classList.toggle("on", t.id === st.tab);
      b.onclick = () => { st.tab = t.id; tabs.querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b)); P.showTab(); }; tabs.append(b); });
    P.panel = h(`<div class="dpanel"></div>`);
    P.body.append(tabs, P.fbar, P.panel);
    P.refresh();
  };
  /* dashboards with lazyTabs work out only the tab on screen; each tab's result is
     kept until a filter changes, so going back to a tab is instant */
  P.sig = () => { const a = {...def.viewArgs(P)}; delete a.tab; return JSON.stringify(a); };
  P.showTab = () => {
    if (!def.lazyTabs) return P.renderTab();
    const hit = P.vc && P.vcSig === P.sig() && P.vc[st.tab];
    if (hit){ P.v = hit; P.renderTab(); } else P.refresh();
  };
  P.refresh = async () => {
    const my = ++P.seq;
    P.panel.classList.add("loading"); P.panel.querySelectorAll(".dspin").forEach(n => n.remove()); P.panel.prepend(spinner("Updating…"));
    try {
      const args = def.viewArgs(P), sig = P.sig(), tab = st.tab;
      const v = (await PY.call(def.viewFn, args, {needs: def.needs})).result;
      if (my !== P.seq || !P.alive()) return;
      P.v = v;
      if (def.lazyTabs){ if (P.vcSig !== sig){ P.vc = {}; P.vcSig = sig; } P.vc[tab] = v; }
      /* cascading filters: drop picks that the other filters have made impossible */
      let dropped = false;
      (def.cascade || []).forEach(([k, optKey]) => { const ok = new Set((v.options && v.options[k]) || v[optKey] || []); const keep = (st.f[k] || []).filter(x => ok.has(x)); if (keep.length !== (st.f[k] || []).length){ st.f[k] = keep; dropped = true; } });
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
    if (P.v.plan2_comments && !(def.noComments || []).includes(t.id)) P.panel.append(plan2CommentsCard(P.v.plan2_comments, P.v));
  };
  /* async work inside a tab (outages, maps...) - ignored if the tab was re-drawn meanwhile */
  P.tabCall = async (fn, kwargs, needs) => { const my = P.tseq; const r = await PY.call(fn, kwargs, {needs: needs || def.needs}); if (my !== P.tseq || !P.alive()) throw {stale: true}; return r.result; };
  P.start();
}

function landing(P){
  const d = DASHES.find(x => x.key === P.key), def = P.def;
  const el = h(`<div class="dland"><div><h2></h2><p>Drop the file(s) here – they stay in your browser and are shared by all three dashboards.</p><p class="lg"></p></div><div class="zones"></div></div>`);
  $("h2", el).textContent = "Load the data for " + d.title;
  const lg = $(".lg", el); lg.append(document.createTextNode("Optional files go in the strip above. What each file is for: ")); lg.append(guideBtn("dash/" + P.key));
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
    r1.append(lbl("Date field"), segc(spec.dateFields, f[spec.dateKey], v => { f[spec.dateKey] = v; el._showRange && el._showRange(); changed(); }));
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
    const rl = lbl("Range"); r1.append(rl, sel, custom);
    const showRange = () => { const off = f[spec.dateKey] === "Unplanned"; [rl, sel].forEach(x => x.style.display = off ? "none" : ""); custom.classList.toggle("hidden", off || sel.value !== "Custom range"); };
    el._showRange = showRange; showRange();
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
    spec.multis.forEach(m => { const o = v.options && v.options[m.key]; if (o) ms[m.key].setOptions(o); else if (m.cascade) ms[m.key].setOptions(v[m.cascade] || []); ms[m.key].set(f[m.key] || []); });
    cap.innerHTML = spec.caption ? spec.caption(v) : "";
    if (custom){ if (!custom.from.value && v.date_from) custom.from.value = v.date_from; if (!custom.to.value && v.date_to) custom.to.value = v.date_to; }
    drawChips();
  };
  drawChips();
  return el;
}
const capHTML = parts => parts.filter(Boolean).map(p => Array.isArray(p) ? `<b>${esc(p[0])}</b> ${esc(p[1] || "")}` : esc(p)).join(" · ");

/* ---------- most repeated comments in the Master for the Plan2 / Plan1 dates in view ---------- */
function plan2CommentsCard(c, v){ return planCommentsCard(c, v, "Plan2"); }
function planCommentsCard(c, v, field){
  const wrap = h(`<div class="dgrid" style="margin-top:16px"></div>`);
  const card = dcard(`${field} comments — most repeated`, "", 12); wrap.append(card);
  const blanks = field === "Blanks", unpl = field === "Unplanned", dword = blanks ? "Plan1" : field;
  const who = blanks ? "a passed Plan1 date and no Done date" : unpl ? "no Plan1 and no Done date" : `a ${field}${/date$/i.test(field) ? "" : " date"}`;
  const range = v.date_from ? `${fmtDate(v.date_from)} → ${fmtDate(v.date_to)}` : `all ${field} dates`;
  if (!c.available){ card.body.append(emptyMsg("This Master file has no comment column.")); return wrap; }
  $("p", card) || $(".tt", card).append(h(`<p></p>`));
  $(".tt p", card).textContent = `Comments on the rows with ${who}${unpl ? "" : ` in ${range}`} (same filters as above) · ${intf(c.with_comment)} of ${intf(c.rows)} rows have a comment · ${intf(c.distinct)} different comments`;
  if (!c.top.length){ card.body.append(emptyMsg(`No comments on the rows with ${who}${unpl ? "" : " in this range"}.`)); return wrap; }
  const [tc, tn, tj] = c.top[0];
  const call = h(`<div class="ctop"><div class="cq">“<span></span>”</div><div class="cm"><b></b><span></span></div></div>`);
  $(".cq span", call).textContent = tc; $(".cm b", call).textContent = `${intf(tn)} times`;
  $(".cm span", call).textContent = `on ${intf(tj)} job(s) · ${c.with_comment ? Math.round(tn / c.with_comment * 100) : 0}% of all comments in view · most repeated`;
  card.body.append(call);
  const top = c.top.slice(0, 15);
  let mode = "chart"; const holder = h(`<div></div>`);
  card.acts.append(segc([["chart", "Top 15"], ["table", `All ${intf(c.top.length)}`]], mode, m => { mode = m; draw(); }));
  card.acts.append(pngButton(card, {}));
  card.body.append(holder);
  function draw(){
    holder.textContent = "";
    if (mode === "table"){ holder.append(dtable(unpl ? {name: `${field} comments`, columns: ["Comment", "Times", "Jobs"], rows: c.top.map(r => r.slice(0, 3)), height: 460}
                                   : {name: `${field} comments`, columns: ["Comment", "Times", "Jobs", `First ${dword}`, `Last ${dword}`], rows: c.top, height: 460})); return; }
    const host = h(`<div class="viz"></div>`); holder.append(host);
    Viz.bars(host, {labels: top.map(x => x[0]), labelMax: 420, rowH: 30, fmt: intf, valueLabel: i => `${intf(top[i][1])}×`,
      series: [{name: "Times", values: top.map(x => x[1]), color: "var(--s1)"}],
      tip: i => ({title: top[i][0], rows: [{color: "var(--s1)", name: "Times written", value: intf(top[i][1])}, {name: "Jobs", value: intf(top[i][2])},
        ...(unpl ? [] : [{name: `${dword} dates`, value: `${fmtDate(top[i][3])} → ${fmtDate(top[i][4])}`}])]})});
  }
  requestAnimationFrame(draw);
  return wrap;
}

/* ---------- shared renderers ---------- */
function tableCard(title, sub, table, span = 12){ const c = dcard(title, sub, span); c.body.append(table.rows && table.rows.length ? dtable(table) : emptyMsg("No rows for the current filters.")); return c; }
const tbl = (t, name, extra) => ({columns: t.columns, rows: t.rows, name, ...(extra || {})});

/* ======================================================================
   Financial Dashboard (was Master Control)
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
    f.preset = f.preset || "All time"; f.granularity = f.granularity || "Auto"; f.pid_mode = f.pid_mode || "All";
    P.st.teams = P.st.teams || mcTeamsLoad(); },
  viewFn: "mc_view",
  lazyTabs: true,
  noComments: ["teams"],
  viewArgs(P){ const f = P.st.f; return {path: FILES.master.path, date_field: f.date_field, preset: f.preset, custom_from: f.custom_from || null, custom_to: f.custom_to || null,
    districts: f.districts || [], projects: f.projects || [], pids: f.pids || [], pms: f.pms || [], sourcefiles: f.sourcefiles || [],
    circuits: f.circuits || [], poles: f.poles || [], granularity: f.granularity, pid_mode: f.pid_mode, tab: P.st.tab, ...mcTeamArgs(P)}; },
  cascade: [["districts"], ["projects"], ["pids"], ["pms"], ["sourcefiles"], ["circuits", "circuit_options"], ["poles", "pole_options"]],
  filters(P){ const L = P.st.L, o = L.options;
    return {dateFields: L.date_fields, dateKey: "date_field", presets: L.presets, gran: ["Auto", "Day", "Week", "Month", "Year"],
      multis: [{key: "districts", label: "District", options: o.district}, {key: "projects", label: "Project", options: o.project},
               {key: "pids", label: "PID", options: o.pid}, {key: "pms", label: "Project Manager", options: o.pm},
               {key: "sourcefiles", label: "Source file", options: o.sourcefile},
               {key: "circuits", label: "Circuit", cascade: "circuit_options", hint: "Narrowed to the circuits the filters to its left allow."},
               {key: "poles", label: "Pole (enid)", cascade: "pole_options", hint: "Narrowed to the poles the filters to its left allow (including Circuit)."}],
      caption: v => capHTML([[fmtDate(v.date_from) + " → " + fmtDate(v.date_to)], [intf(v.records), "records"], `by ${v.granularity}${v.granularity_auto ? " (auto)" : ""}`])}; },
  tabs: [
    {id: "trends", label: "Trends", icon: "trend", render: mcTrends},
    {id: "pid", label: "PID Breakdown", icon: "table", render: mcPid},
    {id: "finance", label: "Finance", icon: "pound", render: mcFinance},
    {id: "teams", label: "Teams", icon: "team", render: mcTeams},
  ],
};
/* ---- Teams tab: which teams build the Teams chart (Trends) + optional leader names.
   Kept in this browser (localStorage) so it is there next time. */
const MC_TEAMS_KEY = "gaeltec.financial.teams";
function mcTeamsLoad(){
  try { const x = JSON.parse(localStorage.getItem(MC_TEAMS_KEY) || "null"); if (x && typeof x === "object") return {off: x.off || [], leaders: x.leaders || {}}; } catch(e){}
  return {off: [], leaders: {}};
}
function mcTeamsSave(t){ try { localStorage.setItem(MC_TEAMS_KEY, JSON.stringify(t)); } catch(e){} }
function mcTeamArgs(P){
  const t = P.st.teams || {off: [], leaders: {}}, all = ((P.st.L && P.st.L.teams) || []).map(x => x.team);
  const on = all.filter(x => !t.off.includes(x));
  // leader shown next to each team: the one typed in the Teams tab, otherwise the one found in the control file
  const leaders = {}; ((P.st.L && P.st.L.teams) || []).forEach(x => { const typed = (t.leaders || {})[x.team];
    const v = typed && typed.trim() ? typed.trim() : x.leader; if (v) leaders[x.team] = v; });
  return {teams: on.length === all.length ? null : (on.length ? on : ["(none)"]), team_leaders: leaders};
}
function mcTeams(P, panel){
  const L = P.st.L, t = P.st.teams, list = L.teams || [];
  if (!list.length){ panel.append(note("No team column values found in the control file.")); return; }
  const c = dcard("Teams", "Tick the teams to show in the Teams chart (Trends). The leader comes from the control files (the name on that team's own rows) – type a name to change it. Jobs done by two or more teams (AB, DE …) are split equally between them. Saved in this browser.");
  const all = h(`<button class="btn sm" type="button">Tick all</button>`), none = h(`<button class="btn sm" type="button">Untick all</button>`), fill = h(`<button class="btn sm" type="button">Reset to file leaders</button>`);
  c.acts.append(all, none, fill);
  const tb = h(`<div class="dt teamtbl"><div class="tw"><table><thead><tr><th style="width:44px"></th><th>Team</th><th style="text-align:right">Rows</th><th style="text-align:right">Shared rows</th><th>Team leader</th></tr></thead><tbody></tbody></table></div></div>`);
  const save = () => mcTeamsSave(t);
  const rows = list.map(r => {
    const tr = h(`<tr><td><input type="checkbox"></td><td class="tn"></td><td style="text-align:right" class="rn"></td><td style="text-align:right" class="sn"></td><td><input type="text" class="tl" style="width:100%;max-width:280px"></td></tr>`);
    const cb = $("input[type=checkbox]", tr), inp = $("input.tl", tr);
    $(".tn", tr).textContent = r.team; $(".rn", tr).textContent = intf(r.rows); $(".sn", tr).textContent = r.shared ? intf(r.shared) : "–";
    cb.checked = !t.off.includes(r.team);
    inp.value = t.leaders[r.team] || ""; inp.placeholder = r.leader ? `${r.leader} (from file)` : "leader name";
    cb.onchange = () => { t.off = t.off.filter(x => x !== r.team); if (!cb.checked) t.off.push(r.team); save(); };
    inp.oninput = () => { t.leaders[r.team] = inp.value; save(); };
    $("tbody", tb).append(tr);
    return {r, cb, inp};
  });
  all.onclick = () => { t.off = []; rows.forEach(x => x.cb.checked = true); save(); };
  none.onclick = () => { t.off = list.map(x => x.team); rows.forEach(x => x.cb.checked = false); save(); };
  fill.onclick = () => { rows.forEach(x => { x.inp.value = ""; }); t.leaders = {}; save(); };
  c.body.append(tb);
  panel.append(c);
}
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
  g.append(vizCard({title: "Revenue",
    legend: [{name: "Materials", color: "var(--s4)"}, {name: "Construction base (original / the rest)", color: "var(--s3)"}, {name: "Positive variation", color: "var(--s1)"}, {name: "Negative variation", color: "var(--s8)"}],
    draw: host => Viz.columns(host, {labels: p1.labels, height: 340, axis: v => moneyS(v),
      series: [{name: "Materials", values: p1.material, color: "var(--s4)"}, {name: "Base", values: p1.base, color: "var(--s3)"}, {name: "Variance", values: p1.cap, colorAt: varCol}],
      tip: i => ({title: p1.labels[i], rows: [{color: "var(--s3)", name: "Construction total", value: money(p1.total[i])}, {name: "Original", value: money(p1.orig[i])},
        {color: varCol(i), name: p1.variance[i] >= 0 ? "Positive variation" : "Negative variation", value: signed(p1.variance[i])}, {color: "var(--s4)", name: "Materials", value: money(p1.material[i])}],
        total: {name: "Total", value: money(p1.total[i] + p1.material[i])}})}),
    table: () => ({name: "Revenue", columns: ["Period", "Construction total (£)", "Original (£)", "Variance (£)", "Materials (£)", "Total (£)"],
      rows: p1.labels.map((l, i) => [l, p1.total[i], p1.orig[i], p1.variance[i], p1.material[i], p1.total[i] + p1.material[i]])})}));
  const p2 = v.trends.panel2;
  if (!p2 || !p2.leaders.length){ g.append(note(p2 ? "No teams ticked – choose them in the Teams tab." : "No construction records under these filters.")); return; }
  const people = p2.leaders.filter(l => l !== "Other"), col = P.ck.assign(people);
  const cOf = l => l === "Other" ? "var(--neutral)" : col(l);
  const c2 = vizCard({title: "Teams",
    legend: p2.leaders.map(l => ({name: l, color: cOf(l)})),
    draw: host => Viz.columns(host, {labels: p2.labels, height: 320, axis: v => moneyS(v), fmt: money,
      series: p2.leaders.map(l => ({name: l, values: p2.values[l], color: cOf(l)})),
      tip: i => { const rows = p2.leaders.map(l => ({color: cOf(l), name: l, value: money(p2.values[l][i] || 0)})).filter((r, j) => (p2.values[p2.leaders[j]][i] || 0) !== 0);
        return {title: p2.labels[i], rows, total: {name: "Total", value: money(p2.leaders.reduce((a, l) => a + (p2.values[l][i] || 0), 0))}}; }}),
    table: () => ({name: "Teams", columns: ["Period", ...p2.leaders.map(l => l + " (£)")], rows: p2.labels.map((lab, i) => [lab, ...p2.leaders.map(l => p2.values[l][i] || 0)])})});
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
  const c = dcard("PID Breakdown");
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
   Planning Dashboard (was Network Job Tracker)
   ====================================================================== */
DASHDEF.tracker = {
  slots: [["master", true], ["outages", false], ["forecast", false], ["ics", false], ["misc", false]], loadSlots: ["master"], viewSlots: ["misc"], needs: NEED_NT,
  lazyTabs: true,
  about: ["Master file (.parquet or .csv) – required. The others are optional and can be dropped any time in the strip at the top:",
          "High level planning (High-level_planning_2026.xlsx) → the Jobs tab calendar with the poles & tasks for each day.",
          "Service Partner Workbank → Pole Position.  Outage calendar (.ics) → the events list.",
          "Service codes (miscelaneous.parquet) → the service codes in Totals – remembered in this browser once dropped."],
  loadFn: "nt_load", loadArgs: P => ({path: FILES.master.path, cols: P.st.cols}),
  init(P){ const st = P.st, L = st.L, f = st.f; st.cols = L.cols;
    if (L.date_fields && !L.date_fields.includes(f.date_field)) f.date_field = L.default_date_field;
    f.preset = f.preset || "All time"; st.gran = st.gran || "Month"; st.o = st.o || {view: "calendar"}; st.fc = st.fc || {};
    f.pid_mode = f.pid_mode || "All"; restoreMisc(); },
  pre(P){   /* the columns are found automatically; only say so when one can't be found */
    const L = P.st.L;
    if (L.missing_required.length) return note(`This Master file has no column for: ${L.missing_required.map(k => (L.col_fields.find(x => x.key === k) || {}).label || k).join(", ")}. Check it's the Master_DD-MM-YYYY file from Build Master.`, "warn", "warn");
    if (L.warn_no_pole) return note("This Master file has no pole (enid) column – pole de-duplication and the pole filter are off.", "warn", "warn");
    return null;
  },
  ready: P => !P.st.L.missing_required.length,
  noComments: ["forecast"],
  viewFn: "nt_view",
  ntFilters(P){ const f = P.st.f; return {date_field: f.date_field, preset: f.preset, custom_from: f.custom_from || null, custom_to: f.custom_to || null,
    districts: f.districts || [], projects: f.projects || [], pids: f.pids || [], jobs: f.jobs || [], sourcefiles: f.sourcefiles || [], circuits: f.circuits || [], poles: f.poles || []}; },
  viewArgs(P){ return {path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P), granularity: P.st.gran, detail: P.st.detail || null,
    misc: FILES.misc ? FILES.misc.path : "", pid_mode: P.st.f.pid_mode || "All", tab: P.st.tab}; },
  cascade: [["districts"], ["projects"], ["pids"], ["jobs"], ["sourcefiles"], ["circuits", "circuit_options"], ["poles", "pole_options"]],
  filters(P){ const L = P.st.L, o = L.options;
    const hint = "Narrowed to what the filters to its left allow.";
    return {dateFields: L.date_fields, dateKey: "date_field", presets: L.presets,
      multis: [{key: "districts", label: "District", options: o.district}, {key: "projects", label: "Project", options: o.project, hint},
               {key: "pids", label: "PID", options: o.pid, hint}, {key: "jobs", label: "Job name", options: o.job, hint},
               {key: "sourcefiles", label: "Source File", options: o.sourcefile || [], hint},
               {key: "circuits", label: "Circuit", cascade: "circuit_options", hint},
               {key: "poles", label: "Pole", cascade: "pole_options", hint: "Narrowed to the poles of the circuits (and other filters) to its left."}],
      caption: v => v.unplanned ? capHTML([["Unplanned", "– no Plan1 or Done date"], [intf(v.rows), `of ${intf(v.rows_total)} rows`]])
                                : capHTML([P.st.f.date_field === "Blanks" ? ["Blanks", "– Plan1 passed, no Done"] : "", v.date_from ? [fmtDate(v.date_from) + " → " + fmtDate(v.date_to)] : "", [intf(v.rows), `of ${intf(v.rows_total)} rows`]])}; },
  tabs: [
    {id: "overview", label: "Overview", icon: "trend", render: ntOverview},
    {id: "jobs", label: "Jobs", icon: "cal", render: ntJobs},
    {id: "items", label: "Mapped Items", icon: "grid", render: ntItems},
    {id: "forecast", label: "Pole Position", icon: "chart", render: ntForecastAndPid},
    {id: "totals", label: "Totals", icon: "pound", render: ntTotals},
  ],
};
const periodLabel = (iso, gran) => { const [y, m, d] = iso.split("-"); return gran === "Month" ? `${MONTHS3[+m - 1]} ${y}` : `${+d} ${MONTHS3[+m - 1]}${gran === "Week" ? "" : ""} ${gran === "Day" || gran === "Week" ? y.slice(2) : ""}`.trim(); };
function ntOverview(P, panel){
  const o = P.v.overview;
  panel.append(kpis([{label: "Total CV7_recover count", value: intf(o.recover_total), hero: true}, {label: "Total poles (all categories)", value: intf(o.pole_total)},
    {label: "Rows after filters", value: intf(P.v.rows), note: `of ${intf(P.v.rows_total)}`}]));
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  if (P.v.unplanned_poles){
    const u = P.v.unplanned_poles;
    if (!u.labels.length){ const c = dcard(`Unplanned poles by ${u.by.toLowerCase()}`, "", 12); c.body.append(emptyMsg("No unplanned poles under the current filters.")); g.append(c); }
    else {
      const cols = ["var(--s1)", "var(--s2)", "var(--s3)", "var(--s4)", "var(--s5)", "var(--s6)"];
      const series = u.series.map((x, i) => ({name: x.name, values: x.values, color: cols[i % cols.length]}));
      g.append(vizCard({title: `Unplanned poles by ${u.by.toLowerCase()}`, span: 12,
        sub: (u.by === "Circuit" ? "Rows with no Plan1 and no Done date, for the selected project(s) · by circuit" : "Rows with no Plan1 and no Done date · pick a Project to see it by circuit") + ` · ${intf(u.labels.length)} ${u.by.toLowerCase()}(s)`,
        legend: series.length > 1 ? series.map(x => ({name: x.name, color: x.color})) : null,
        draw: host => Viz.columns(host, {labels: u.labels, height: 360, axis: compact, fmt: intf, valueLabels: u.labels.length <= 40, series,
          tip: i => ({title: `${u.by}: ${u.labels[i]}`, rows: series.filter(x => x.values[i]).map(x => ({color: x.color, name: x.name, value: intf(x.values[i])})), total: {name: "Unplanned poles", value: intf(u.totals[i])}, ...barList(u.detail, i)})}),
        table: () => ({name: `Unplanned poles by ${u.by}`, columns: [u.by, ...series.map(x => x.name), "Total"], rows: u.labels.map((l, i) => [l, ...series.map(x => x.values[i]), u.totals[i]])})}));
    }
  }
  const gsel = segc(["Day", "Week", "Month"], P.st.gran, x => { P.st.gran = x; P.refresh(); });
  if (P.v.unplanned_poles){ /* no dates to plot over time */ }
  else if (!o.trend){ const c = dcard("CV7_recover — count over time", "", 12); c.acts.append(gsel); c.body.append(emptyMsg("No CV7_recover records (with a date) for the current filters.")); g.append(c); }
  else { const labs = o.trend.x.map(x => periodLabel(x, P.st.gran)), ys = o.trend.y, mx = Math.max(...ys);
    const cvOpts = {labels: labs, height: 340, axis: compact, fmt: intf,
        series: [{name: "CV7_recover", values: ys, color: "var(--s1)"}]};
    g.append(vizCard({title: "CV7_recover — count over time", sub: `Grouped by ${P.st.gran.toLowerCase()} · hover a bar for its count`, span: 12, extra: gsel,
      exportDraw: host => Viz.columns(host, {...cvOpts, height: 420, valueLabels: "all"}),          // picture: value on every bar
      exportWidth: Math.max(1400, labs.length * 26 + 120),
      draw: host => Viz.columns(host, {labels: labs, height: 340, axis: compact, fmt: intf, valueLabels: ys.length <= 40,
        series: [{name: "CV7_recover", values: ys, color: "var(--s1)"}],
        tip: i => ({title: labs[i], rows: [{color: "var(--s1)", name: "CV7_recover", value: intf(ys[i])}], note: ys[i] === mx ? "Highest period" : "",
          ...barList(o.trend.detail, i)})}),
      table: () => ({name: "CV7_recover over time", columns: ["Date", "Count"], rows: o.trend.x.map((x, i) => [x, o.trend.y[i]])})})); }
  if (!o.poles.length){ const c = dcard("All pole categories (erect only)", "", 12); c.body.append(emptyMsg("No pole records for the current filters.")); g.append(c); }
  else g.append(vizCard({title: "All pole categories (erect only)", sub: "Count by pole type", span: 12,
    draw: host => Viz.bars(host, {labels: o.poles.map(p => p.type), series: [{name: "Count", values: o.poles.map(p => p.count), color: "var(--s1)"}], fmt: intf, labelMax: 200}),
    table: () => ({name: "Pole categories", columns: ["Pole type", "Count"], rows: o.poles.map(p => [p.type, p.count])})}));
  if (P.v.sf_pivot) panel.append(sfPivotCard(P.v.sf_pivot));
  if (P.v.plan1_comments) panel.append(planCommentsCard(P.v.plan1_comments, P.v, P.st.f.date_field || "Combined date"));
}
/* hover list for a bar: the circuits and jobs behind it (biggest first) */
function barList(detail, i){
  const d = detail && detail[i]; if (!d || !d[0].length) return {};
  return {list: d[0].map(x => [x[0], x[1], intf(x[2])]), more: d[1]};
}
const dateSpan = (a, b, n) => !a ? "" : (a === b ? fmtDate(a) : `${fmtDate(a)} → ${fmtDate(b)}${n > 2 ? ` · ${intf(n)} dates` : ""}`);
/* ---------- Overview: source file → circuit → most repeated job, opened like a pivot table ---------- */
function sfPivotCard(pv){
  const wrap = h(`<div class="dgrid" style="margin-top:16px"></div>`);
  const card = dcard("Source files → circuits → most repeated job", "", 12); wrap.append(card);
  if (!pv.available){ card.body.append(emptyMsg("This Master file has no source file or circuit column.")); return wrap; }
  const files = pv.files || [];
  if (!files.length){ card.body.append(emptyMsg("No rows for the current filters.")); return wrap; }
  $(".tt", card).append(h(`<p>${intf(files.length)} source file(s) · click a source file to open its circuits and the job repeated most on each · dates = that job's dates on the circuit, inside the date range above</p>`));
  const el = h(`<div class="dt sfpv"><div class="dth"><input type="text" placeholder="Search source file, circuit or job…"><span class="cnt"></span>
    <button class="btn sm ex">Open all</button><button class="btn sm co">Close all</button><button class="btn sm dl" title="Download as CSV">${ICON2.csv} CSV</button></div>
    <div class="tw" style="max-height:560px"><table><thead><tr><th style="width:26%">Source file / Circuit</th><th class="pn">Rows</th><th class="pn">Circuits</th><th class="pn">Jobs</th><th>Most repeated job</th><th class="pn">Times</th><th>Dates</th></tr></thead><tbody></tbody></table></div><div class="foot"></div></div>`);
  card.body.append(el);
  const tb = $("tbody", el), inp = $("input", el), foot = $(".foot", el), cnt = $(".cnt", el);
  const open = new Set(); let limit = 100, shown = files;
  const td = (t, cls) => { const c = document.createElement("td"); if (cls) c.className = cls; c.textContent = t; return c; };
  function draw(){
    const q = inp.value.trim().toLowerCase();
    shown = !q ? files : files.filter(f => f[0].toLowerCase().includes(q) || f[9].some(c => c[0].toLowerCase().includes(q) || String(c[3]).toLowerCase().includes(q)));
    cnt.textContent = `${intf(shown.length)} of ${intf(files.length)}`;
    tb.textContent = "";
    shown.slice(0, limit).forEach(f => {
      const isOpen = open.has(f[0]) || (q && !f[0].toLowerCase().includes(q));
      const tr = document.createElement("tr"); tr.className = "sfp" + (isOpen ? " open" : "");
      const c0 = td(""); c0.innerHTML = `<span class="car">▸</span>`; c0.append(document.createTextNode(f[0]));
      tr.append(c0, td(intf(f[1]), "pn"), td(intf(f[2]), "pn"), td(intf(f[3]), "pn"), td(f[4]), td(intf(f[5]), "pn"), td(dateSpan(f[6], f[7], f[8]), "dts"));
      tr.onclick = () => { open.has(f[0]) ? open.delete(f[0]) : open.add(f[0]); draw(); };
      tb.append(tr);
      if (isOpen) f[9].forEach(c => {
        if (q && !f[0].toLowerCase().includes(q) && !(c[0].toLowerCase().includes(q) || String(c[3]).toLowerCase().includes(q))) return;
        const r = document.createElement("tr"); r.className = "sfc";
        r.append(td(c[0], "ind"), td(intf(c[1]), "pn"), td("", "pn"), td(intf(c[2]), "pn"), td(c[3]), td(intf(c[4]), "pn"), td(dateSpan(c[5], c[6], c[7]), "dts"));
        tb.append(r);
      });
    });
    foot.textContent = "";
    if (shown.length > limit){ const b = h(`<button class="btn sm">Show ${intf(Math.min(100, shown.length - limit))} more</button>`); b.onclick = () => { limit += 100; draw(); }; foot.append(b); }
  }
  inp.oninput = () => { limit = 100; draw(); };
  $(".ex", el).onclick = () => { shown.slice(0, limit).forEach(f => open.add(f[0])); draw(); };
  $(".co", el).onclick = () => { open.clear(); draw(); };
  $(".dl", el).onclick = () => downloadCSV("Source files - circuits - most repeated job",
    ["Source file", "Circuit", "Rows", "Jobs", "Most repeated job", "Times", "First date", "Last date", "Number of dates"],
    files.flatMap(f => f[9].map(c => [f[0], c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7]])));
  draw();
  return wrap;
}
const GROUP_ICON = {"Poles": "pole", "Transformers": "tx", "Conductor": "cable", "Switch gear": "switch"};
const GROUP_IMAGE = {"Poles": "Poles", "Transformers": "Transformer", "Conductor": "Cable", "Switch gear": "Switchgear"};
const ITEM_IMAGE = {"Fuses": "Fuses"};
function ntItems(P, panel){
  const it = P.v.items;
  if (!it.detail_names.length){ panel.append(note("No mapped items for the current filters.")); return; }
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const chosen = it.detail && it.detail.chosen;
  const cardEl = c => { const b = h(`<button class="icard${c.name === chosen ? " on" : ""}"><small></small><b></b></button>`); $("small", b).textContent = c.label; $("b", b).textContent = c.value;
    if (ITEM_IMAGE[c.name]){ b.classList.add("withimg"); b.prepend(imageEl(ITEM_IMAGE[c.name], "", "iimg")); }
    b.title = "Show the rows behind " + c.label; b.onclick = () => { P.st.detail = c.name; P.refresh(); }; return b; };
  it.groups.forEach(gr => {
    const c = dcard(gr.title, "", 6); c.querySelector(".ch").remove();
    const w = h(`<div class="igrp centre"><div class="ih"><h3></h3><p></p></div><div class="icards"></div></div>`);
    $(".ih", w).after(imageEl(GROUP_IMAGE[gr.title], `<div class="ico">${ICON2[GROUP_ICON[gr.title]] || I.box3}</div>`, "gimg"));
    $("h3", w).textContent = gr.title; $("p", w).textContent = gr.cards.length ? `${gr.cards.length} item type(s)` : "No records for this group under the current filters.";
    gr.cards.forEach(x => $(".icards", w).append(cardEl(x)));
    c.body.append(w); g.append(c);
  });
  if (it.other.length){ const c = dcard("Other items", "", 12); const w = h(`<div class="icards centre"></div>`); it.other.forEach(x => w.append(cardEl(x))); c.body.append(w); g.append(c); }
  if (it.detail){
    const c = dcard("Details", "The rows behind the selected item – click any card above, or pick here", 12);
    const sel = h(`<select></select>`); it.detail_names.forEach(n => { const o = document.createElement("option"); o.textContent = n; sel.append(o); }); sel.value = chosen;
    sel.onchange = () => { P.st.detail = sel.value; P.refresh(); };
    c.acts.append(sel); c.body.append(dtable(tbl(it.detail, "Details - " + chosen, {height: 380}))); g.append(c);
  }
}
/* the miscelaneous file dropped earlier (here or in the Outputs tool) is put back in its slot */
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
  if (t.total == null){ panel.append(note("The Master file has no 'total' column, so there are no totals to show.", "warn", "warn")); return; }
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
  const allRows = vr.detail_rows || vr.detail.rows.length, shown = vr.detail.rows.length;
  const ec = tableCard("Variations — every item", (allRows > shown ? `Showing the first ${intf(shown)} of ${intf(allRows)} rows` : `${intf(allRows)} row(s)`) +
    " · the same rows and columns as under the pivot on the Excel Variations sheet", tbl(vr.detail, "Variations - every item", {height: 460}));
  if (allRows > shown){
    const b = h(`<button class="btn sm">${ICON2.csv}<span></span></button>`); $("span", b).textContent = `Download all ${intf(allRows)} rows (CSV)`;
    b.onclick = async () => { const lab = $("span", b), was = lab.textContent; b.disabled = true; lab.textContent = "Preparing the CSV…";
      try { const r = await PY.call("nt_variations_csv", {path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P), misc: FILES.misc ? FILES.misc.path : ""}, {needs: NEED_NT});
            (r.files || []).forEach(f => downloadBlob(f.name, f.data)); }
      catch(e){ alert(e.message); }
      b.disabled = false; lab.textContent = was; };
    ec.acts.append(b);
  }
  g.append(ec);
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
/* ---- Jobs: High level planning calendar + .ics ---- */
function ntJobs(P, panel){
  const g = h(`<div class="dgrid"></div>`); panel.append(g);
  const oc = dcard("High level planning", "Every year sheet · colour = district (darker = At Risk, red = Cancelled) · hover an outage for its figures · click a day or an outage for the full details (same as the Outlook calendar)", 12);
  oc.classList.add("calwide");
  g.append(oc);
  if (!FILES.outages){ oc.body.append(slotZone("outages", false, "Drop the High level planning workbook – the calendar shows here")); }
  else ntOutages(P, oc);
  const ic = dcard("Outage calendar (.ics)", "Events from the outage calendar file", 12); g.append(ic);
  if (!FILES.ics){ ic.body.append(slotZone("ics", false, "Drop outage_calendar_2026 (vN).ics")); }
  else { ic.body.append(spinner("Reading the calendar…"));
    P.tabCall("nt_ics", {ics_path: FILES.ics.path}).then(r => { ic.body.textContent = ""; $("p", ic).textContent = `📅 Using: ${r.name} · ${intf(r.count)} event(s)`;
      ic.body.append(r.count ? dtable({columns: r.columns, rows: r.rows, name: "Outage calendar"}) : emptyMsg("No VEVENT entries found in this calendar file.")); })
      .catch(e => { if (!e.stale){ ic.body.textContent = ""; ic.body.append(errNote(e)); } }); }
}
const calMonthOf = (o, r) => {
  if (o.month) return o.month;
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`, t = iso(new Date());
  if (r && r.min && r.max && t >= r.min.slice(0, 7) && t <= r.max.slice(0, 7)) return t;
  return r && r.min ? r.min.slice(0, 7) : t;
};
function ntOutages(P, card){
  const o = P.st.o, body = card.body;
  body.append(spinner("Reading the High level planning…"));
  const month = calMonthOf(o, o._range);
  P.tabCall("nt_outages", {outage_path: FILES.outages.path, master_path: FILES.master.path, cols: P.st.cols, filters: DASHDEF.tracker.ntFilters(P),
    o_districts: o.districts || [], o_pms: o.pms || [], o_from: o.from || null, o_to: o.to || null, selected_date: o.date || null,
    password: CALX.password || "", roster: rosterText(), absence: CALX.absence, month}, NEED_NT)
  .then(r => {
    body.textContent = "";
    const redo = () => { body.textContent = ""; ntOutages(P, card); };
    if (r.need_password){
      const pw = h(`<div class="row" style="align-items:flex-end;gap:10px"><div class="field"><label>🔒 Password of the High level planning</label><input type="password" autocomplete="off" style="width:220px"></div><button class="btn">Open</button></div>`);
      const go = () => { CALX.password = $("input", pw).value; redo(); };
      $("button", pw).onclick = go; $("input", pw).addEventListener("keydown", e => { if (e.key === "Enter") go(); });
      body.append(pw); if (CALX.password) body.append(note("That password didn't open it – try again.")); return;
    }
    o._range = {min: r.min, max: r.max};
    if (!o.month && month !== calMonthOf(o, o._range)){ o.month = calMonthOf(o, o._range); redo(); return; }
    const bar = h(`<div class="fr" style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px"></div>`);
    bar.append(msel({label: "District", options: r.districts, value: o.districts || [], onChange: v => { o.districts = v; redo(); }}),
               msel({label: "SPEN PM", options: r.pms, value: o.pms || [], onChange: v => { o.pms = v; redo(); }}));
    const di = (k, def) => { const i = h(`<input type="date">`); i.value = o[k] || def || ""; i.onchange = () => { o[k] = i.value || null; redo(); }; return i; };
    bar.append(h(`<span class="lbl" style="font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;margin-left:6px">Outage dates</span>`), di("from", r.min), h(`<span class="muted">→</span>`), di("to", r.max));
    bar.append(segc([["calendar", "Calendar"], ["table", "Table"]], o.view || "calendar", v => { o.view = v; redo(); }));
    const cnt = h(`<span class="muted" style="margin-left:auto;font-size:13px"></span>`); cnt.textContent = `${intf(r.rows)} of ${intf(r.rows_total)} outages`; bar.append(cnt);
    body.append(bar);
    body.append(calExtrasBar(P, r, month, redo));
    if ((o.view || "calendar") === "table"){ body.append(dtable({columns: r.table.columns, rows: r.table.rows, name: "High level planning", height: 520})); return; }
    body.append(calLegend(r));
    body.append(calendarEl(r, o, month, iso => { o.date = iso; o.sel = null; redo(); }, (iso, id) => { o.date = iso; o.sel = id; redo(); }, m => { o.month = m; redo(); }));
    if (r.day) body.append(dayPanel(r.day, o.sel));
    else body.append(h(`<p class="muted" style="margin:12px 0 0">Click a date or an outage above to see the day summary, the teams and each outage's details here.</p>`));
  })
  .catch(e => { if (!e.stale){ body.textContent = ""; body.append(errNote(e)); } });
}
/* roster / absence / folder look-ups – compact bar above the calendar */
function calExtrasBar(P, r, month, redo){
  const el = h(`<div class="calx"><span class="cxs"></span><button class="btn sm ts">👷 Team breakdown…</button><button class="btn sm fl">📁 Look inside this month's outage folders</button><span class="status"></span></div>`);
  const rs = rosterSummary(rosterText()), st = $(".status", el);
  const bits = [];
  bits.push(r.has_master ? "✔ Master figures" : "Master not loaded – no poles / switches");
  bits.push(r.has_team ? "✔ Teams available" : (rs.teams ? (CALX.absence.length ? "Teams: roster + absence loaded" : "Teams: drop the absence CSVs") : "Teams: load the roster"));
  $(".cxs", el).textContent = bits.join(" · ");
  const panel = h(`<div class="calxp hidden"></div>`);
  $(".ts", el).onclick = () => { if (!panel.firstChild) panel.append(teamBlock(() => redo())); panel.classList.toggle("hidden"); };
  $(".fl", el).onclick = async () => {
    st.className = "status";
    try {
      const f = (await PY.call("nt_outage_folders", {outage_path: FILES.outages.path, password: CALX.password || "", o_districts: P.st.o.districts || [], o_pms: P.st.o.pms || [], month}, {})).result;
      if (!f.all){ st.textContent = "No outage this month has a folder link."; return; }
      if (!f.folders.length){ st.className = "status ok"; st.textContent = `All ${f.all} folder(s) already looked up.`; return; }
      const res = await outageFolderLookups(f.folders, st);
      st.className = "status ok"; st.textContent = `✔ ${res.n} folder(s) looked up`; st.title = folderSummary(res);
      setTimeout(redo, 400);
    } catch(e){ if (e.name === "AbortError"){ st.textContent = ""; return; } st.className = "status err"; st.textContent = e.message; }
  };
  const w = h(`<div></div>`); w.append(el, panel); return w;
}
function calLegend(r){
  const el = h(`<div class="calleg"></div>`);
  r.legend.forEach(d => { const x = h(`<span><i></i><i></i><b></b></span>`); const [a, b] = x.querySelectorAll("i"); a.style.background = d.green; b.style.background = d.risk; $("b", x).textContent = d.district; el.append(x); });
  const c = h(`<span><i></i><b>Cancelled</b></span>`); $("i", c).style.background = r.cancelled; el.append(c);
  el.append(h(`<span class="muted">light = on track · dark = At Risk</span>`));
  return el;
}
const _n0 = v => v ? (Number.isInteger(v) ? String(v) : String(+v.toFixed(2))) : "0";
function evTip(e){
  const rows = [{name: "Recover poles", value: _n0(e.poles)}];
  const tx = Object.entries(e.tx || {}); rows.push({name: "Transformers", value: tx.length ? tx.map(([k, v]) => `${_n0(v)}× ${k}`).join(", ") : "0"});
  rows.push({name: "Noja", value: _n0(e.noja)}, {name: "ABSW", value: _n0(e.absw)}, {name: "Soule", value: _n0(e.soule)});
  rows.push({name: "Customers", value: e.customers || (e.folder ? (e.looked ? "–" : "look inside the folder") : "–")});
  rows.push({name: "Teams available", value: e.teams === null || e.teams === undefined ? "load roster + absence" : String(e.teams)});
  return {title: [e.outage, e.scheme].filter(Boolean).join(" · ") || "Outage",
    rows, note: [e.district, e.circuit && "Circuit " + e.circuit, e.project, e.job && "Job: " + e.job, e.pm && "PM " + e.pm, e.time, e.status].filter(Boolean).join(" · ")};
}
function calendarEl(r, o, month, onDay, onEv, onMonth){
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const today = iso(new Date());
  const [yy, mm] = month.split("-").map(Number);
  const first = new Date(yy, mm - 1, 1), start = new Date(first); start.setDate(1 - ((first.getDay() + 6) % 7));   // weeks start Monday
  const byDay = new Map(); r.events.forEach(e => { if (!byDay.has(e.start)) byDay.set(e.start, []); byDay.get(e.start).push(e); });
  const el = h(`<div class="cal big"><div class="calh"><button class="btn sm" title="Previous month">‹</button><h4></h4><button class="btn sm">Today</button><button class="btn sm" title="Next month">›</button></div><div class="calg"></div></div>`);
  const mNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const mEv = r.events.filter(e => e.start.slice(0, 7) === month).length;
  $("h4", el).textContent = `${mNames[mm - 1]} ${yy}  ·  ${mEv} outage(s)`;
  const shift = n => { const d = new Date(yy, mm - 1 + n, 1); onMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`); };
  const [bp, bt, bn] = el.querySelectorAll(".calh .btn");
  bp.onclick = () => shift(-1); bn.onclick = () => shift(1); bt.onclick = () => onMonth(today.slice(0, 7));
  const grid = $(".calg", el), days = r.days || {};
  ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"].forEach(d => { const x = h(`<div class="dow"></div>`); x.textContent = d; grid.append(x); });
  const weeks = Math.ceil((((first.getDay() + 6) % 7) + new Date(yy, mm, 0).getDate()) / 7);
  const MAXV = 5;
  for (let i = 0; i < weeks * 7; i++){
    const d = new Date(start); d.setDate(start.getDate() + i); const k = iso(d), evs = byDay.get(k) || [], ds = days[k];
    const inM = d.getMonth() === mm - 1;
    const cell = h(`<div class="day${inM ? "" : " oth"}${k === today ? " today" : ""}${k === o.date ? " sel" : ""}${ds && ds.flag ? " flag" : ""}"><div class="dn"><span class="dnum"></span><span class="badges"></span></div><div class="evs"></div></div>`);
    $(".dnum", cell).textContent = d.getDate();
    if (inM && ds){
      const bd = $(".badges", cell);
      if (ds.poles) bd.append(h(`<span class="bdg" title="Recover poles">⬆${_n0(ds.poles)}</span>`));
      if (ds.switches) bd.append(h(`<span class="bdg sw" title="Switches (Noja / Soule / ABSW)">⚡${_n0(ds.switches)}</span>`));
      if (ds.tx) bd.append(h(`<span class="bdg sw" title="Transformers">TX${_n0(ds.tx)}</span>`));
      if (ds.teams !== null && ds.teams !== undefined) bd.append(h(`<span class="bdg tm${ds.teams < ds.teams_total ? " low" : ""}" title="Complete teams available">👷${ds.teams}/${ds.teams_total}</span>`));
    }
    const box = $(".evs", cell);
    evs.slice(0, MAXV).forEach(e => {
      const ev = h(`<div class="ev2${e.id === o.sel ? " on" : ""}"><b></b><span class="l2"></span><span class="l3"></span></div>`);
      ev.style.background = e.fill; ev.style.color = e.ink;
      $("b", ev).textContent = [e.outage, e.scheme].filter(Boolean).join(" · ") || "Outage";
      $(".l2", ev).textContent = [e.circuit, e.project].filter(Boolean).join(" · ");
      $(".l3", ev).textContent = e.job || e.pm || "";
      if (!$(".l3", ev).textContent) $(".l3", ev).remove();
      ev.onmouseenter = x => Tip.show(x, evTip(e)); ev.onmousemove = x => Tip.move(x); ev.onmouseleave = () => Tip.hide();
      ev.onclick = x => { x.stopPropagation(); Tip.hide(); onEv(k, e.id); };
      box.append(ev);
    });
    if (evs.length > MAXV){ const m = h(`<div class="more"></div>`); m.textContent = `+${evs.length - MAXV} more – click the day`; box.append(m); }
    const dn = $(".dn", cell);
    if (inM && (ds || evs.length)){
      dn.onmouseenter = x => Tip.show(x, {title: fmtDate(k), rows: [
        {name: "Outages", value: String(evs.length)},
        ...(ds ? [{name: "Recover poles", value: _n0(ds.poles)}, {name: "Switches", value: _n0(ds.switches)}, {name: "Transformers", value: _n0(ds.tx)}] : []),
        ...(ds && ds.teams !== null && ds.teams !== undefined ? [{name: "Complete teams", value: `${ds.teams} of ${ds.teams_total}`}] : [])], note: "Click for the full day summary"});
      dn.onmousemove = x => Tip.move(x); dn.onmouseleave = () => Tip.hide();
    }
    cell.onclick = () => { Tip.hide(); onDay(k); };
    grid.append(cell);
  }
  return el;
}
function sectionsEl(sections){
  const w = h(`<div class="ocs"></div>`);
  sections.forEach(([title, items]) => {
    const sec = h(`<div class="ocsec"><div class="octh"></div><table></table></div>`);
    $(".octh", sec).textContent = title.toUpperCase();
    const tb = $("table", sec);
    items.forEach(([label, text, url]) => {
      const tr = document.createElement("tr"), val = document.createElement("td");
      val.className = "v";
      if (label){
        const l = document.createElement("td"); l.className = "l"; l.textContent = label;
        if (url && /^file:/i.test(url) && label === "Outage Folder") val.append(sectionsLink(filePathOf(url), url, true));
        else val.append(url ? sectionsLink(text, url) : text);
        tr.append(l, val);
      } else { val.colSpan = 2; val.append("• ", url ? sectionsLink(text, url) : text); tr.append(val); }
      tb.append(tr);
    });
    w.append(sec);
  });
  return w;
}
const filePathOf = url => "\\\\" + decodeURIComponent(url.replace(/^file:\/+/i, "")).replace(/\//g, "\\");
function sectionsLink(text, url, showPath){
  if (/^file:/i.test(url)){   /* a browser can't open network folders from a web page – copy the path instead */
    const path = filePathOf(url);
    const s = h(`<span class="fpath"></span>`); s.textContent = text;
    const cp = h(`<button class="btn sm ghost cpy" title="Copy the path, then paste it in File Explorer's address bar">Copy path</button>`);
    cp.onclick = ev => { ev.stopPropagation(); ev.preventDefault(); navigator.clipboard.writeText(path).then(() => { cp.textContent = "Copied ✔"; setTimeout(() => cp.textContent = "Copy path", 1500); }); };
    const w = h(`<span></span>`); w.title = path; w.append(s, " ", cp); return w;
  }
  const a = h(`<a target="_blank" rel="noopener"></a>`); a.href = url; a.textContent = text; return a;
}
function dayPanel(day, selId){
  const w = h(`<div class="dgrid calwide" style="margin-top:14px"></div>`);
  const a = dcard(`${fmtDate(day.date)} – day summary`, "Master figures (date field of the filter bar) and the team breakdown", 4);
  const lines = [...(day.summary || [])];
  if (day.team && day.team.length){ if (lines.length) lines.push(""); lines.push(...day.team); }
  if (!lines.length) a.body.append(emptyMsg("No Master figures or team data for this day."));
  else { const pre = h(`<div class="daysum"></div>`); lines.forEach(l => { const d = document.createElement("div"); d.textContent = l || "\u00a0";
    if (/^(Team liders off - [1-9]|.* off: |Supervisors off - [1-9])/.test(l)) d.className = "warn"; pre.append(d); }); a.body.append(pre); }
  const b = dcard(`Outages on ${fmtDate(day.date)}`, day.outages.length ? `${day.outages.length} outage(s) · the same details as the Outlook calendar` : "", 8);
  if (!day.outages.length) b.body.append(emptyMsg("No outages on this date under the current filters."));
  day.outages.forEach((ou, i) => {
    const open = selId ? ou.id === selId : i === 0;
    const det = h(`<details class="ocard"${open ? " open" : ""}><summary><i></i><span></span></summary></details>`);
    $("i", det).style.background = ou.fill; $("span", det).textContent = ou.title;
    det.append(sectionsEl(ou.sections));
    if (ou.id === selId) setTimeout(() => det.scrollIntoView({behavior: "smooth", block: "center"}), 80);
    b.body.append(det);
  });
  const c = dcard(`Poles & MD Poling tasks on ${fmtDate(day.date)}`, "From the Master file, under the filters above", 12);
  if (day.note) c.body.append(note(day.note));
  else if (!day.tasks || !day.tasks.rows.length) c.body.append(emptyMsg("No pole/task records for this date in the main dataset under the current filters."));
  else { $("p", c).textContent = `${intf(day.pole_count)} pole(s) · ${intf(day.tasks.rows.length)} task row(s)`; c.body.append(poleTaskTable(day.tasks)); }
  w.append(a, b, c); return w;
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
    const main = dcard("Poles disposed vs. forecasted", "District → Voltage → Project → Circuit / PID · each bar runs from Start Date to Finish Date, coloured by Status · scroll sideways for more dates · click a job for its details below", 12);
    const side = h(`<section class="dcard span12 jobpanel"></section>`);
    g.append(main, side);
    let mode = "gantt"; const holder = h(`<div></div>`);
    fc.zoom = fc.zoom || "month";
    const zoomSel = segc([["fit", "Fit"], ["quarter", "Quarter"], ["month", "Month"], ["week", "Week"]], fc.zoom, z => { fc.zoom = z; draw(); });
    zoomSel.title = "Zoom the timeline";
    main.acts.append(zoomSel, segc([["gantt", "Gantt"], ["table", "Table"]], mode, m => { mode = m; zoomSel.style.display = m === "gantt" ? "" : "none"; draw(); }));
    main.body.append(holder);
    if (fc.pick == null || !jobs.some(j => j.row === fc.pick)) fc.pick = jobs[0].row;
    const pick = row => { fc.pick = row; holder.querySelectorAll(".gjob").forEach(x => x.classList.toggle("on", +x.dataset.row === row)); jobPanel(side, jobs.find(j => j.row === row));
      const r = side.getBoundingClientRect(); if (r.top > innerHeight - 120) side.scrollIntoView({behavior: "smooth", block: "nearest"}); };
    function draw(){
      holder.textContent = "";
      if (mode === "table"){ holder.append(dtable({name: "Pole position", columns: ["District", "Voltage", "Project", "Circuit", "PID", "Status", "Start Date", "Finish Date", "Forecast", "Disposed", "Remaining", "Comment", "Control File", "Project link"],
        rows: jobs.map(j => [j.district, j.voltage, j.project, j.circuit, j.pid, j.status, j.start, j.finish, j.forecast, j.disposed, j.remaining, j.comment, j.control_file, j.project_link])})); return; }
      const counts = new Map(); jobs.forEach(j => { const key = statusKey(r.statuses, j); counts.set(key.label, (counts.get(key.label) || 0) + 1); });
      holder.append(legendEl(r.statuses.filter(s => counts.has(s.label)).map(s => ({name: `${s.label} (${counts.get(s.label)})`, color: s.colour}))));
      holder.append(ganttEl(jobs, fc.pick, pick, fc.zoom));
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
function ganttEl(jobs, picked, onPick, zoom){
  const starts = jobs.map(j => toDay(j.start)).filter(isFinite), ends = jobs.map(j => toDay(j.finish || j.start)).filter(isFinite);
  const el = h(`<div class="gantt"><div class="gscroll"><div class="gin"><div class="gh"><div class="glab">Circuit · PID</div><div class="gaxis"></div></div><div class="gbody"></div></div></div></div>`);
  if (!starts.length){ $(".gbody", el).append(emptyMsg("No Start Dates in these rows – map the Start Date column under Column mapping.")); return el; }
  let t0 = new Date(Math.min(...starts)), t1 = new Date(Math.max(...ends, ...starts) + 21 * dayMs);
  t0 = Date.UTC(t0.getUTCFullYear(), t0.getUTCMonth(), 1); t1 = Date.UTC(t1.getUTCFullYear(), t1.getUTCMonth() + 1, 1);
  const days = (t1 - t0) / dayMs, LAB = 200;
  const PPD = {quarter: 1.6, month: 4, week: 12};
  const scroller = $(".gscroll", el), inner = $(".gin", el), axis = $(".gaxis", el), body = $(".gbody", el);
  let ppd = 4, built = false;
  const now = Date.now();
  const sum = (arr, k) => arr.reduce((a, j) => a + (j[k] || 0), 0);
  const tot = arr => `${intf(sum(arr, "disposed"))} / ${intf(sum(arr, "forecast"))} poles`;
  const lab = x => x || "(blank)";
  const groupBy = (arr, key) => { const m = new Map(); arr.forEach(j => { const k = lab(j[key]); if (!m.has(k)) m.set(k, []); m.get(k).push(j); }); return [...m]; };
  function build(){
    ppd = zoom === "fit" ? Math.max(0.5, (scroller.clientWidth - LAB - 2) / days) : PPD[zoom] || 4;
    const W = Math.round(days * ppd), X = t => ((t - t0) / dayMs * ppd);
    inner.style.width = (LAB + W) + "px";
    el.style.setProperty("--glab", LAB + "px");
    const months = []; for (let d = new Date(t0); d.getTime() < t1; d = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1))) months.push(d.getTime());
    const monthPx = 30.4 * ppd, step = monthPx < 34 ? 3 : monthPx < 60 ? 2 : 1;
    axis.textContent = ""; axis.style.width = W + "px";
    months.forEach((m, i) => { const d = new Date(m);
      if (i % step === 0){ const sp = h(`<span></span>`); sp.style.left = X(m) + "px"; sp.textContent = `${MONTHS3[d.getUTCMonth()]}${d.getUTCMonth() === 0 || i === 0 ? " " + d.getUTCFullYear() : monthPx > 90 ? " " + String(d.getUTCFullYear()).slice(2) : ""}`; axis.append(sp); }
      if (d.getUTCMonth() === 0 && i){ const y = h(`<em></em>`); y.style.left = X(m) + "px"; axis.append(y); } });
    let weekLines = "";
    if (ppd >= 8){ const first = new Date(t0); const off = (8 - first.getUTCDay()) % 7; for (let t = t0 + off * dayMs; t < t1; t += 7 * dayMs){ weekLines += `<i class="wk" style="left:${X(t).toFixed(1)}px"></i>`;
      const sp = h(`<span class="wkl"></span>`); sp.style.left = X(t) + "px"; sp.textContent = new Date(t).getUTCDate(); axis.append(sp); } }
    const grid = months.map(m => `<i class="${new Date(m).getUTCMonth() === 0 ? "yr" : ""}" style="left:${X(m).toFixed(1)}px"></i>`).join("") + weekLines;
    const todayX = now >= t0 && now <= t1 ? `<b class="gtoday" style="left:${X(now).toFixed(1)}px"></b>` : "";
    if (todayX){ const sp = h(`<span class="today">Today</span>`); sp.style.left = X(now) + "px"; axis.append(sp); }
    const parts = [];
    groupBy(jobs, "district").forEach(([dn, dj]) => {
      parts.push(`<div class="g1"><span class="gst"><span>${esc(dn)}</span><span>${esc(tot(dj))}</span></span></div>`);
      groupBy(dj, "voltage").forEach(([vn, vj]) => {
        parts.push(`<div class="g2"><span class="gst"><span>${esc(vn)}</span><span>${esc(tot(vj))}</span></span></div>`);
        groupBy(vj, "project").forEach(([pn, pj]) => {
          const link = pj.find(j => j.project_link);
          parts.push(`<div class="g3"><span class="gst"><span>${link ? `<a href="${esc(link.project_link)}" target="_blank" rel="noopener" title="${esc(link.project_link)}">${esc(pn)} ${I.ext.replace("<svg", "<svg width=13 height=13")}</a>` : esc(pn)}</span><span>${esc(tot(pj))}</span></span></div>`);
          pj.forEach(j => {
            const s0 = toDay(j.start), f = toDay(j.finish);
            let barH = "";
            if (isFinite(s0)){
              const e = isFinite(f) && f >= s0 ? f + dayMs : s0 + 21 * dayMs;
              const left = X(s0), w = Math.max(4, X(e) - left), ink = inkOn(j.colour), prog = j.forecast ? Math.min(100, j.disposed / j.forecast * 100) : 0;
              barH = `<div class="gbar${isFinite(f) ? "" : " open"}" title="${isFinite(f) ? "" : "No Finish Date in the workbook"}" style="left:${left.toFixed(1)}px;width:${w.toFixed(1)}px;background:${j.colour};color:${ink}"><em style="width:${prog.toFixed(1)}%;background:${ink}"></em><span data-long="Disposed ${intf(j.disposed)} · Forecast ${intf(j.forecast)}" data-short="${intf(j.disposed)} / ${intf(j.forecast)}">Disposed ${intf(j.disposed)} · Forecast ${intf(j.forecast)}</span></div>`;
            } else barH = `<div class="gnodate">No Start Date · ${intf(j.disposed)} / ${intf(j.forecast)} poles</div>`;
            parts.push(`<div class="gjob${j.row === picked ? " on" : ""}" data-row="${j.row}"><div class="glab"><b>${esc(j.circuit || "—")}</b><small>${esc(j.pid ? "PID " + j.pid : "")}</small></div><div class="gtrack" style="width:${W}px">${grid}${todayX}${barH}</div></div>`);
          });
        });
      });
    });
    body.innerHTML = parts.join("");
    body.querySelectorAll(".gbar").forEach(bar => { const sp = $("span", bar);
      if (sp.scrollWidth > bar.clientWidth - 10){ sp.textContent = sp.dataset.short; if (sp.scrollWidth > bar.clientWidth - 10) bar.classList.add("outside"); } });
    if (!built && zoom !== "fit"){   // open on today (or the first job) instead of the far left
      const focus = now >= t0 && now <= t1 ? X(now) - scroller.clientWidth * 0.35 : 0;
      scroller.scrollLeft = Math.max(0, focus);
    }
    built = true;
  }
  let w0 = 0;
  new ResizeObserver(() => { const w = scroller.clientWidth; if (!w) return; el.style.setProperty("--vis", w + "px");
    if (built && (zoom !== "fit" || w === w0)) return; w0 = w; requestAnimationFrame(build); }).observe(scroller);
  body.addEventListener("click", e => { if (e.target.closest("a")) return; const r = e.target.closest(".gjob"); if (r) onPick(+r.dataset.row); });
  body.addEventListener("mousemove", e => { const r = e.target.closest(".gjob"); if (!r){ Tip.hide(); return; } const j = jobs.find(x => x.row === +r.dataset.row);
    Tip.show(e, {title: `${j.project} — ${j.circuit}`, rows: [{color: j.colour, name: "Status", value: j.status || "—"}, {name: "Start → Finish", value: `${fmtDate(j.start) || "—"} → ${fmtDate(j.finish) || "—"}`},
      {name: "Disposed / Forecast", value: `${intf(j.disposed)} / ${intf(j.forecast)}`}], note: "Click for the comment, control file and link"}); });
  body.addEventListener("mouseleave", () => Tip.hide());
  /* drag sideways with the mouse, as well as the scrollbar / shift+wheel */
  let drag = null;
  scroller.addEventListener("mousedown", e => { if (e.button !== 0 || e.target.closest("a,.glab")) return; drag = {x: e.clientX, l: scroller.scrollLeft, moved: false}; });
  addEventListener("mousemove", e => { if (!drag) return; const dx = e.clientX - drag.x; if (Math.abs(dx) > 4) drag.moved = true; if (drag.moved){ scroller.scrollLeft = drag.l - dx; scroller.classList.add("dragging"); } });
  addEventListener("mouseup", () => { if (drag && drag.moved) setTimeout(() => scroller.classList.remove("dragging"), 0); drag = null; });
  scroller.addEventListener("click", e => { if (scroller.classList.contains("dragging")) e.stopPropagation(); }, true);
  return el;
}
function linkish(v){ return /^(https?:|mailto:|file:)/i.test(v || "") ? v : /^\\\\/.test(v || "") ? "file:" + v.replace(/\\/g, "/") : ""; }
function jobPanel(side, j){
  side.textContent = "";
  if (!j){ side.append(emptyMsg("Click a job in the chart to see its details here.")); return; }
  const w = h(`<div class="jp"><div class="jph"><small>Job details · click another bar to switch</small><h3></h3><div class="jst"><i></i><span></span></div></div>
    <div class="jleft"><dl></dl><div class="jmeter"></div></div>
    <div class="jside"><div class="jpc"><h4>Comment (column L)</h4><p></p></div><div class="jpc jcf"><h4>Control File (column O)</h4><div></div></div></div></div>`);
  const hh = $("h3", w);
  if (j.project_link){ const a = h(`<a target="_blank" rel="noopener"></a>`); a.href = j.project_link; a.title = j.project_link; a.textContent = j.project + " ↗"; hh.append(a); } else hh.textContent = j.project;
  $(".jst i", w).style.background = j.colour; $(".jst span", w).textContent = j.status || "No status";
  if (j.project_link){ const b = h(`<a class="btn sm" target="_blank" rel="noopener">${I.ext} Open project link</a>`); b.href = j.project_link; $(".jph", w).append(b); }
  const dl = $("dl", w);
  [["District", j.district], ["Voltage", j.voltage], ["Circuit", j.circuit], ["PID", j.pid], ["Start Date", fmtDate(j.start) || "—"], ["Finish Date", fmtDate(j.finish) || "—"],
   ["Forecasted Total poles", intf(j.forecast)], ["Poles Disposed", intf(j.disposed)], ["Remaining", intf(j.remaining)], ["Workbook row", String(j.row)]].forEach(([k, v]) => {
    const d = document.createElement("div"), dt = document.createElement("dt"), dd = document.createElement("dd"); dt.textContent = k; dd.textContent = v || "—"; d.append(dt, dd); dl.append(d); });
  const pct = j.forecast ? Math.min(100, j.disposed / j.forecast * 100) : 0;
  const pm = h(`<div style="margin-top:14px"><div style="display:flex;justify-content:space-between;font-size:12.5px;color:var(--muted);margin-bottom:4px"><span>Poles disposed</span><b class="tnum" style="color:var(--ink)"></b></div><div class="meter"><i></i></div></div>`);
  $("b", pm).textContent = `${Math.round(pct)}%`; $("i", pm).style.width = pct + "%"; $(".jmeter", w).append(pm);
  $(".jpc p", w).textContent = j.comment || "No comment in column L for this job.";
  const cf = $(".jcf div", w);
  if (!j.control_file && !j.control_link) cf.textContent = "No control file in column O.";
  else {
    const target = j.control_link || j.control_file, href = linkish(target);
    const line = h(`<div class="jcfl"></div>`);
    if (href){ const a = h(`<a target="_blank" rel="noopener"></a>`); a.href = href; a.textContent = j.control_file || target; a.title = target; line.append(a); }
    else { const sp = h(`<span></span>`); sp.textContent = j.control_file; line.append(sp); }
    const cp = h(`<button class="btn sm" title="Copy to the clipboard">Copy</button>`);
    cp.onclick = () => { navigator.clipboard && navigator.clipboard.writeText(target).then(() => { cp.textContent = "Copied ✔"; setTimeout(() => cp.textContent = "Copy", 1500); }); };
    line.append(cp); cf.append(line);
    if (/^\\\\/.test(target)) cf.append(h(`<p class="muted" style="font-size:12px;margin:6px 0 0">Network paths may not open straight from the browser – use Copy and paste it into File Explorer.</p>`));
  }
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
  noComments: ["maps", "export"],
  loadFn: "mb_load", loadArgs: () => ({master_path: FILES.materials.path, control_path: FILES.master.path}),
  init(P){ const f = P.st.f, L = P.st.L; if (!L.date_fields.includes(f.date_field_label)) f.date_field_label = L.default_date_field || "Combined date";
    P.st.scope = P.st.scope || "All poles matching filters"; P.st.pdf = P.st.pdf || []; P.st.mapsMode = P.st.mapsMode || "Manual folder(s)"; },
  viewFn: "mb_view",
  mbFilters(P){ const f = P.st.f; return {date_field_label: f.date_field_label, date_from: f.date_from || null, date_to: f.date_to || null, districts: f.districts || [], projects: f.projects || [],
    pids: f.pids || [], sourcefiles: f.sourcefiles || [], circuits: f.circuits || [], enids: f.enids || []}; },
  viewArgs(P){ const s = P.st; return {master_path: FILES.materials.path, control_path: FILES.master.path, filters: DASHDEF.materials.mbFilters(P),
    fi_cats: s.fi_cats || [], fi_mats: s.fi_mats || [], guk_cats: s.guk_cats || [], guk_mats: s.guk_mats || [], guk_submats: s.guk_submats || []}; },
  cascade: [["districts"], ["projects"], ["pids"], ["sourcefiles"], ["circuits", "circuit_options"], ["enids", "pole_options"]],
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