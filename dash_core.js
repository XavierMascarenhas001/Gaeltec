/* ======================================================================
   Dashboards - shared pieces: loaded files, theme, formatting, tooltip,
   multi-select filters, data tables and SVG charts.
   ====================================================================== */
const NEED_MC = ["pyarrow"];
const NEED_NT = ["pyarrow"];
const NEED_MB = ["pyarrow", "pymupdf"];   // engine.py checks for PyMuPDF when it's imported

/* ---------- theme (Auto / Light / Dark) ---------- */
const Theme = {
  get(){ return store.get("theme", "auto"); },
  apply(){ const t = Theme.get(); if (t === "auto") delete document.documentElement.dataset.theme; else document.documentElement.dataset.theme = t; },
  next(){ const order = ["auto", "light", "dark"]; store.set("theme", order[(order.indexOf(Theme.get()) + 1) % 3]); Theme.apply(); },
  label(){ return {auto: "◐ Auto", light: "☀ Light", dark: "☾ Dark"}[Theme.get()]; },
};
Theme.apply();

/* ---------- files dropped for the dashboards (shared by all three) ---------- */
const FILES = {};
const SLOTS = {
  master:    {label: "Master file", hint: "Master_DD-MM-YYYY.parquet", types: [".parquet", ".csv"], icon: "db"},
  materials: {label: "Materials master", hint: "materials_all.parquet", types: [".parquet"], icon: "box3"},
  outages:   {label: "Outages Programme", hint: "High-level_planning_2026.xlsx", types: [".xlsx", ".xlsm"], icon: "grid"},
  forecast:  {label: "Service Partner Workbank", hint: "Workbank .xlsx", types: [".xlsx", ".xlsm"], icon: "chart"},
  ics:       {label: "Outage calendar", hint: "outage_calendar_2026 (vN).ics", types: [".ics"], icon: "file"},
};
const fileListeners = new Set();
async function setSlotFile(slot, file){
  const exts = SLOTS[slot].types;
  if (!exts.includes(extOf(file.name))) { alert(`“${file.name}” isn't a ${exts.join(" / ")} file.`); return false; }
  const dir = await PY.mount([{rel: file.name, file}]);
  const old = FILES[slot];
  FILES[slot] = {name: file.name, size: file.size, dir, path: dir + "/" + file.name, v: Date.now()};
  if (old) PY.unmount(old.dir);
  fileListeners.forEach(cb => cb(slot));
  return true;
}
function clearSlot(slot){
  const old = FILES[slot]; if (!old) return;
  delete FILES[slot]; PY.unmount(old.dir);
  fileListeners.forEach(cb => cb(slot));
}
function pickFile(accept){
  return new Promise(res => {
    const inp = document.createElement("input"); inp.type = "file"; inp.accept = accept;
    inp.onchange = () => res(inp.files[0] || null); inp.click();
  });
}
/* compact drop target used in the dashboard's file strip */
function slotEl(slot, required){
  const s = SLOTS[slot];
  const el = h(`<div class="slot" tabindex="0" title="Drop a file here or click to choose"><div class="si">${I[s.icon] || I.file}</div>
    <div class="st"><small>${esc(s.label)}</small><span></span></div></div>`);
  const txt = $(".st span", el);
  const render = () => {
    const f = FILES[slot];
    el.classList.toggle("has", !!f);
    el.querySelectorAll(".x,.req").forEach(n => n.remove());
    if (f){ txt.textContent = f.name; txt.className = ""; el.title = `${f.name} · ${fmtSize(f.size)} – click or drop to replace`;
      const x = h(`<button class="x" title="Remove">×</button>`); x.onclick = e => { e.stopPropagation(); clearSlot(slot); }; el.append(x); }
    else { txt.textContent = "Drop " + s.hint; txt.className = "none"; if (required) el.append(h(`<span class="req">Required</span>`)); }
  };
  const take = async file => { if (!file) return; el.classList.add("busy"); try { await setSlotFile(slot, file); } catch(e){ alert("Couldn't load the file: " + e.message); } el.classList.remove("busy"); };
  el.onclick = async () => take(await pickFile(s.types.join(",")));
  el.onkeydown = e => { if (e.key === "Enter" || e.key === " "){ e.preventDefault(); el.click(); } };
  el.ondragover = e => { e.preventDefault(); el.classList.add("over"); };
  el.ondragleave = () => el.classList.remove("over");
  el.ondrop = async e => { e.preventDefault(); el.classList.remove("over"); const ents = await entriesFromDataTransfer(e.dataTransfer); take(ents[0] && ents[0].file); };
  const cb = k => { if (k === slot){ if (!el.isConnected) fileListeners.delete(cb); else render(); } };
  fileListeners.add(cb);
  render();
  return el;
}
/* the big drop zone used on a dashboard's landing screen */
function slotZone(slot, required, note){
  const s = SLOTS[slot];
  const el = h(`<div class="field"><label>${esc(s.label)}${required ? "" : " <span class='muted' style='font-weight:400'>(optional)</span>"}</label>
    <div class="dz" tabindex="0"><div class="dzi">${I[s.icon] || I.drop}</div>
    <div class="dzt"><span><b>Drop the file here</b> or <u>click to choose</u></span><small></small></div></div></div>`);
  const dz = $(".dz", el), sm = $("small", el);
  const render = () => { const f = FILES[slot]; dz.classList.toggle("has", !!f); sm.textContent = f ? `✔ ${f.name} · ${fmtSize(f.size)}` : (note || s.hint); };
  const take = async file => { if (!file) return; dz.classList.add("busy"); try { await setSlotFile(slot, file); } catch(e){ alert("Couldn't load the file: " + e.message); } dz.classList.remove("busy"); };
  dz.onclick = async () => take(await pickFile(s.types.join(",")));
  dz.onkeydown = e => { if (e.key === "Enter" || e.key === " "){ e.preventDefault(); dz.click(); } };
  dz.ondragover = e => { e.preventDefault(); dz.classList.add("over"); };
  dz.ondragleave = () => dz.classList.remove("over");
  dz.ondrop = async e => { e.preventDefault(); dz.classList.remove("over"); const ents = await entriesFromDataTransfer(e.dataTransfer); take(ents[0] && ents[0].file); };
  const cb = k => { if (k === slot){ if (!el.isConnected) fileListeners.delete(cb); else render(); } };
  fileListeners.add(cb); render();
  return el;
}

/* ---------- formatting ---------- */
const nf2 = new Intl.NumberFormat("en-GB", {minimumFractionDigits: 2, maximumFractionDigits: 2});
const nf0 = new Intl.NumberFormat("en-GB", {maximumFractionDigits: 0});
const nfq = new Intl.NumberFormat("en-GB", {maximumFractionDigits: 3});
const isNum = v => typeof v === "number" && isFinite(v);
const money = v => !isNum(v) ? "—" : (v < 0 ? "-£" : "£") + nf2.format(Math.abs(v));
const moneyS = v => { if (!isNum(v)) return "—"; const a = Math.abs(v), s = v < 0 ? "-£" : "£";
  return a >= 1e6 ? s + (a/1e6).toFixed(a >= 1e7 ? 1 : 2).replace(/\.?0+$/, "") + "m" : a >= 1e3 ? s + (a/1e3).toFixed(a >= 1e4 ? 0 : 1).replace(/\.0$/, "") + "k" : s + nf0.format(a); };
const signed = (v, f = money) => !isNum(v) ? "—" : (v > 0 ? "+" : "") + f(v);
const intf = v => !isNum(v) ? "—" : nf0.format(v);
const qty = v => !isNum(v) ? (v ?? "") : nfq.format(v);
const compact = v => { const a = Math.abs(v); return a >= 1e6 ? (v/1e6).toFixed(1).replace(/\.0$/, "") + "m" : a >= 1e4 ? (v/1e3).toFixed(0) + "k" : a >= 1e3 ? (v/1e3).toFixed(1).replace(/\.0$/, "") + "k" : nf0.format(v); };
const MONTHS3 = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const fmtDate = s => { if (!s) return ""; const [y, m, d] = String(s).slice(0, 10).split("-"); return d ? `${+d} ${MONTHS3[+m-1]} ${y}` : s; };

/* ---------- colour that follows the entity, never its rank ---------- */
const SLOT_VARS = ["--s1","--s2","--s3","--s4","--s5","--s6","--s7","--s8"];
function colorKeeper(){
  const map = new Map(); let lastSeen = new Map(), tick = 0;
  return {
    /* keys = the entities on screen now; returns key -> css colour */
    assign(keys){
      tick++;
      keys.forEach(k => lastSeen.set(k, tick));
      const used = new Set([...map.values()]);
      for (const k of keys){
        if (map.has(k)) continue;
        let slot = SLOT_VARS.find(v => !used.has(v));
        if (!slot){   // all 8 taken: reuse the slot of an entity that isn't on screen now
          const free = [...map].filter(([kk]) => !keys.includes(kk)).sort((a, b) => lastSeen.get(a[0]) - lastSeen.get(b[0]))[0];
          slot = free ? free[1] : SLOT_VARS[0]; if (free) map.delete(free[0]);
        }
        map.set(k, slot); used.add(slot);
      }
      return k => `var(${map.get(k) || "--neutral"})`;
    },
  };
}

/* ---------- tooltip (built with textContent) ---------- */
const Tip = (() => {
  const el = document.createElement("div"); el.className = "vtip"; document.body.append(el);
  function place(ev){
    const pad = 14, r = el.getBoundingClientRect();
    let x = ev.clientX + pad, y = ev.clientY + pad;
    if (x + r.width > innerWidth - 8) x = ev.clientX - r.width - pad;
    if (y + r.height > innerHeight - 8) y = ev.clientY - r.height - pad;
    el.style.left = Math.max(4, x) + "px"; el.style.top = Math.max(4, y) + "px";
  }
  /* spec: {title, rows:[{color, name, value}], total:{name, value}, note} */
  function show(ev, spec){
    el.textContent = "";
    if (spec.title){ const t = document.createElement("div"); t.className = "th"; t.textContent = spec.title; el.append(t); }
    const row = (r, cls) => { const d = document.createElement("div"); d.className = "tr" + (cls ? " " + cls : "");
      const s = document.createElement("span"); if (r.color){ const i = document.createElement("i"); i.style.background = r.color; s.append(i); }
      s.append(document.createTextNode(r.name)); const b = document.createElement("b"); b.textContent = r.value; d.append(s, b); el.append(d); };
    (spec.rows || []).forEach(r => row(r));
    if (spec.total) row(spec.total, "tot");
    if (spec.note){ const n = document.createElement("div"); n.style.opacity = ".75"; n.style.marginTop = "3px"; n.textContent = spec.note; el.append(n); }
    el.classList.add("on"); place(ev);
  }
  return {show, move: place, hide(){ el.classList.remove("on"); }};
})();

/* ---------- small UI pieces ---------- */
function segc(options, value, onChange){
  const el = h(`<div class="segc" role="group"></div>`);
  const render = v => el.querySelectorAll("button").forEach(b => b.classList.toggle("on", b.dataset.v === v));
  options.forEach(o => { const [v, l] = Array.isArray(o) ? o : [o, o];
    const b = h(`<button type="button"></button>`); b.dataset.v = v; b.textContent = l;
    b.onclick = () => { if (b.classList.contains("on")) return; render(v); onChange(v); }; el.append(b); });
  render(value); el.set = render;
  return el;
}
function dcard(title, sub, span = 12){
  const el = h(`<section class="dcard span${span}"><div class="ch"><div class="tt"><h3></h3>${sub ? "<p></p>" : ""}</div><div class="acts"></div></div><div class="cb"></div></section>`);
  $("h3", el).textContent = title; if (sub) $("p", el).textContent = sub;
  el.body = $(".cb", el); el.acts = $(".acts", el);
  return el;
}
function note(text, kind = "", icon = "info"){
  const el = h(`<div class="dnote ${kind}">${ICON2[icon] || ICON2.info}<div></div></div>`);
  $("div", el).textContent = text; return el;
}
function emptyMsg(text){ const e = h(`<div class="empty"></div>`); e.textContent = text; return e; }
/* stat tiles: [{label, value, hero, note, tone:'up'|'down', swatch}] - one hero per view */
function kpis(items){
  const el = h(`<div class="kpis"></div>`);
  items.forEach(k => {
    const t = h(`<div class="kpi${k.hero ? " hero" : ""}"><div class="kl"></div><div class="kv"></div></div>`);
    const kl = $(".kl", t);
    if (k.swatch){ const i = document.createElement("i"); i.style.background = k.swatch; kl.append(i); }
    kl.append(document.createTextNode(k.label));
    $(".kv", t).textContent = k.value;
    if (k.title) t.title = k.title;
    if (k.note != null){ const d = h(`<div class="kd"></div>`); if (k.tone) d.classList.add(k.tone); d.textContent = k.note; t.append(d); }
    el.append(t);
  });
  return el;
}
const ICON2 = {
  info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><path d="M12 8h.01"/></svg>',
  warn: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3 2 20h20z"/><path d="M12 10v4"/><path d="M12 17h.01"/></svg>',
  err: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6"/><path d="m15 9-6 6"/></svg>',
  up: '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M12 5 20 17H4z"/></svg>',
  down: '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M12 19 4 7h16z"/></svg>',
  pole: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 2v20"/><path d="M5 6h14"/><path d="M7 10h10"/><path d="M5 6v2"/><path d="M19 6v2"/><path d="M8 22h8"/></svg>',
  tx: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><rect x="5" y="7" width="14" height="12" rx="2"/><path d="M9 7V4"/><path d="M15 7V4"/><path d="m11 10-2 4h4l-2 4"/></svg>',
  cable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 7c4 0 4 10 9 10s5-10 9-10"/><circle cx="3" cy="7" r="1"/><circle cx="21" cy="7" r="1"/></svg>',
  switch: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="5" cy="16" r="2"/><circle cx="19" cy="16" r="2"/><path d="m7 15 10-7"/><path d="M2 16h1"/><path d="M21 16h1"/></svg>',
  cal: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18"/><path d="M8 3v4"/><path d="M16 3v4"/></svg>',
  table: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18"/><path d="M9 10v10"/></svg>',
  trend: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m3 17 6-6 4 4 8-8"/><path d="M15 7h6v6"/></svg>',
  csv: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/></svg>',
};

/* ---------- multi-select with search (replaces st.multiselect) ---------- */
let openMs = null;
document.addEventListener("mousedown", e => { if (openMs && !openMs.contains(e.target)) openMs.close(); });
document.addEventListener("keydown", e => { if (e.key === "Escape" && openMs) openMs.close(); });
function msel({label, options, value = [], onChange, cascade, hint, right}){
  const el = h(`<div class="ms"><button type="button" class="btn${cascade ? " cascade" : ""}"><span></span><span class="n all">All</span><span style="opacity:.6">▾</span></button></div>`);
  const btn = $("button", el), cnt = $(".n", el);
  $("span", btn).textContent = label;
  let opts = options || [], sel = new Set(value), pop = null, dirty = false;
  const upd = () => { cnt.textContent = sel.size ? String(sel.size) : "All"; cnt.classList.toggle("all", !sel.size); btn.title = sel.size ? [...sel].join(", ") : `${label}: all`; };
  el.close = () => { if (!pop) return; pop.remove(); pop = null; openMs = null; if (dirty){ dirty = false; onChange && onChange([...sel]); } };
  btn.onclick = () => {
    if (pop){ el.close(); return; }
    if (openMs) openMs.close();
    pop = h(`<div class="mspop${right ? " right" : ""}"><input type="text" placeholder="Search ${esc(label.toLowerCase())}…"><div class="opts"></div><div class="acts"><span class="cnt"></span><button class="btn sm">Select shown</button><button class="btn sm">Clear</button><button class="btn sm primary">Done</button></div></div>`);
    if (hint){ const p = h(`<p class="note"></p>`); p.textContent = hint; pop.insertBefore(p, pop.firstChild); }
    const q = $("input", pop), list = $(".opts", pop), c2 = $(".acts .cnt", pop);
    const shown = () => { const t = q.value.trim().toLowerCase(); return t ? opts.filter(o => String(o).toLowerCase().includes(t)) : opts; };
    const draw = () => {
      list.textContent = "";
      const vis = shown(), frag = document.createDocumentFragment();
      /* selected-but-hidden values (e.g. after a cascade) stay visible so they can be unticked */
      const extra = [...sel].filter(v => !opts.includes(v));
      [...extra, ...vis.slice(0, 400)].forEach(v => {
        const l = h(`<label><input type="checkbox"><span></span></label>`); $("span", l).textContent = v; l.title = v;
        const cb = $("input", l); cb.checked = sel.has(v);
        cb.onchange = () => { cb.checked ? sel.add(v) : sel.delete(v); dirty = true; upd(); c2.textContent = `${sel.size} selected`; };
        frag.append(l);
      });
      list.append(frag);
      if (!vis.length && !extra.length) list.append(h(`<div class="more">No matches</div>`));
      if (vis.length > 400){ const m = h(`<div class="more"></div>`); m.textContent = `…${intf(vis.length - 400)} more – keep typing to narrow`; list.append(m); }
      c2.textContent = `${sel.size} selected · ${intf(opts.length)} options`;
    };
    q.oninput = draw;
    q.onkeydown = e => { if (e.key === "Enter"){ const vis = shown(); if (vis.length === 1){ sel.add(vis[0]); dirty = true; upd(); draw(); } } };
    const [bAll, bClr, bOk] = pop.querySelectorAll(".acts .btn");
    bAll.onclick = () => { shown().forEach(v => sel.add(v)); dirty = true; upd(); draw(); };
    bClr.onclick = () => { if (sel.size){ sel.clear(); dirty = true; upd(); draw(); } };
    bOk.onclick = () => el.close();
    el.append(pop); openMs = el; draw(); q.focus();
  };
  el.setOptions = o => { opts = o || []; };
  el.set = v => { sel = new Set(v || []); upd(); };
  el.get = () => [...sel];
  upd();
  return el;
}

/* ---------- data table: sort, search, show more, CSV ---------- */
function toCSV(columns, rows){
  const q = v => { const s = v == null ? "" : String(v); return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  return [columns.map(q).join(","), ...rows.map(r => r.map(q).join(","))].join("\r\n");
}
function downloadCSV(name, columns, rows){ downloadBlob(name.replace(/[\\/:*?"<>|]+/g, "_") + ".csv", new TextEncoder().encode("﻿" + toCSV(columns, rows))); }
/* spec: {columns, rows, money:[cols], signed:[cols], name, page, height, search, pct:[cols]} */
function dtable(spec){
  const {columns, rows} = spec;
  const money_ = new Set(spec.money || columns.filter(c => /£/.test(c)));
  const signed_ = new Set(spec.signed || columns.filter(c => /difference|variance/i.test(c)));
  const pct_ = new Set(spec.pct || []);
  const numeric = columns.map((c, i) => { let n = 0, any = 0; for (const r of rows.slice(0, 300)){ const v = r[i]; if (v == null || v === "") continue; any++; if (isNum(v)) n++; } return any > 0 && n === any; });
  const el = h(`<div class="dt"><div class="dth">${spec.search === false ? "" : `<input type="text" placeholder="Search this table…">`}<span class="cnt"></span><button class="btn sm" title="Download these rows as CSV">${ICON2.csv} CSV</button></div><div class="tw"><table><thead><tr></tr></thead><tbody></tbody></table></div><div class="foot"></div></div>`);
  if (spec.height) $(".tw", el).style.maxHeight = spec.height + "px";
  const thr = $("thead tr", el), tb = $("tbody", el), cnt = $(".cnt", el), foot = $(".foot", el), q = $("input", el);
  let sortI = -1, sortDir = 1, limit = spec.page || 200, view = rows;
  const fmt = (v, i) => {
    const c = columns[i];
    if (v == null) return "";
    if (!isNum(v)) return String(v);
    if (pct_.has(c)) return (v > 0 ? "+" : "") + v.toFixed(1) + "%";
    if (money_.has(c)) return signed_.has(c) ? signed(v) : money(v);
    return signed_.has(c) ? signed(v, qty) : qty(v);
  };
  columns.forEach((c, i) => { const th = document.createElement("th"); th.textContent = c; if (numeric[i]) th.className = "nr";
    const ar = document.createElement("span"); ar.className = "ar"; th.append(ar);
    th.onclick = () => { sortDir = sortI === i ? -sortDir : (numeric[i] ? -1 : 1); sortI = i; apply(); }; thr.append(th); });
  function apply(){
    const t = q ? q.value.trim().toLowerCase() : "";
    view = t ? rows.filter(r => r.some(v => v != null && String(v).toLowerCase().includes(t))) : rows.slice();
    if (sortI >= 0){ const i = sortI; view.sort((a, b) => { const x = a[i], y = b[i];
      if (x == null || x === "") return 1; if (y == null || y === "") return -1;
      return (isNum(x) && isNum(y) ? x - y : String(x).localeCompare(String(y), undefined, {numeric: true})) * sortDir; }); }
    thr.querySelectorAll(".ar").forEach((a, i) => a.textContent = i === sortI ? (sortDir > 0 ? "▲" : "▼") : "");
    draw();
  }
  function draw(){
    const part = view.slice(0, limit);
    tb.innerHTML = part.map(r => "<tr>" + r.map((v, i) => {
      const cls = [numeric[i] ? "nr" : "", signed_.has(columns[i]) && isNum(v) ? (v > 0 ? "pos" : v < 0 ? "neg" : "") : ""].filter(Boolean).join(" ");
      const s = esc(fmt(v, i)); return `<td${cls ? ` class="${cls}"` : ""} title="${s}">${s}</td>`; }).join("") + "</tr>").join("")
      || `<tr><td colspan="${columns.length}" class="muted" style="text-align:center;padding:18px">No rows</td></tr>`;
    cnt.textContent = view.length === rows.length ? `${intf(rows.length)} row${rows.length === 1 ? "" : "s"}` : `${intf(view.length)} of ${intf(rows.length)} rows`;
    foot.textContent = "";
    if (view.length > limit){ const b = h(`<button class="btn sm"></button>`); b.textContent = `Show ${intf(Math.min(limit, view.length - limit))} more (${intf(view.length - limit)} hidden)`; b.onclick = () => { limit *= 2; draw(); }; foot.append(b); }
  }
  if (q) q.oninput = () => { limit = spec.page || 200; apply(); };
  $(".dth .btn", el).onclick = () => downloadCSV(spec.name || "table", columns, view);
  apply();
  return el;
}
/* card whose body flips between a chart and the same data as a table */
function vizCard({title, sub, span = 12, draw, table, legend, extra}){
  const c = dcard(title, sub, span);
  const holder = h(`<div></div>`);
  let mode = "chart";
  const render = () => {
    holder.textContent = "";
    if (mode === "chart"){ if (legend) holder.append(legendEl(legend)); const host = h(`<div class="viz"></div>`); holder.append(host); draw(host); }
    else holder.append(dtable(table()));
  };
  if (extra) c.acts.append(extra);
  if (table) c.acts.append(segc([["chart", "Chart"], ["table", "Table"]], "chart", v => { mode = v; render(); }));
  c.body.append(holder);
  requestAnimationFrame(render);
  return c;
}
function legendEl(items){
  const el = h(`<div class="legend"></div>`);
  items.forEach(it => { const s = document.createElement("span"); const i = document.createElement("i"); if (it.shape) i.className = it.shape; i.style.background = it.color; s.append(i, document.createTextNode(it.name)); el.append(s); });
  return el;
}

/* ======================================================================
   SVG charts - thin marks, 4px rounded data ends, 2px surface gaps,
   hairline grid, one axis, hover tooltip on every mark.
   ====================================================================== */
const Viz = (() => {
  const ctx = document.createElement("canvas").getContext("2d");
  const tw = (s, px = 11.5) => { ctx.font = `${px}px system-ui,-apple-system,"Segoe UI",sans-serif`; return ctx.measureText(String(s)).width; };
  function nice(max, count = 5, min = 0){
    const span = (max - min) || Math.abs(max) || 1, raw = span / count, p = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map(m => m * p).find(s => s >= raw) || 10 * p;
    const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step || step;
    const ticks = []; for (let v = lo; v <= hi + step / 2; v += step) ticks.push(+v.toFixed(10));
    return {lo, hi, ticks};
  }
  const trunc = (s, w, px = 11.5) => { s = String(s ?? ""); if (tw(s, px) <= w) return s; while (s.length > 1 && tw(s + "…", px) > w) s = s.slice(0, -1); return s + "…"; };
  /* column with rounded data end, square at the baseline */
  const colPath = (x, yTop, w, yBase, r) => { const hgt = yBase - yTop; r = Math.min(r, w / 2, hgt); if (r <= 0.5) return `M${x},${yBase}V${yTop}H${x+w}V${yBase}Z`;
    return `M${x},${yBase}V${yTop + r}Q${x},${yTop} ${x + r},${yTop}H${x + w - r}Q${x + w},${yTop} ${x + w},${yTop + r}V${yBase}Z`; };
  const rowPath = (xBase, y, xEnd, hgt, r) => { const w = xEnd - xBase; r = Math.min(r, hgt / 2, Math.abs(w)); if (r <= 0.5) return `M${xBase},${y}H${xEnd}V${y+hgt}H${xBase}Z`;
    if (w >= 0) return `M${xBase},${y}H${xEnd - r}Q${xEnd},${y} ${xEnd},${y + r}V${y + hgt - r}Q${xEnd},${y + hgt} ${xEnd - r},${y + hgt}H${xBase}Z`;
    return `M${xBase},${y}H${xEnd + r}Q${xEnd},${y} ${xEnd},${y + r}V${y + hgt - r}Q${xEnd},${y + hgt} ${xEnd + r},${y + hgt}H${xBase}Z`; };
  function mount(host, render){
    let w0 = 0, raf = 0;
    const go = () => { const w = host.clientWidth; if (!w || w === w0) return; w0 = w; render(w); };
    const ro = new ResizeObserver(() => { cancelAnimationFrame(raf); raf = requestAnimationFrame(go); });
    ro.observe(host); go();
  }
  function wire(svgHost, onEnter){
    svgHost.addEventListener("mousemove", e => { const t = e.target.closest("[data-i]"); if (!t) { Tip.hide(); return; } onEnter(e, t); });
    svgHost.addEventListener("mouseleave", () => Tip.hide());
  }

  /* vertical stacked columns.  o = {labels, series:[{name, values, color|colorAt(i)}], fmt, axis, height,
     fill (histogram: columns touch with 2px gaps), tip(i) -> tooltip spec, onClick(i), marker:{x (index-ish float), label}} */
  function columns(host, o){
    mount(host, W => {
      const n = o.labels.length, H = o.height || 300;
      const sums = o.labels.map((_, i) => o.series.reduce((a, s) => a + Math.max(0, s.values[i] || 0), 0));
      const sc = nice(Math.max(...sums, 0) || 1, 5);
      const axis = o.axis || compact;
      const L = Math.max(...sc.ticks.map(t => tw(axis(t)))) + 12, R = 6, T = 10;
      const pw = W - L - R, band = pw / Math.max(n, 1);
      const maxLab = Math.max(...o.labels.map(l => tw(l)));
      const rotate = !o.fill && maxLab + 10 > band && n > 1;
      const every = o.fill ? Math.ceil(46 / band) : rotate ? Math.max(1, Math.ceil(16 / band)) : 1;
      const B = rotate ? Math.min(110, maxLab * 0.72 + 16) : 24;
      const ph = H - T - B, y = v => T + ph - (v / sc.hi) * ph;
      const bw = o.fill ? Math.max(1, band - 2) : Math.min(24, Math.max(3, band * 0.62));
      let s = `<svg height="${H}" viewBox="0 0 ${W} ${H}" role="img">`;
      s += `<g class="gl">` + sc.ticks.map(t => `<line x1="${L}" x2="${W - R}" y1="${y(t)}" y2="${y(t)}"/>`).join("") + `</g>`;
      s += `<g class="tick">` + sc.ticks.map(t => `<text x="${L - 8}" y="${y(t) + 4}" text-anchor="end">${esc(axis(t))}</text>`).join("") + `</g>`;
      for (let i = 0; i < n; i++){
        const x0 = L + band * i + (band - bw) / 2;
        let acc = 0; const segs = [];
        o.series.forEach((se, k) => { const v = Math.max(0, se.values[i] || 0); if (v > 0) segs.push({k, v}); });
        segs.forEach((g, j) => {
          const yb = y(acc) - (j > 0 ? 2 : 0), yt = y(acc + g.v); acc += g.v;
          if (yb - yt < 0.5) return;
          const se = o.series[g.k], col = se.colorAt ? se.colorAt(i) : se.color;
          const d = j === segs.length - 1 ? colPath(x0, yt, bw, yb, 4) : `M${x0},${yb}V${yt}H${x0 + bw}V${yb}Z`;
          s += `<path d="${d}" style="fill:${col}"/>`;
        });
        s += `<rect class="hit${o.onClick ? " clickable" : ""}" data-i="${i}" x="${L + band * i}" y="${T}" width="${band}" height="${ph}"/>`;
        if (i % every === 0){
          const lab = esc(rotate ? trunc(o.labels[i], 150) : trunc(o.labels[i], band * every - 4));
          const cx = L + band * i + band / 2;
          s += rotate ? `<text x="${cx}" y="${T + ph + 12}" text-anchor="end" transform="rotate(-40 ${cx} ${T + ph + 12})">${lab}</text>`
                      : `<text x="${o.fill ? L + band * i : cx}" y="${T + ph + 16}" text-anchor="${o.fill ? "start" : "middle"}">${lab}</text>`;
        }
      }
      if (o.marker){ const mx = L + band * o.marker.x; s += `<line x1="${mx}" x2="${mx}" y1="${T}" y2="${T + ph}" style="stroke:var(--ink);stroke-width:1.5"/><text class="lbl" x="${mx + 5}" y="${T + 11}">${esc(o.marker.label)}</text>`; }
      s += `<line class="base" x1="${L}" x2="${W - R}" y1="${T + ph}" y2="${T + ph}"/></svg>`;
      host.innerHTML = s;
    });
    wire(host, (e, t) => { const i = +t.dataset.i; Tip.show(e, o.tip ? o.tip(i) : {title: o.labels[i], rows: o.series.map(se => ({color: se.colorAt ? se.colorAt(i) : se.color, name: se.name, value: (o.fmt || qty)(se.values[i] || 0)}))}); });
    if (o.onClick) host.addEventListener("click", e => { const t = e.target.closest("[data-i]"); if (t) o.onClick(+t.dataset.i); });
  }

  /* horizontal stacked bars, value at the tip.  o = {labels, series, fmt, rowH, tip(i), valueLabel(i)} */
  function bars(host, o){
    mount(host, W => {
      const n = o.labels.length, rowH = o.rowH || 30, th = Math.min(18, rowH - 10), T = 4, B = 22;
      const sums = o.labels.map((_, i) => o.series.reduce((a, s) => a + Math.max(0, s.values[i] || 0), 0));
      const sc = nice(Math.max(...sums, 0) || 1, 4);
      const fmt = o.fmt || qty, vlab = o.valueLabel || (i => fmt(sums[i]));
      const Lw = o.labels.every(l => !l) ? 2 : Math.min(o.labelMax || 300, Math.max(60, W * 0.34), Math.max(...o.labels.map(l => tw(l))) + 12);
      const Rw = Math.max(...o.labels.map((_, i) => tw(vlab(i)))) + 10;
      const pw = W - Lw - Rw, x = v => Lw + (v / sc.hi) * pw, H = T + n * rowH + B;
      let s = `<svg height="${H}" viewBox="0 0 ${W} ${H}" role="img"><g class="gl">` + sc.ticks.map(t => `<line x1="${x(t)}" x2="${x(t)}" y1="${T}" y2="${T + n * rowH}"/>`).join("") + `</g>`;
      s += `<g class="tick">` + sc.ticks.map(t => `<text x="${x(t)}" y="${H - 6}" text-anchor="middle">${esc((o.axis || compact)(t))}</text>`).join("") + `</g>`;
      for (let i = 0; i < n; i++){
        const yy = T + i * rowH + (rowH - th) / 2;
        let acc = 0; const segs = []; o.series.forEach((se, k) => { const v = Math.max(0, se.values[i] || 0); if (v > 0) segs.push({k, v}); });
        segs.forEach((g, j) => { const x0 = x(acc) + (j > 0 ? 2 : 0), x1 = x(acc + g.v); acc += g.v; if (x1 - x0 < 0.5) return;
          const se = o.series[g.k], col = se.colorAt ? se.colorAt(i) : se.color;
          s += `<path d="${j === segs.length - 1 ? rowPath(x0, yy, x1, th, 4) : `M${x0},${yy}H${x1}V${yy + th}H${x0}Z`}" style="fill:${col}"/>`; });
        s += `<text x="${Lw - 8}" y="${yy + th / 2 + 4}" text-anchor="end" style="fill:var(--ink2)">${esc(trunc(o.labels[i], Lw - 12))}</text>`;
        s += `<text class="lbl tnum" x="${x(sums[i]) + 6}" y="${yy + th / 2 + 4}">${esc(vlab(i))}</text>`;
        s += `<rect class="hit" data-i="${i}" x="0" y="${T + i * rowH}" width="${W}" height="${rowH}"/>`;
      }
      s += `<line class="base" x1="${Lw}" x2="${Lw}" y1="${T}" y2="${T + n * rowH}"/></svg>`;
      host.innerHTML = s;
    });
    wire(host, (e, t) => { const i = +t.dataset.i; Tip.show(e, o.tip ? o.tip(i) : {title: o.labels[i], rows: o.series.map(se => ({color: se.colorAt ? se.colorAt(i) : se.color, name: se.name, value: (o.fmt || qty)(se.values[i] || 0)}))}); });
  }

  /* diverging bars around zero (variance).  o = {labels, values, fmt, valueLabel(i), tip(i), pos, neg} */
  function diverging(host, o){
    mount(host, W => {
      const n = o.labels.length, rowH = o.rowH || 30, th = Math.min(18, rowH - 10), T = 4, B = 22;
      const vmax = Math.max(0, ...o.values), vmin = Math.min(0, ...o.values);
      const sc = nice(vmax || (vmin ? 0 : 1), 4, vmin);
      const fmt = o.fmt || qty, vlab = o.valueLabel || (i => signed(o.values[i], fmt));
      const Lw = Math.min(o.labelMax || 280, Math.max(60, W * 0.3), Math.max(...o.labels.map(l => tw(l))) + 12);
      const lab = Math.max(...o.labels.map((_, i) => tw(vlab(i)))) + 10;
      const x0 = Lw + (sc.lo < 0 ? lab : 0), x1 = W - lab;
      const x = v => x0 + (v - sc.lo) / ((sc.hi - sc.lo) || 1) * (x1 - x0), H = T + n * rowH + B, zx = x(0);
      let s = `<svg height="${H}" viewBox="0 0 ${W} ${H}" role="img"><g class="gl">` + sc.ticks.map(t => `<line x1="${x(t)}" x2="${x(t)}" y1="${T}" y2="${T + n * rowH}"/>`).join("") + `</g>`;
      s += `<g class="tick">` + sc.ticks.map(t => `<text x="${x(t)}" y="${H - 6}" text-anchor="middle">${esc((o.axis || compact)(t))}</text>`).join("") + `</g>`;
      for (let i = 0; i < n; i++){
        const v = o.values[i] || 0, yy = T + i * rowH + (rowH - th) / 2, xe = x(v);
        if (Math.abs(xe - zx) >= 0.5) s += `<path d="${rowPath(zx, yy, xe, th, 4)}" style="fill:${v >= 0 ? (o.pos || "var(--s1)") : (o.neg || "var(--s8)")}"/>`;
        s += `<text x="${Lw - 8}" y="${yy + th / 2 + 4}" text-anchor="end" style="fill:var(--ink2)">${esc(trunc(o.labels[i], Lw - 12))}</text>`;
        s += `<text class="lbl tnum" x="${v >= 0 ? xe + 6 : xe - 6}" y="${yy + th / 2 + 4}" text-anchor="${v >= 0 ? "start" : "end"}">${esc(vlab(i))}</text>`;
        s += `<rect class="hit" data-i="${i}" x="0" y="${T + i * rowH}" width="${W}" height="${rowH}"/>`;
      }
      s += `<line class="base" x1="${zx}" x2="${zx}" y1="${T}" y2="${T + n * rowH}"/></svg>`;
      host.innerHTML = s;
    });
    wire(host, (e, t) => { const i = +t.dataset.i; Tip.show(e, o.tip ? o.tip(i) : {title: o.labels[i], rows: [{name: "Value", value: signed(o.values[i], o.fmt || qty)}]}); });
  }

  /* line with a 10% area wash, crosshair + tooltip, end-point label.  o = {labels, values, name, fmt, height, color} */
  function line(host, o){
    let geo = null;
    mount(host, W => {
      const n = o.labels.length, H = o.height || 280, col = o.color || "var(--s1)";
      const sc = nice(Math.max(...o.values, 0) || 1, 5), axis = o.axis || compact;
      const L = Math.max(...sc.ticks.map(t => tw(axis(t)))) + 12, endLab = tw((o.fmt || qty)(o.values[n - 1] ?? 0)) + 12, R = Math.max(10, endLab), T = 14, B = 24;
      const pw = W - L - R, ph = H - T - B;
      const x = i => L + (n > 1 ? i * pw / (n - 1) : pw / 2), y = v => T + ph - (v / sc.hi) * ph;
      const pts = o.values.map((v, i) => [x(i), y(v || 0)]);
      let s = `<svg height="${H}" viewBox="0 0 ${W} ${H}" role="img"><g class="gl">` + sc.ticks.map(t => `<line x1="${L}" x2="${W - R}" y1="${y(t)}" y2="${y(t)}"/>`).join("") + `</g>`;
      s += `<g class="tick">` + sc.ticks.map(t => `<text x="${L - 8}" y="${y(t) + 4}" text-anchor="end">${esc(axis(t))}</text>`).join("") + `</g>`;
      const every = Math.max(1, Math.ceil((Math.max(...o.labels.map(l => tw(l))) + 14) / (pw / Math.max(n - 1, 1))));
      o.labels.forEach((l, i) => { if (i % every === 0) s += `<text x="${x(i)}" y="${T + ph + 17}" text-anchor="middle">${esc(l)}</text>`; });
      if (n){
        const dl = pts.map((p, i) => (i ? "L" : "M") + p[0].toFixed(1) + "," + p[1].toFixed(1)).join("");
        s += `<path d="${dl}L${pts[n - 1][0]},${T + ph}L${pts[0][0]},${T + ph}Z" style="fill:${col};opacity:.1"/>`;
        s += `<path d="${dl}" style="fill:none;stroke:${col};stroke-width:2;stroke-linejoin:round;stroke-linecap:round"/>`;
        const [ex, ey] = pts[n - 1];
        s += `<circle cx="${ex}" cy="${ey}" r="4" style="fill:${col};stroke:var(--card);stroke-width:2"/><text class="lbl tnum" x="${ex + 8}" y="${ey + 4}">${esc((o.fmt || qty)(o.values[n - 1]))}</text>`;
      }
      s += `<line class="base" x1="${L}" x2="${W - R}" y1="${T + ph}" y2="${T + ph}"/><g class="xh" style="display:none"><line y1="${T}" y2="${T + ph}" style="stroke:var(--axis);stroke-width:1"/><circle r="5" style="fill:${col};stroke:var(--card);stroke-width:2"/></g>`;
      s += `<rect class="hit" data-i="0" x="${L}" y="${T}" width="${pw}" height="${ph}"/></svg>`;
      host.innerHTML = s; geo = {x, pts, n, L, pw};
    });
    host.addEventListener("mousemove", e => {
      if (!geo || !geo.n) return; const svg = $("svg", host), r = svg.getBoundingClientRect();
      const px = e.clientX - r.left; if (px < geo.L - 4 || px > geo.L + geo.pw + 4) { Tip.hide(); $(".xh", host).style.display = "none"; return; }
      const i = Math.max(0, Math.min(geo.n - 1, Math.round(geo.n > 1 ? (px - geo.L) / geo.pw * (geo.n - 1) : 0)));
      const g = $(".xh", host); g.style.display = ""; $("line", g).setAttribute("x1", geo.pts[i][0]); $("line", g).setAttribute("x2", geo.pts[i][0]);
      $("circle", g).setAttribute("cx", geo.pts[i][0]); $("circle", g).setAttribute("cy", geo.pts[i][1]);
      Tip.show(e, {title: o.labels[i], rows: [{color: o.color || "var(--s1)", name: o.name || "Value", value: (o.fmt || qty)(o.values[i])}]});
    });
    host.addEventListener("mouseleave", () => { Tip.hide(); const g = $(".xh", host); if (g) g.style.display = "none"; });
  }
  return {columns, bars, diverging, line};
})();
