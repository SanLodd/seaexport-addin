/* Seaexport AI – Outlook taakvenster (v2, tabs)
   Leest de AI-analyse uit de SharePoint-lijst "Seaexport email routing log"
   op basis van het internetMessageId van de geopende mail. */

const CFG = window.APP_CONFIG;
const SCOPES = ["Sites.Read.All"];
const TABS = ["routing", "analyse", "cw", "acties", "links"];
let pca = null;
let siteId = null;

// Huidige toestand van het paneel
const state = {
  tab: loadPref("tab", "analyse"),
  row: null,        // SharePoint-item (of null als niet gevonden)
  count: 0,         // aantal gevonden logregels
  facts: null,      // geparste ShipmentFactsJson
  subview: null,    // { title, html } als een subscherm open staat
  notFound: false
};

// ---------- Start ----------
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    show(`<div class="state">Deze invoegtoepassing werkt alleen in Outlook.</div>`);
    return;
  }
  document.getElementById("refresh").onclick = () => loadCurrentItem();
  document.querySelectorAll("#tabs button").forEach((b) => {
    b.onclick = () => { state.subview = null; setTab(b.getAttribute("data-tab")); };
  });
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => loadCurrentItem());
  } catch (e) { /* oudere clients */ }
  try {
    pca = await msal.createNestablePublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname
      },
      cache: { cacheLocation: "localStorage" }
    });
  } catch (e) {
    showError("MSAL kon niet starten", e);
    return;
  }
  loadCurrentItem();
});

// ---------- Token & Graph ----------
async function getToken() {
  const req = { scopes: SCOPES };
  try {
    const r = await pca.acquireTokenSilent(req);
    return r.accessToken;
  } catch (silentErr) {
    console.log("Stil token mislukt, popup proberen:", silentErr);
    const r = await pca.acquireTokenPopup(req);
    return r.accessToken;
  }
}

async function graph(path) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly"
    }
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Graph ${res.status}: ${txt.substring(0, 500)}`);
  }
  return res.json();
}

async function getSiteId() {
  if (siteId) return siteId;
  const s = await graph(`/sites/${CFG.spHost}:${CFG.spSitePath}?$select=id`);
  siteId = s.id;
  return siteId;
}

// ---------- Data laden ----------
async function loadCurrentItem() {
  state.row = null; state.facts = null; state.count = 0; state.subview = null; state.notFound = false;
  const item = Office.context.mailbox.item;
  if (!item) { hideTabs(); show(`<div class="state">Selecteer een e-mail.</div>`); return; }
  const mid = item.internetMessageId;
  if (!mid) { hideTabs(); show(`<div class="state">Geen internetMessageId beschikbaar voor dit item (bijvoorbeeld een concept).</div>`); return; }

  hideTabs();
  show(`<div class="state">AI-analyse ophalen…</div>`);
  try {
    const sid = await getSiteId();
    const list = encodeURIComponent(CFG.spListName);
    const filter = encodeURIComponent(`fields/InternetMessageId eq '${mid.replace(/'/g, "''")}'`);
    const data = await graph(`/sites/${sid}/lists/${list}/items?$expand=fields&$filter=${filter}&$top=10`);
    const rows = (data.value || []).sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
    state.count = rows.length;
    if (rows.length) {
      state.row = rows[0];
      const fj = txt(get("ShipmentFactsJson"));
      if (fj) { try { state.facts = JSON.parse(fj); } catch (e) { state.facts = null; } }
    } else {
      state.notFound = true;
    }
    showTabs();
    setTab(state.tab);
  } catch (e) {
    hideTabs();
    showError("Ophalen mislukt", e);
  }
}

// ---------- Veldhulpen ----------
function F() { return (state.row && state.row.fields) || {}; }
function get(name) {
  const f = F();
  if (f[name] !== undefined) return f[name];
  const k = Object.keys(f).find((x) => x.toLowerCase() === name.toLowerCase());
  return k ? f[k] : undefined;
}
function txt(v) { return v === undefined || v === null ? "" : typeof v === "object" ? (v.Url || v.url || JSON.stringify(v)) : String(v); }
function isTrue(v) { const s = String(v).toLowerCase(); return v === true || s === "true" || s === "ja" || s === "yes"; }
function list(col) { return txt(get(col)).split(",").map((s) => s.trim()).filter(Boolean); }

// ---------- Tabs ----------
function setTab(tab) {
  if (!TABS.includes(tab)) tab = "analyse";
  state.tab = tab; savePref("tab", tab);
  document.querySelectorAll("#tabs button").forEach((b) => b.classList.toggle("active", b.getAttribute("data-tab") === tab));
  render();
}
function showTabs() {
  const nav = document.getElementById("tabs");
  nav.classList.remove("hidden");
  // Rode stip op Acties als er actie vereist is
  const btn = nav.querySelector('[data-tab="acties"]');
  const old = btn.querySelector(".dot"); if (old) old.remove();
  if (state.row && isTrue(get("ActionRequired"))) btn.insertAdjacentHTML("beforeend", `<span class="dot"></span>`);
}
function hideTabs() { document.getElementById("tabs").classList.add("hidden"); }

function render() {
  if (state.subview) {
    show(`<div class="subhead"><button id="back">← Terug</button><b>${esc(state.subview.title)}</b></div>${state.subview.html}`);
    document.getElementById("back").onclick = () => { state.subview = null; render(); };
    bindCopy();
    return;
  }
  const subj = (Office.context.mailbox.item && Office.context.mailbox.item.subject) || txt(get("Title"));
  let html = subj ? `<div class="subject">${esc(subj)}</div>` : "";
  const notFoundMsg = `<div class="state">Geen AI-analyse gevonden voor deze mail.<br><br><span class="muted">Mogelijk is de mail binnengekomen voordat de flow het InternetMessageId opsloeg, of heeft de flow deze mail (nog) niet verwerkt.</span></div>`;

  switch (state.tab) {
    case "routing": html += state.notFound ? notFoundMsg : renderRouting(); break;
    case "analyse": html += state.notFound ? notFoundMsg : renderAnalyse(); break;
    case "cw":      html += state.notFound ? notFoundMsg : renderCW(); break;
    case "acties":  html += renderActies(); break;
    case "links":   html += renderLinks(); break;
  }
  show(html);
  bindCopy();
  bindActions();
}

// ---------- Tab: Routing ----------
function renderRouting() {
  const folder = txt(get("Folder"));
  const conf = txt(get("Confidence")).toLowerCase();
  const moved = isTrue(get("Moved"));
  const source = txt(get("FolderSource"));
  let h = `<section><div class="badges">`;
  if (folder) h += `<span class="badge b-folder">📁 ${esc(folder)}</span>`;
  if (conf) h += `<span class="badge b-${esc(conf)}">${esc(conf)}</span>`;
  if (source) h += `<span class="badge">bron: ${esc(source)}</span>`;
  h += `<span class="badge">${moved ? "verplaatst" : "niet verplaatst"}</span></div></section>`;

  const rows = [
    ["Map", folder],
    ["Klant", txt(get("Customer"))],
    ["Collega", txt(get("Colleague"))],
    ["Regel", txt(get("Rule"))],
    ["CW operator", txt(get("CWOperator"))],
    ["Reden", txt(get("Reason"))],
    ["Ontvangen", fmtDate(txt(get("Received")))]
  ].filter((r) => r[1]);
  h += `<section><h2>Routering</h2>${kv(rows)}</section>`;

  const runLink = txt(get("RunLink"));
  h += `<section><h2>Links</h2><div class="links">`;
  if (state.row.webUrl) h += a(state.row.webUrl, "Open in SharePoint");
  if (runLink) h += a(runLink, "Open flow-run");
  h += `</div></section>`;
  if (state.count > 1) h += `<p class="muted">Let op: ${state.count} logregels gevonden; de nieuwste wordt getoond.</p>`;
  h += `<section><details><summary>Alle velden (debug)</summary><table>${renderAllFields(F())}</table></details></section>`;
  return h;
}

// ---------- Tab: Analyse ----------
function renderAnalyse() {
  const intent = txt(get("Intent"));
  const actionReq = isTrue(get("ActionRequired"));
  let h = `<section><div class="badges">`;
  if (intent) h += `<span class="badge">${esc(intent.replace(/_/g, " "))}</span>`;
  h += actionReq ? `<span class="badge b-action">Actie vereist</span>` : `<span class="badge">Geen actie</span>`;
  h += `</div></section>`;

  const action = txt(get("Action"));
  if (action) h += `<section><h2>Voorgestelde actie</h2><div class="action ${actionReq ? "" : "none"}">${esc(action)}</div></section>`;
  const summary = txt(get("Summary"));
  if (summary) h += `<section><h2>Samenvatting</h2><div class="summary">${esc(summary)}</div></section>`;

  const refs = [["Shipment", "ShipmentNumbers"], ["Carrier / MBL", "CarrierRefs"], ["Container", "ContainerNumbers"], ["Klantref.", "CustomerRefs"], ["Overig", "OtherRefs"]];
  let r = "";
  for (const [label, col] of refs) {
    const vals = list(col);
    if (!vals.length) continue;
    r += `<tr><td class="k">${label}</td><td><div class="chips">${vals.map(chip).join("")}</div></td></tr>`;
  }
  if (r) h += `<section><h2>Referenties</h2><table>${r}</table></section>`;
  return h;
}

// ---------- Tab: CargoWise ----------
function renderCW() {
  const f = state.facts;
  if (!f || !f.shipment) {
    const facts = txt(get("ShipmentFacts"));
    if (facts) return `<section><h2>Shipment (CargoWise)</h2><div style="white-space:pre-wrap">${esc(facts)}</div></section>`;
    return `<div class="state">Geen CargoWise-gegevens voor deze mail.<br><br><span class="muted">Er is geen geldig SM-nummer gevonden, of de lookup leverde niets op.</span></div>`;
  }
  const date = (v) => (v ? String(v).substring(0, 10) : "");
  const rows = [
    ["Shipment", f.shipment], ["Consol", f.consol], ["Status", f.shipment_status], ["Geannuleerd", f.cancelled],
    ["Operator", [f.operator_name, f.operator_code].filter(Boolean).join(" · ")],
    ["Klant", f.customer], ["Consignee", f.consignee], ["Carrier booking", f.carrier_booking],
    ["Vessel / voyage", [f.vessel, f.voyage].filter(Boolean).join(" / ")],
    ["POL → POD", [f.pol, f.pod].filter(Boolean).join(" → ")],
    ["ETD", date(f.etd)], ["ETA", date(f.eta)], ["Containers", f.container_count], ["Goederen", f.goods],
    ["Colli / kg", [f.packs, f.weight_kg].filter((x) => x !== undefined && x !== null && x !== "").join(" / ")],
    ["Release type", f.release_type], ["BL-status", f.bl_status], ["VGM", f.vgm], ["SHI", f.shi]
  ].filter((r) => r[1] !== undefined && r[1] !== null && String(r[1]) !== "");

  let h = `<section><h2>Shipment</h2>${kv(rows)}</section>`;
  const cnts = containerList();
  if (cnts.length) h += `<section><h2>Containers</h2>${cnts.map((c) => `<div class="cnt">${esc(c.display)}</div>`).join("")}</section>`;
  const recv = fmtDate(txt(get("Received")));
  h += `<p class="muted">Momentopname bij binnenkomst van de mail${recv ? " (" + esc(recv) + ")" : ""}.</p>`;
  if (CFG.cargowiseUrl) h += `<div class="links">${a(CFG.cargowiseUrl.replace("{SM}", f.shipment), "Open in CargoWise")}</div>`;
  return h;
}

// ---------- Tab: Acties ----------
function renderActies() {
  const actionReq = state.row && isTrue(get("ActionRequired"));
  const action = state.row ? txt(get("Action")) : "";
  let h = "";
  if (action) h += `<section><h2>Voorgestelde actie</h2><div class="action ${actionReq ? "" : "none"}">${esc(action)}</div></section>`;
  h += `<section><h2>Acties</h2>
    <button class="actionbtn" data-action="edocs"><span class="ic">📎</span><span>Opslaan in eDocs<small>Hele mail of losse bijlagen naar CargoWise</small></span></button>
    <button class="actionbtn" disabled><span class="ic">✉️</span><span>Conceptantwoord (AI)<small>Binnenkort</small></span></button>
    <button class="actionbtn" disabled><span class="ic">🕒</span><span>Closings bijwerken<small>Binnenkort</small></span></button>
  </section>`;
  return h;
}

function openEdocs() {
  state.subview = {
    title: "Opslaan in eDocs",
    html: `<div class="state">Deze actie wordt in de volgende stap gebouwd (backend via Power Automate).</div>`
  };
  render();
}

// ---------- Tab: Links ----------
function renderLinks() {
  const L = CFG.links || {};
  let h = "";

  // Containers
  const cnts = containerList();
  if (cnts.length) {
    h += `<section><h2>Containers</h2>`;
    for (const c of cnts) {
      const links = (L.container || []).map((t) => a(fill(t.url, { container: c.number }), t.label));
      const carrier = (L.carrierByPrefix || []).find((x) => x.prefixes.includes(c.number.substring(0, 4)));
      if (carrier) links.unshift(a(fill(carrier.url, { container: c.number }), carrier.carrier));
      h += `<div class="linkrow"><div class="ref">${chip(c.number)}</div><div class="links">${links.join("")}</div></div>`;
    }
    h += `</section>`;
  }

  // Carrier booking / B/L
  const bookings = uniq([...(state.facts && state.facts.carrier_booking ? [state.facts.carrier_booking] : []), ...list("CarrierRefs")]);
  if (bookings.length && (L.booking || []).length) {
    h += `<section><h2>Booking / B/L</h2>`;
    for (const b of bookings) {
      h += `<div class="linkrow"><div class="ref">${chip(b)}</div><div class="links">${L.booking.map((t) => a(fill(t.url, { booking: b }), t.label)).join("")}</div></div>`;
    }
    h += `</section>`;
  }

  // Vessel
  const vessel = state.facts && state.facts.vessel;
  if (vessel && (L.vessel || []).length) {
    h += `<section><h2>Vessel</h2><div class="linkrow"><div class="ref">${esc(vessel)}</div><div class="links">${L.vessel.map((t) => a(fill(t.url, { vessel: vessel }), t.label)).join("")}</div></div></section>`;
  }

  // Shipment
  const sms = uniq([...(state.facts && state.facts.shipment ? [state.facts.shipment] : []), ...list("ShipmentNumbers")]);
  if (sms.length && (L.shipment || []).length) {
    h += `<section><h2>Shipment</h2>`;
    for (const s of sms) h += `<div class="linkrow"><div class="ref">${chip(s)}</div><div class="links">${L.shipment.map((t) => a(fill(t.url, { shipment: s }), t.label)).join("")}</div></div>`;
    h += `</section>`;
  }

  if (!h) h = `<div class="state">Geen referenties gevonden om links voor te maken.</div>`;
  return h;
}

// Containers uit CargoWise-feiten + AI-referenties, ontdubbeld
function containerList() {
  const out = new Map();
  const re = /[A-Z]{4}\d{7}/;
  const f = state.facts;
  if (f && Array.isArray(f.containers)) {
    for (const c of f.containers) {
      const display = typeof c === "object" ? Object.values(c).filter(Boolean).join(" · ") : String(c);
      const m = display.toUpperCase().match(re);
      if (m && !out.has(m[0])) out.set(m[0], { number: m[0], display });
    }
  }
  for (const v of list("ContainerNumbers")) {
    const m = v.toUpperCase().replace(/\s/g, "").match(re);
    if (m && !out.has(m[0])) out.set(m[0], { number: m[0], display: m[0] });
  }
  return [...out.values()];
}

// ---------- Hulpfuncties ----------
const SYSTEM_FIELDS = new Set(["@odata.etag","ContentType","Modified","Created","AuthorLookupId","EditorLookupId","_UIVersionString","Attachments","Edit","LinkTitleNoMenu","LinkTitle","ItemChildCount","FolderChildCount","_ComplianceFlags","_ComplianceTag","_ComplianceTagWrittenTime","_ComplianceTagUserId","AppAuthorLookupId","AppEditorLookupId","id","_ColorTag","ContentTypeId","DocIcon"]);

function renderAllFields(f) {
  return Object.keys(f).filter((k) => !SYSTEM_FIELDS.has(k)).sort().map((k) => {
    const v = f[k];
    const s = typeof v === "object" ? JSON.stringify(v) : String(v);
    return `<tr><td class="k">${esc(k)}</td><td style="word-break:break-word">${esc(s.length > 400 ? s.substring(0, 400) + "…" : s)}</td></tr>`;
  }).join("");
}

function kv(rows) { return `<table>${rows.map((r) => `<tr><td class="k">${r[0]}</td><td>${esc(String(r[1]))}</td></tr>`).join("")}</table>`; }
function chip(s) { return `<span class="chip" data-copy="${esc(s)}" title="Klik om te kopiëren">${esc(s)}</span>`; }
function a(url, label) { return `<a href="${esc(url)}" data-url="${esc(url)}" target="_blank" rel="noopener">${esc(label)}</a>`; }
function fill(tpl, vals) { return tpl.replace(/\{(\w+)\}/g, (m, k) => encodeURIComponent(vals[k] || "")); }
function uniq(arr) { return [...new Set(arr.map((s) => String(s).trim()).filter(Boolean))]; }

function bindCopy() {
  document.querySelectorAll("[data-copy]").forEach((el) => { el.onclick = () => copy(el.getAttribute("data-copy")); });
  // Links openen in de standaardbrowser (werkt ook in klassieke Outlook)
  document.querySelectorAll("a[data-url]").forEach((el) => {
    el.onclick = (ev) => {
      try {
        if (Office.context.ui && Office.context.ui.openBrowserWindow) {
          ev.preventDefault();
          Office.context.ui.openBrowserWindow(el.getAttribute("data-url"));
        }
      } catch (e) { /* val terug op gewone link */ }
    };
  });
}

function bindActions() {
  document.querySelectorAll("[data-action]").forEach((el) => {
    el.onclick = () => { if (el.getAttribute("data-action") === "edocs") openEdocs(); };
  });
}

function show(html) { document.getElementById("content").innerHTML = html; }

function showError(title, e) {
  console.error(title, e);
  show(`<div class="state error"><b>${esc(title)}</b>\n${esc(e && e.message ? e.message : String(e))}</div>`);
}

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function fmtDate(v) {
  if (!v) return "";
  const d = new Date(v);
  return isNaN(d) ? v : d.toLocaleString("nl-NL", { dateStyle: "short", timeStyle: "short" });
}

function copy(text) {
  const ta = document.createElement("textarea");
  ta.value = text; document.body.appendChild(ta); ta.select();
  try { document.execCommand("copy"); } catch (e) {}
  document.body.removeChild(ta);
  const t = document.getElementById("toast");
  t.textContent = `Gekopieerd: ${text}`; t.style.display = "block";
  setTimeout(() => (t.style.display = "none"), 1500);
}

function loadPref(k, d) { try { return localStorage.getItem("sxai_" + k) || d; } catch (e) { return d; } }
function savePref(k, v) { try { localStorage.setItem("sxai_" + k, v); } catch (e) {} }
