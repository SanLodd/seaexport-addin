/* Seaexport AI – Outlook taakvenster
   Leest de AI-analyse uit de SharePoint-lijst "Seaexport email routing log"
   op basis van het internetMessageId van de geopende mail. */

const CFG = window.APP_CONFIG;
const SCOPES = ["Sites.Read.All"];
let pca = null;
let siteId = null;

// ---------- Start ----------
Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    show(`<div class="state">Deze invoegtoepassing werkt alleen in Outlook.</div>`);
    return;
  }
  document.getElementById("refresh").onclick = () => loadCurrentItem();
  // Bij vastgezet taakvenster: herladen als de gebruiker een andere mail kiest
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

// ---------- Token ----------
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

// ---------- Hoofdlogica ----------
async function loadCurrentItem() {
  const item = Office.context.mailbox.item;
  if (!item) {
    show(`<div class="state">Selecteer een e-mail.</div>`);
    return;
  }
  const mid = item.internetMessageId;
  if (!mid) {
    show(`<div class="state">Geen internetMessageId beschikbaar voor dit item (bijvoorbeeld een concept).</div>`);
    return;
  }
  show(`<div class="state">AI-analyse ophalen…</div>`);

  try {
    const sid = await getSiteId();
    const list = encodeURIComponent(CFG.spListName);
    const filterValue = mid.replace(/'/g, "''");
    const filter = encodeURIComponent(`fields/InternetMessageId eq '${filterValue}'`);
    const data = await graph(`/sites/${sid}/lists/${list}/items?$expand=fields&$filter=${filter}&$top=10`);
    const rows = (data.value || []).sort((a, b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));

    if (rows.length === 0) {
      show(`<div class="state">Geen AI-analyse gevonden voor deze mail.<br><br>
        <span class="muted">Mogelijk is de mail binnengekomen voordat de flow het InternetMessageId opsloeg, of heeft de flow deze mail (nog) niet verwerkt.</span></div>`);
      return;
    }
    render(rows[0], rows.length);
  } catch (e) {
    showError("Ophalen mislukt", e);
  }
}

// ---------- Weergave ----------
function render(row, count) {
  const F = row.fields || {};
  const get = (name) => {
    if (F[name] !== undefined) return F[name];
    const k = Object.keys(F).find((x) => x.toLowerCase() === name.toLowerCase());
    return k ? F[k] : undefined;
  };
  const txt = (v) => (v === undefined || v === null ? "" : typeof v === "object" ? (v.Url || v.url || JSON.stringify(v)) : String(v));

  const folder = txt(get("Folder"));
  const conf = txt(get("Confidence")).toLowerCase();
  const intent = txt(get("Intent"));
  const actionReq = get("ActionRequired");
  const actionReqBool = actionReq === true || String(actionReq).toLowerCase() === "true" || String(actionReq).toLowerCase() === "ja";
  const moved = get("Moved");
  const movedBool = moved === true || String(moved).toLowerCase() === "true";
  const source = txt(get("FolderSource"));

  let html = "";

  // Badges
  html += `<div class="badges">`;
  if (folder) html += `<span class="badge b-folder">📁 ${esc(folder)}</span>`;
  if (conf) html += `<span class="badge b-${esc(conf)}">${esc(conf)}</span>`;
  if (intent) html += `<span class="badge">${esc(intent.replace(/_/g, " "))}</span>`;
  html += actionReqBool ? `<span class="badge b-action">Actie vereist</span>` : `<span class="badge b-noaction">Geen actie</span>`;
  if (source) html += `<span class="badge">bron: ${esc(source)}</span>`;
  html += `<span class="badge">${movedBool ? "verplaatst" : "niet verplaatst"}</span>`;
  html += `</div>`;

  // Actie
  const action = txt(get("Action"));
  if (action) html += `<section><h2>Actie</h2><div class="action">${esc(action)}</div></section>`;

  // Samenvatting
  const summary = txt(get("Summary"));
  if (summary) html += `<section><h2>Samenvatting</h2><div class="summary">${esc(summary)}</div></section>`;

  // Referenties
  const refs = [
    ["Shipment", "ShipmentNumbers"],
    ["Carrier / MBL", "CarrierRefs"],
    ["Container", "ContainerNumbers"],
    ["Klantref.", "CustomerRefs"],
    ["Overig", "OtherRefs"]
  ];
  let refHtml = "";
  for (const [label, col] of refs) {
    const v = txt(get(col));
    if (!v) continue;
    const chips = v.split(",").map((s) => s.trim()).filter(Boolean)
      .map((s) => `<span class="chip" data-copy="${esc(s)}" title="Klik om te kopiëren">${esc(s)}</span>`).join("");
    refHtml += `<tr><td class="k">${label}</td><td><div class="chips">${chips}</div></td></tr>`;
  }
  if (refHtml) html += `<section><h2>Referenties</h2><table>${refHtml}</table></section>`;

  // Shipmentgegevens (CargoWise)
  const factsJson = txt(get("ShipmentFactsJson"));
  if (factsJson) {
    try {
      html += renderFacts(JSON.parse(factsJson));
    } catch (e) {
      const facts = txt(get("ShipmentFacts"));
      if (facts) html += `<section><h2>Shipment (CargoWise)</h2><div style="white-space:pre-wrap">${esc(facts)}</div></section>`;
    }
  }

  // Routering
  const routeRows = [
    ["Klant", txt(get("Customer"))],
    ["Collega", txt(get("Colleague"))],
    ["Regel", txt(get("Rule"))],
    ["CW operator", txt(get("CWOperator"))],
    ["Reden", txt(get("Reason"))],
    ["Ontvangen", fmtDate(txt(get("Received")))]
  ].filter((r) => r[1]);
  if (routeRows.length) {
    html += `<section><h2>Routering</h2><table>${routeRows.map((r) => `<tr><td class="k">${r[0]}</td><td>${esc(r[1])}</td></tr>`).join("")}</table></section>`;
  }

  // Links
  const runLink = txt(get("RunLink"));
  html += `<section><h2>Links</h2><div class="links">`;
  if (row.webUrl) html += `<a href="${esc(row.webUrl)}" target="_blank" rel="noopener">Open in SharePoint</a>`;
  if (runLink) html += `<a href="${esc(runLink)}" target="_blank" rel="noopener">Open flow-run</a>`;
  html += `</div></section>`;

  if (count > 1) html += `<p class="muted">Let op: ${count} logregels gevonden voor deze mail; de nieuwste wordt getoond.</p>`;

  // Debug: alle velden
  html += `<section><details><summary>Alle velden (debug)</summary><table>${renderAllFields(F)}</table></details></section>`;

  show(html);

  // Kopiëren
  document.querySelectorAll("[data-copy]").forEach((el) => {
    el.onclick = () => copy(el.getAttribute("data-copy"));
  });
}

function renderFacts(f) {
  if (!f || !f.shipment) return "";
  const date = (v) => (v ? String(v).substring(0, 10) : "");
  const rows = [
    ["Shipment", f.shipment],
    ["Consol", f.consol],
    ["Status", f.shipment_status],
    ["Geannuleerd", f.cancelled],
    ["Operator", [f.operator_name, f.operator_code].filter(Boolean).join(" · ")],
    ["Klant", f.customer],
    ["Consignee", f.consignee],
    ["Carrier booking", f.carrier_booking],
    ["Vessel / voyage", [f.vessel, f.voyage].filter(Boolean).join(" / ")],
    ["POL → POD", [f.pol, f.pod].filter(Boolean).join(" → ")],
    ["ETD", date(f.etd)],
    ["ETA", date(f.eta)],
    ["Containers", f.container_count],
    ["Goederen", f.goods],
    ["Colli / kg", [f.packs, f.weight_kg].filter((x) => x !== undefined && x !== null && x !== "").join(" / ")],
    ["Release type", f.release_type],
    ["BL-status", f.bl_status],
    ["VGM", f.vgm],
    ["SHI", f.shi]
  ].filter((r) => r[1] !== undefined && r[1] !== null && String(r[1]) !== "");

  let html = `<section><h2>Shipment (CargoWise)</h2><table>`;
  html += rows.map((r) => `<tr><td class="k">${r[0]}</td><td>${esc(String(r[1]))}</td></tr>`).join("");
  html += `</table>`;

  if (Array.isArray(f.containers) && f.containers.length) {
    const items = f.containers.map((c) => typeof c === "object" ? Object.values(c).filter(Boolean).join(" · ") : String(c));
    html += `<div style="margin-top:6px">${items.map((i) => `<div class="cnt">${esc(i)}</div>`).join("")}</div>`;
  }
  if (CFG.cargowiseUrl) {
    html += `<div class="links" style="margin-top:6px"><a href="${esc(CFG.cargowiseUrl.replace("{SM}", f.shipment))}" target="_blank" rel="noopener">Open in CargoWise</a></div>`;
  }
  return html + `</section>`;
}

const SYSTEM_FIELDS = new Set(["@odata.etag","ContentType","Modified","Created","AuthorLookupId","EditorLookupId","_UIVersionString","Attachments","Edit","LinkTitleNoMenu","LinkTitle","ItemChildCount","FolderChildCount","_ComplianceFlags","_ComplianceTag","_ComplianceTagWrittenTime","_ComplianceTagUserId","AppAuthorLookupId","AppEditorLookupId","id","_ColorTag","ContentTypeId","DocIcon"]);

function renderAllFields(F) {
  return Object.keys(F).filter((k) => !SYSTEM_FIELDS.has(k)).sort()
    .map((k) => {
      const v = F[k];
      const s = typeof v === "object" ? JSON.stringify(v) : String(v);
      return `<tr><td class="k">${esc(k)}</td><td style="word-break:break-word">${esc(s.length > 400 ? s.substring(0, 400) + "…" : s)}</td></tr>`;
    }).join("");
}

// ---------- Hulpfuncties ----------
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
  // Clipboard API is in Outlook-iframes vaak geblokkeerd → fallback via textarea
  const ta = document.createElement("textarea");
  ta.value = text; document.body.appendChild(ta); ta.select();
  try { document.execCommand("copy"); } catch (e) {}
  document.body.removeChild(ta);
  const t = document.getElementById("toast");
  t.textContent = `Gekopieerd: ${text}`; t.style.display = "block";
  setTimeout(() => (t.style.display = "none"), 1500);
}
