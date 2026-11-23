// ===== Pas alleen deze waarden aan =====
window.APP_CONFIG = {
  // Entra app-registratie → Overzicht → "Toepassings-id (client)"
  clientId: "cf80d352-e6f7-4fed-ac27-25b39cf0432f",
  // Entra app-registratie → Overzicht → "Map-id (tenant)"
  tenantId: "d94482f8-7979-4a8d-839c-d848fb70f46a",

  // SharePoint
  spHost: "totabitbv.sharepoint.com",
  spSitePath: "/sites/MerzarioAutomations",
  spListName: "Seaexport email routing log",

  // Link naar CargoWise shipment (leeg laten = geen knop). {SM} wordt vervangen.
  cargowiseUrl: "",

  // ===== Tab "Links" =====
  // Placeholders: {container} {booking} {vessel} {shipment}
  links: {
    // Altijd getoond per container
    container: [
      { label: "Track-Trace", url: "https://www.track-trace.com/container?number={container}" }
    ],
    // Extra carrier-link op basis van containerprefix (eerste 4 letters). Controleer de URL's.
    carrierByPrefix: [
      { carrier: "Maersk", prefixes: ["MAEU", "MSKU", "MRKU", "MRSU"], url: "https://www.maersk.com/tracking/{container}" },
      { carrier: "Hapag-Lloyd", prefixes: ["HLXU", "HLBU", "HLCU", "UACU"], url: "https://www.hapag-lloyd.com/en/online-business/track/track-by-container-solution.html?container={container}" },
      { carrier: "CMA CGM", prefixes: ["CMAU", "CGMU", "APZU", "ECMU"], url: "https://www.cma-cgm.com/ebusiness/tracking/search?SearchBy=Container&Reference={container}" }
    ],
    // Per carrier booking / B/L
    booking: [
      { label: "Track-Trace B/L", url: "https://www.track-trace.com/bol?number={booking}" }
    ],
    // Per vessel
    vessel: [
      { label: "MarineTraffic", url: "https://www.marinetraffic.com/en/ais/index/search/all/keyword:{vessel}" }
    ],
    // Per shipment (SM-nummer), bijv. later MyMerzario of CargoWise
    shipment: []
  },

  // ===== Tab "Acties" (nog niet actief) =====
  docTypes: [
    { code: "MAI", label: "MAI – E-mail" },
    { code: "MSC", label: "MSC – Overig" }
  ]
};
