type Env = {
  DB: D1Database;
  EVIDENCIAS?: R2Bucket;
  MGA_API_KEY: string;
  SERVICE_NAME?: string;
};

type JsonObject = Record<string, unknown>;
type D1Row = Record<string, unknown>;

class HttpError extends Error {
  readonly status: number;

  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

const JSON_HEADERS = {
  "content-type": "application/json; charset=utf-8",
  "cache-control": "no-store, max-age=0",
  "access-control-allow-origin": "*",
  "access-control-allow-methods": "GET,POST,OPTIONS",
  "access-control-allow-headers": "Content-Type,X-MGA-API-Key"
};

const HTML_HEADERS = {
  "content-type": "text/html; charset=utf-8",
  "cache-control": "no-store, max-age=0",
  "access-control-allow-origin": "*"
};

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: JSON_HEADERS });
    }

    const url = new URL(request.url);
    try {
      if ((url.pathname === "/" || url.pathname === "/web" || url.pathname === "/programa" || url.pathname === "/almacen-filtros") && request.method === "GET") return webHome(env);
      if (url.pathname === "/favicon.ico" && request.method === "GET") return favicon();

      if (url.pathname === "/health") {
        return json({
          ok: true,
          service: env.SERVICE_NAME || "mga-cloud-sync-worker",
          storage: env.EVIDENCIAS ? "cloudflare-d1-r2" : "cloudflare-d1",
          generated_at: nowIso()
        });
      }

      await requireApiKey(request, env);

      if (url.pathname === "/api/stats" && request.method === "GET") return stats(env);
      if (url.pathname === "/api/catalog" && request.method === "GET") return getCatalog(env);
      if (url.pathname === "/api/catalog" && request.method === "POST") return saveSnapshot(request, env.DB, "catalog_snapshot", "catalog");
      if (url.pathname === "/api/portal" && request.method === "GET") return getSnapshot(env.DB, "portal_snapshot", emptyPortalSnapshot());
      if (url.pathname === "/api/portal/snapshot" && request.method === "POST") return saveSnapshot(request, env.DB, "portal_snapshot", "default");
      if (url.pathname === "/api/portal/snapshot" && request.method === "GET") return getSnapshot(env.DB, "portal_snapshot", emptyPortalSnapshot());
      if (url.pathname === "/api/filter-inventory/snapshot" && request.method === "GET") return getSnapshot(env.DB, "filter_inventory_snapshot", emptyFilterSnapshot());
      if (url.pathname === "/api/filter-inventory/snapshot" && request.method === "POST") return saveFilterSnapshot(request, env.DB);
      if (url.pathname === "/api/filter-inventory" && request.method === "GET") return getSnapshot(env.DB, "filter_inventory_snapshot", emptyFilterSnapshot());
      if (url.pathname === "/api/epp/snapshot" && request.method === "GET") return getSnapshot(env.DB, "epp_snapshot", emptyEppSnapshot());
      if (url.pathname === "/api/epp/snapshot" && request.method === "POST") return saveSnapshot(request, env.DB, "epp_snapshot", "default");
      if (url.pathname === "/api/epp" && request.method === "GET") return getSnapshot(env.DB, "epp_snapshot", emptyEppSnapshot());
      if (url.pathname === "/api/diesel" && request.method === "GET") return getDieselSnapshot(env.DB);
      if (url.pathname === "/api/sync" && request.method === "POST") return syncMobile(request, env);
      if (url.pathname === "/api/desktop/pending" && request.method === "GET") return desktopPending(url, env);
      if (url.pathname === "/api/desktop/ack" && request.method === "POST") return desktopAck(request, env);
      if (url.pathname === "/api/mobile/status" && request.method === "POST") return mobileStatus(request, env);
      if (url.pathname === "/api/hose-changes" && request.method === "GET") return getHoseChanges(url, env.DB);
      if (url.pathname === "/api/hose-changes" && request.method === "POST") return saveHoseChange(request, env.DB);
      if (url.pathname === "/api/hose-changes/delete" && request.method === "POST") return deleteHoseChange(request, env.DB);

      return json({ ok: false, detail: "Ruta no encontrada." }, 404);
    } catch (error) {
      const status = error instanceof HttpError ? error.status : 500;
      const detail = error instanceof Error ? error.message : "Error interno.";
      console.error(JSON.stringify({ status, path: url.pathname, detail }));
      return json({ ok: false, detail }, status);
    }
  }
};

async function requireApiKey(request: Request, env: Env): Promise<void> {
  const expected = String(env.MGA_API_KEY || "");
  if (!expected) throw new HttpError(500, "Falta configurar MGA_API_KEY en Cloudflare.");
  const received = request.headers.get("X-MGA-API-Key") || "";
  if (!(await constantTimeEqual(received, expected))) {
    throw new HttpError(401, "Unauthorized");
  }
}

async function constantTimeEqual(a: string, b: string): Promise<boolean> {
  const encoder = new TextEncoder();
  const [left, right] = await Promise.all([
    crypto.subtle.digest("SHA-256", encoder.encode(a)),
    crypto.subtle.digest("SHA-256", encoder.encode(b))
  ]);
  const leftBytes = new Uint8Array(left);
  const rightBytes = new Uint8Array(right);
  let diff = leftBytes.length ^ rightBytes.length;
  for (let index = 0; index < leftBytes.length && index < rightBytes.length; index += 1) {
    diff |= leftBytes[index] ^ rightBytes[index];
  }
  return diff === 0 && a.length === b.length;
}

function json(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), { status, headers: JSON_HEADERS });
}

function html(payload: string, status = 200): Response {
  return new Response(payload, { status, headers: HTML_HEADERS });
}

function favicon(): Response {
  return new Response(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="10" fill="#071f49"/><text x="32" y="39" text-anchor="middle" font-family="Arial,Helvetica,sans-serif" font-size="20" font-weight="700" fill="#fff">MGA</text><rect x="7" y="47" width="50" height="5" fill="#078f8c"/></svg>`,
    { headers: { "content-type": "image/svg+xml", "cache-control": "public, max-age=86400" } }
  );
}

async function webHome(env: Env): Promise<Response> {
  const captures = await captureCounts(env.DB);
  const catalogSnapshot = await loadSnapshot(env.DB, "catalog_snapshot", "catalog", emptyCatalog());
  const portalSnapshot = await loadSnapshot(env.DB, "portal_snapshot", "default", emptyPortalSnapshot());
  const filterSnapshot = await loadSnapshot(env.DB, "filter_inventory_snapshot", "default", emptyFilterSnapshot());
  const eppSnapshot = await loadSnapshot(env.DB, "epp_snapshot", "default", emptyEppSnapshot());
  const mobileRows = await recentMobileCaptures(env.DB, 90);
  const hoseRows = await getHoseRowsForDates(env.DB, monthStartIso(), todayIso());
  const hose = hoseReport(hoseRows, monthStartIso(), todayIso());

  const catalogEquipment = rowsFrom(catalogSnapshot.payload.equipment);
  const portalEquipment = rowsFrom(portalSnapshot.payload.equipment);
  const equipmentRows = portalEquipment.length ? portalEquipment : catalogEquipment;
  const portalCaptures = rowsFrom(portalSnapshot.payload.captures);
  const filterRows = rowsFrom(filterSnapshot.payload.inventory);
  const portalEpp = objectFrom(portalSnapshot.payload.epp);
  const eppPayload = rowsFrom(eppSnapshot.payload.items).length ? eppSnapshot.payload : portalEpp;
  const eppRows = rowsFrom(eppPayload.items);
  const eppSummary = objectFrom(eppPayload.summary);
  const tirePayload = objectFrom(portalSnapshot.payload.tire_kpi);
  const tireRows = rowsFrom(tirePayload.rows);
  const oilPayload = objectFrom(portalSnapshot.payload.oil_kpi);
  const oilRows = rowsFrom(oilPayload.rows);
  const kpiReports = objectFrom(portalSnapshot.payload.kpi_reports);
  const kpiSummary = firstKpiSummary(kpiReports);
  const filterEquipmentRows = catalogEquipment.map((item) => {
    const filters = rowsFrom(item.filters);
    const shortageCount = filters.filter((row) => numberValue(row.shortage) > 0 || /FALTANTE|SIN INVENTARIO/i.test(stringValue(row.inventory_status))).length;
    return {
      code: item.code,
      description: item.description,
      family: item.family,
      filters: filters.length,
      shortage_count: shortageCount,
      status: shortageCount > 0 ? "FALTANTES" : "OK"
    };
  });

  const storage = env.EVIDENCIAS ? "D1 + R2" : "D1";
  const lastUpdate = latestText([
    catalogSnapshot.updatedAt,
    portalSnapshot.updatedAt,
    filterSnapshot.updatedAt,
    eppSnapshot.updatedAt,
    stringValue(portalSnapshot.payload.updated_at),
    stringValue(portalSnapshot.payload.generated_at)
  ]);
  return html(`<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MGA Mantenimiento Cloud</title>
  <style>
    :root { color-scheme: light; --blue:#071f49; --teal:#078f8c; --red:#e11d48; --amber:#b45309; --green:#047857; --bg:#f4f6f8; --card:#fff; --line:#d7dee8; --text:#102033; --muted:#5d6a7c; --soft:#eef4fb; }
    * { box-sizing:border-box; }
    body { margin:0; font-family:Arial, Helvetica, sans-serif; background:var(--bg); color:var(--text); }
    header { background:var(--blue); color:white; padding:18px 24px 14px; }
    header .bar { max-width:1260px; margin:0 auto; display:flex; align-items:center; justify-content:space-between; gap:16px; }
    header h1 { margin:0; font-size:24px; letter-spacing:0; }
    header p { margin:5px 0 0; color:#dbe7ff; font-size:14px; }
    main { max-width:1260px; margin:0 auto; padding:18px; }
    .status { display:inline-block; padding:7px 10px; border-radius:999px; background:#d9f7e8; color:#076b43; font-weight:700; font-size:13px; white-space:nowrap; }
    .notice { background:#fff7ed; border:1px solid #fed7aa; border-radius:8px; color:#7c2d12; padding:10px 12px; margin-bottom:14px; font-size:13px; }
    .grid { display:grid; grid-template-columns:repeat(6, minmax(0,1fr)); gap:10px; margin-bottom:14px; }
    .card, .panel { background:var(--card); border:1px solid var(--line); border-radius:8px; }
    .card { padding:13px; min-height:86px; }
    .card span { display:block; color:var(--muted); font-size:12px; margin-bottom:7px; }
    .card strong { font-size:26px; color:var(--blue); line-height:1; }
    .card small { display:block; margin-top:7px; color:var(--muted); font-size:11px; }
    .tabs { display:flex; gap:6px; flex-wrap:wrap; margin:8px 0 14px; }
    .tabs button { border:1px solid var(--line); border-radius:7px; background:white; color:var(--blue); padding:9px 12px; font-weight:700; cursor:pointer; }
    .tabs button.active { background:var(--teal); border-color:var(--teal); color:white; }
    .view { display:none; }
    .view.active { display:block; }
    .panel { margin-top:12px; overflow:hidden; }
    .panel-head { display:flex; justify-content:space-between; align-items:center; gap:12px; padding:12px 14px; border-bottom:1px solid var(--line); background:#fbfdff; }
    .panel h2 { margin:0; font-size:17px; }
    .panel .hint { color:var(--muted); font-size:12px; }
    .tools { display:flex; gap:8px; flex-wrap:wrap; align-items:center; }
    input.search { border:1px solid var(--line); border-radius:6px; padding:8px 10px; min-width:230px; }
    .table-wrap { overflow:auto; max-height:560px; }
    table { width:100%; border-collapse:collapse; font-size:12px; }
    th { position:sticky; top:0; background:var(--blue); color:white; z-index:1; }
    th, td { padding:8px 9px; border-bottom:1px solid #e5ebf2; text-align:left; vertical-align:top; white-space:nowrap; }
    tbody tr:nth-child(even) { background:#f8fbff; }
    td.num, th.num { text-align:right; }
    .tag { display:inline-block; border-radius:999px; padding:3px 8px; font-weight:700; font-size:11px; background:#e8f7f2; color:var(--green); }
    .tag.warn { background:#fff7d6; color:var(--amber); }
    .tag.bad { background:#ffe4e6; color:var(--red); }
    .empty { padding:22px; color:var(--muted); }
    code { background:#edf1f6; padding:2px 5px; border-radius:4px; }
    a { color:var(--teal); font-weight:700; }
    @media (max-width:1100px) { .grid { grid-template-columns:repeat(3, minmax(0,1fr)); } }
    @media (max-width:760px) { header .bar { align-items:flex-start; flex-direction:column; } .grid { grid-template-columns:repeat(2, minmax(0,1fr)); } main { padding:12px; } input.search { width:100%; min-width:0; } .panel-head { align-items:flex-start; flex-direction:column; } }
  </style>
</head>
<body>
  <header>
    <div class="bar">
      <div>
        <h1>MGA Mantenimiento Cloud</h1>
        <p>Portal web y nube de sincronizacion para APK y programa de escritorio</p>
      </div>
      <span class="status">En linea</span>
    </div>
  </header>
  <main>
    <div class="notice">Cloudflare levanta la pagina web, API y base en la nube. El programa de escritorio sigue siendo una aplicacion Windows local y se conecta a esta nube para publicar/importar datos.</div>
    <section class="grid">
      <div class="card"><span>Equipos publicados</span><strong>${equipmentRows.length}</strong><small>${escapeHtml(catalogSnapshot.updatedAt || "catalogo pendiente")}</small></div>
      <div class="card"><span>Capturas APK</span><strong>${numberValue(captures.total)}</strong><small>${portalCaptures.length} en bitacora PC</small></div>
      <div class="card"><span>Pendientes PC</span><strong>${numberValue(captures.pending)}</strong></div>
      <div class="card"><span>Filtros en inventario</span><strong>${filterRows.length}</strong><small>${countFilterAlerts(filterRows)} alerta(s)</small></div>
      <div class="card"><span>EPP / alertas</span><strong>${eppRows.length}</strong><small>${numberValue(eppSummary.low_stock) + numberValue(eppSummary.out_stock) + numberValue(eppSummary.due_soon) + numberValue(eppSummary.overdue)} alerta(s)</small></div>
      <div class="card"><span>Mangueras mes</span><strong>${numberValue(objectFrom(hose.totals).changes)}</strong><small>${formatNumber(objectFrom(hose.totals).real_qty)} piezas</small></div>
    </section>
    <nav class="tabs" aria-label="Modulos">
      <button class="active" data-tab="inicio">Inicio</button>
      <button data-tab="equipos">Equipos</button>
      <button data-tab="capturas">Capturas APK</button>
      <button data-tab="filtros">Filtros</button>
      <button data-tab="epp">EPP</button>
      <button data-tab="aceites">Aceites</button>
      <button data-tab="llantas">Llantas</button>
      <button data-tab="mangueras">Mangueras</button>
      <button data-tab="api">API</button>
    </nav>

    <section id="inicio" class="view active">
      <div class="panel">
        <div class="panel-head"><h2>Resumen operativo</h2><span class="hint">Ultima actualizacion: ${escapeHtml(lastUpdate || "pendiente")}</span></div>
        ${summaryTable(kpiSummary)}
      </div>
      <div class="panel">
        <div class="panel-head"><h2>Ultimas capturas moviles</h2><span class="hint">Recibidas por la nube</span></div>
        ${tableHtml(mobileRows.slice(0, 12), [
          { label: "Fecha", keys: ["work_date"] },
          { label: "Equipo", keys: ["equipment_code"] },
          { label: "Componente", keys: ["component_name"] },
          { label: "Mecanico", keys: ["user_name"] },
          { label: "Recibido", keys: ["received_at"] },
          { label: "Fotos", keys: ["photos"], numeric: true },
          { label: "Estado PC", keys: ["desktop_status"] }
        ], "Todavia no hay capturas moviles en D1.", 12)}
      </div>
    </section>

    <section id="equipos" class="view">
      ${panelTable("Equipos publicados", "Catalogo tomado del programa PC", "Buscar equipo", "equipmentSearch", "equipmentTable", equipmentRows, [
        { label: "Codigo", keys: ["code", "equipment_code"] },
        { label: "Descripcion", keys: ["description", "equipment_description"] },
        { label: "Familia", keys: ["family"] },
        { label: "Estatus", keys: ["status", "condition"] }
      ], "Publica equipos desde el boton Nube del programa de escritorio.")}
      ${panelTable("Filtros por equipo", "Relacion de filtros del catalogo movil", "Buscar filtro/equipo", "equipmentFilterSearch", "equipmentFilterTable", filterEquipmentRows, [
        { label: "Equipo", keys: ["code"] },
        { label: "Descripcion", keys: ["description"] },
        { label: "Familia", keys: ["family"] },
        { label: "Filtros", keys: ["filters"], numeric: true },
        { label: "Faltantes", keys: ["shortage_count"], numeric: true },
        { label: "Estado", keys: ["status"] }
      ], "El catalogo no trae filtros por equipo todavia.")}
    </section>

    <section id="capturas" class="view">
      ${panelTable("Capturas locales publicadas por PC", "Bitacora anual publicada en el snapshot del portal", "Buscar captura", "portalCaptureSearch", "portalCaptureTable", portalCaptures, [
        { label: "Fecha", keys: ["work_date"] },
        { label: "Turno", keys: ["shift"] },
        { label: "Equipo", keys: ["equipment_code"] },
        { label: "Componente", keys: ["component"] },
        { label: "HI", keys: ["hi"], numeric: true },
        { label: "HF", keys: ["hf"], numeric: true },
        { label: "Hrs Trab", keys: ["worked_hours"], numeric: true },
        { label: "MP", keys: ["mp_hours"], numeric: true },
        { label: "MC", keys: ["mc_hours"], numeric: true },
        { label: "Estado", keys: ["status"] }
      ], "Publica el portal desde el programa PC para ver la bitacora aqui.", 160)}
      ${panelTable("Capturas APK en nube", "Estas son las capturas recibidas para importar a la PC", "Buscar APK", "mobileSearch", "mobileTable", mobileRows, [
        { label: "ID", keys: ["mobile_id"] },
        { label: "Fecha", keys: ["work_date"] },
        { label: "Equipo", keys: ["equipment_code"] },
        { label: "Componente", keys: ["component_name"] },
        { label: "Mecanico", keys: ["user_name"] },
        { label: "Recibido", keys: ["received_at"] },
        { label: "Fotos", keys: ["photos"], numeric: true },
        { label: "Estado PC", keys: ["desktop_status"] }
      ], "Sin capturas APK alojadas.", 160)}
    </section>

    <section id="filtros" class="view">
      ${panelTable("Inventario general de filtros", "Existencias publicadas desde el modulo de filtros", "Buscar filtro", "filterSearch", "filterTable", filterRows, [
        { label: "Parte", keys: ["part_number", "part_key"] },
        { label: "Descripcion", keys: ["description"] },
        { label: "Existencia", keys: ["quantity"], numeric: true },
        { label: "Minimo", keys: ["min_stock"], numeric: true },
        { label: "Unidad", keys: ["unit"] },
        { label: "Ubicacion", keys: ["location"] },
        { label: "Archivo", keys: ["source_file"] }
      ], "Publica inventario de filtros desde el programa PC.", 220)}
    </section>

    <section id="epp" class="view">
      ${panelTable("Inventario EPP", "Stock, minimos y vida util publicada desde PC", "Buscar EPP", "eppSearch", "eppTable", eppRows, [
        { label: "Codigo", keys: ["code"] },
        { label: "Descripcion", keys: ["description"] },
        { label: "Categoria", keys: ["category"] },
        { label: "Talla", keys: ["size"] },
        { label: "Existencia", keys: ["quantity"], numeric: true },
        { label: "Minimo", keys: ["min_stock"], numeric: true },
        { label: "Estado", keys: ["status"] },
        { label: "Prox. reposicion", keys: ["next_due_date"] },
        { label: "Trabajador", keys: ["next_worker"] }
      ], "Sin snapshot EPP publicado.", 220)}
    </section>

    <section id="aceites" class="view">
      ${panelTable("Consumo de lubricantes", "Litros por maquina tomados del KPI de aceites", "Buscar aceite/equipo", "oilSearch", "oilTable", oilRows, [
        { label: "Equipo", keys: ["code"] },
        { label: "Descripcion", keys: ["description"] },
        { label: "Grupo", keys: ["group"] },
        { label: "Hrs Trab", keys: ["worked_hours"], numeric: true },
        { label: "Motor 15W40", keys: ["oil_motor_15w40"], numeric: true },
        { label: "HCO ISO68", keys: ["oil_hco_iso68"], numeric: true },
        { label: "SAE30", keys: ["oil_trans_sae30"], numeric: true },
        { label: "85W140", keys: ["oil_85w140"], numeric: true },
        { label: "Total L", keys: ["total_liters"], numeric: true }
      ], "Publica el reporte de aceites desde el programa PC.", 160)}
    </section>

    <section id="llantas" class="view">
      ${panelTable("Vida util de llantas", "Porcentaje por llanta y equipo para KPI mensual", "Buscar llanta/equipo", "tireSearch", "tireTable", tireRows, [
        { label: "Equipo", keys: ["equipment_code"] },
        { label: "Llanta", keys: ["tire_code"] },
        { label: "Pos.", keys: ["position"] },
        { label: "Marca", keys: ["brand"] },
        { label: "Medida", keys: ["size"] },
        { label: "% vida", keys: ["life_percent"], numeric: true },
        { label: "Hrs uso", keys: ["hours_used"], numeric: true },
        { label: "Hrs rest.", keys: ["life_remaining_hours"], numeric: true },
        { label: "Estatus", keys: ["control_status", "status"] }
      ], "Publica el portal desde PC para ver KPI de llantas.", 220)}
    </section>

    <section id="mangueras" class="view">
      ${panelTable("Cambios de mangueras y conexiones", "Periodo actual: ${escapeHtml(monthStartIso())} a ${escapeHtml(todayIso())}", "Buscar manguera/equipo", "hoseSearch", "hoseTable", rowsFrom(hose.records), [
        { label: "Fecha", keys: ["change_date"] },
        { label: "Equipo", keys: ["equipment"] },
        { label: "Sistema", keys: ["system"] },
        { label: "Tipo", keys: ["part_type"] },
        { label: "Diam.", keys: ["diameter"] },
        { label: "Largo m", keys: ["length_m"], numeric: true },
        { label: "Cant.", keys: ["quantity"], numeric: true },
        { label: "Causa", keys: ["failure_reason"] },
        { label: "Tecnico", keys: ["technician"] }
      ], "Sin cambios de mangueras este mes.", 220)}
      ${panelTable("Consumo resumido de mangueras", "Agrupado por equipo y tipo", "Buscar resumen", "hoseSummarySearch", "hoseSummaryTable", rowsFrom(hose.summary), [
        { label: "Equipo", keys: ["equipment"] },
        { label: "Tipo", keys: ["part_type"] },
        { label: "Sistemas", keys: ["systems"] },
        { label: "Cambios", keys: ["changes"], numeric: true },
        { label: "Real", keys: ["quantity"], numeric: true },
        { label: "Aprox.", keys: ["estimated_qty"], numeric: true },
        { label: "Var.", keys: ["variance"], numeric: true }
      ], "Sin resumen de mangueras.", 100)}
    </section>

    <section id="api" class="view">
      <div class="panel">
        <div class="panel-head"><h2>Conexion de app y PC</h2><span class="hint">Almacenamiento: ${escapeHtml(storage)}</span></div>
        <div class="empty">
          URL para configurar: <code>${escapeHtml(new URL("https://mga-cloud-sync.mga-vicozero.workers.dev").origin)}</code><br>
          Estado publico: <a href="/health">/health</a><br>
          Endpoints con clave: <code>/api/catalog</code>, <code>/api/portal</code>, <code>/api/sync</code>, <code>/api/desktop/pending</code>, <code>/api/hose-changes</code>.
        </div>
      </div>
    </section>
  </main>
  <script>
    const buttons = [...document.querySelectorAll("[data-tab]")];
    const views = [...document.querySelectorAll(".view")];
    function activate(tab){
      buttons.forEach(button => button.classList.toggle("active", button.dataset.tab === tab));
      views.forEach(view => view.classList.toggle("active", view.id === tab));
      history.replaceState(null, "", "#" + tab);
    }
    buttons.forEach(button => button.addEventListener("click", () => activate(button.dataset.tab)));
    if(location.hash){
      const tab = location.hash.slice(1);
      if(document.getElementById(tab)) activate(tab);
    }
    document.querySelectorAll("input.search").forEach(input => {
      input.addEventListener("input", () => {
        const table = document.getElementById(input.dataset.table);
        const q = input.value.trim().toUpperCase();
        if(!table) return;
        table.querySelectorAll("tbody tr").forEach(row => {
          row.style.display = !q || row.textContent.toUpperCase().includes(q) ? "" : "none";
        });
      });
    });
  </script>
</body>
</html>`);
}

type SnapshotResult = { payload: JsonObject; updatedAt: string };

type WebColumn = {
  label: string;
  keys: string[];
  numeric?: boolean;
};

async function loadSnapshot(db: D1Database, table: string, name: string, fallback: JsonObject): Promise<SnapshotResult> {
  const row = await db.prepare(`SELECT payload_json, updated_at FROM ${table} WHERE name = ?`).bind(name).first<D1Row>();
  if (!row) return { payload: fallback, updatedAt: "" };
  return { payload: parseJsonObject(row.payload_json, fallback), updatedAt: stringValue(row.updated_at) };
}

async function recentMobileCaptures(db: D1Database, limit: number): Promise<JsonObject[]> {
  const rows = (await db
    .prepare(
      "SELECT c.mobile_id, c.user_name, c.equipment_code, c.component_name, c.work_date, c.received_at, c.desktop_imported_at, COUNT(p.id) AS photos " +
        "FROM mobile_capture c LEFT JOIN mobile_photo p ON p.capture_id = c.id " +
        "GROUP BY c.id ORDER BY c.id DESC LIMIT ?"
    )
    .bind(limit)
    .all<D1Row>()).results || [];
  return rows.map((row) => ({
    ...row,
    photos: numberValue(row.photos),
    desktop_status: row.desktop_imported_at ? "Importada" : "Pendiente"
  }));
}

function rowsFrom(value: unknown): JsonObject[] {
  return Array.isArray(value) ? value.filter(isObject) : [];
}

function objectFrom(value: unknown): JsonObject {
  return isObject(value) ? value : {};
}

function latestText(values: string[]): string {
  return values.filter((item) => item && item.trim()).sort().pop() || "";
}

function firstValue(row: JsonObject, keys: string[]): unknown {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key) && row[key] !== null && row[key] !== undefined && stringValue(row[key]) !== "") {
      return row[key];
    }
  }
  return "";
}

function formatNumber(value: unknown): string {
  const number = numberValue(value);
  if (!Number.isFinite(number)) return "0";
  if (Math.abs(number) >= 100) return number.toFixed(0);
  if (Math.abs(number) >= 10) return number.toFixed(1);
  return number.toFixed(2).replace(/\.?0+$/, "");
}

function statusHtml(value: unknown): string {
  const text = stringValue(value);
  if (!text) return "";
  const upper = text.toUpperCase();
  const kind = /SIN|FALT|VENC|CRIT|NO DISP|BAJA|ALTO/.test(upper) ? "bad" : /BAJO|PROX|REV|PEND|URG|ALERTA/.test(upper) ? "warn" : "";
  return `<span class="tag ${kind}">${escapeHtml(text)}</span>`;
}

function tableHtml(rows: JsonObject[], columns: WebColumn[], emptyText: string, limit = 120): string {
  const visibleRows = rows.slice(0, limit);
  if (!visibleRows.length) return `<div class="empty">${escapeHtml(emptyText)}</div>`;
  const head = columns.map((column) => `<th class="${column.numeric ? "num" : ""}">${escapeHtml(column.label)}</th>`).join("");
  const body = visibleRows
    .map((row) => `<tr>${columns.map((column) => {
      const value = firstValue(row, column.keys);
      const htmlValue = column.numeric ? escapeHtml(formatNumber(value)) : /estado|estatus|status/i.test(column.label) ? statusHtml(value) : escapeHtml(value);
      return `<td class="${column.numeric ? "num" : ""}">${htmlValue}</td>`;
    }).join("")}</tr>`)
    .join("");
  const footer = rows.length > visibleRows.length ? `<div class="empty">Mostrando ${visibleRows.length} de ${rows.length}. Usa el buscador para filtrar lo visible.</div>` : "";
  return `<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>${footer}`;
}

function panelTable(title: string, hint: string, searchPlaceholder: string, inputId: string, tableId: string, rows: JsonObject[], columns: WebColumn[], emptyText: string, limit = 120): string {
  const table = tableHtml(rows, columns, emptyText, limit).replace("<table>", `<table id="${escapeHtml(tableId)}">`);
  return `<div class="panel"><div class="panel-head"><div><h2>${escapeHtml(title)}</h2><span class="hint">${escapeHtml(hint)}</span></div><div class="tools"><input id="${escapeHtml(inputId)}" class="search" data-table="${escapeHtml(tableId)}" placeholder="${escapeHtml(searchPlaceholder)}"></div></div>${table}</div>`;
}

function countFilterAlerts(rows: JsonObject[]): number {
  return rows.filter((row) => {
    const status = stringValue(row.status || row.inventory_status).toUpperCase();
    const quantity = numberValue(row.quantity);
    const minStock = numberValue(row.min_stock);
    return /SIN|BAJO|FALT/.test(status) || (minStock > 0 && quantity <= minStock);
  }).length;
}

function firstKpiSummary(kpiReports: JsonObject): JsonObject {
  for (const value of Object.values(kpiReports)) {
    const report = objectFrom(value);
    const totals = objectFrom(report.totals);
    if (Object.keys(totals).length) {
      return {
        group: report.group,
        availability: totals.availability,
        utilization: totals.utilization,
        tmef: totals.tmef,
        tmpr: totals.tmpr,
        reliability: totals.reliability,
        stops: totals.stops,
        worked_hours: totals.worked_hours
      };
    }
  }
  return {};
}

function summaryTable(summary: JsonObject): string {
  if (!Object.keys(summary).length) {
    return `<div class="empty">Publica el portal desde el programa PC para llenar los KPI generales.</div>`;
  }
  return tableHtml([summary], [
    { label: "Grupo", keys: ["group"] },
    { label: "% Disp", keys: ["availability"], numeric: true },
    { label: "% Util", keys: ["utilization"], numeric: true },
    { label: "TMEF", keys: ["tmef"], numeric: true },
    { label: "TMPR", keys: ["tmpr"], numeric: true },
    { label: "Confiabilidad", keys: ["reliability"], numeric: true },
    { label: "Paradas", keys: ["stops"], numeric: true },
    { label: "Hrs Trab", keys: ["worked_hours"], numeric: true }
  ], "Sin KPI publicado.", 1);
}

async function readJson(request: Request): Promise<JsonObject> {
  const payload = await request.json().catch(() => null);
  if (!isObject(payload)) throw new HttpError(400, "Carga JSON invalida.");
  return payload;
}

function nowIso(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, "");
}

function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
}

function isObject(value: unknown): value is JsonObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseJsonObject(value: unknown, fallback: JsonObject = {}): JsonObject {
  if (isObject(value)) return value;
  if (typeof value !== "string" || !value.trim()) return fallback;
  try {
    const parsed = JSON.parse(value) as unknown;
    return isObject(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function stringValue(value: unknown): string {
  return value === null || value === undefined ? "" : String(value);
}

function escapeHtml(value: unknown): string {
  return stringValue(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;"
  }[character] || character));
}

function numberValue(value: unknown, fallback = 0): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const parsed = Number(String(value ?? "").replace(",", "."));
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeText(value: unknown): string {
  return stringValue(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function isoDateOrEmpty(value: unknown): string {
  const raw = stringValue(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function emptyCatalog(): JsonObject {
  return { ok: true, equipment: [], components: [], summary: {}, generated_at: nowIso() };
}

function emptyFilterSnapshot(): JsonObject {
  return { ok: true, inventory: [], generated_at: nowIso() };
}

function emptyEppSnapshot(): JsonObject {
  return { ok: true, epp: {}, items: [], generated_at: nowIso() };
}

function emptyPortalSnapshot(): JsonObject {
  return {
    ok: true,
    equipment: [],
    products: [],
    preventives: [],
    captures: [],
    availability: [],
    kpi_groups: [],
    kpi_reports: {},
    oil_kpi: { rows: [], totals: {}, columns: [] },
    tire_kpi: { rows: [], summary: {} },
    diesel: { records: [], days: [], rows: [], totals: {} },
    epp: { items: [], summary: {} },
    generated_at: nowIso()
  };
}

async function stats(env: Env): Promise<Response> {
  const captures = await captureCounts(env.DB);
  const catalog = await env.DB.prepare("SELECT updated_at FROM catalog_snapshot WHERE name = ?").bind("catalog").first<D1Row>();
  return json({
    ok: true,
    service: env.SERVICE_NAME || "mga-cloud-sync-worker",
    storage: env.EVIDENCIAS ? "cloudflare-d1-r2" : "cloudflare-d1",
    captures,
    catalog_updated_at: stringValue(catalog?.updated_at),
    generated_at: nowIso()
  });
}

async function captureCounts(db: D1Database): Promise<JsonObject> {
  const row = await db
    .prepare(
      "SELECT COUNT(*) AS total, " +
        "SUM(CASE WHEN desktop_imported_at IS NULL OR desktop_imported_at = '' THEN 1 ELSE 0 END) AS pending, " +
        "SUM(CASE WHEN desktop_imported_at IS NOT NULL AND desktop_imported_at <> '' THEN 1 ELSE 0 END) AS imported " +
        "FROM mobile_capture"
    )
    .first<D1Row>();
  return {
    total: Number(row?.total || 0),
    pending: Number(row?.pending || 0),
    imported: Number(row?.imported || 0)
  };
}

async function getCatalog(env: Env): Promise<Response> {
  const row = await env.DB.prepare("SELECT payload_json, updated_at FROM catalog_snapshot WHERE name = ?").bind("catalog").first<D1Row>();
  if (!row) return json(emptyCatalog());
  const payload = parseJsonObject(row.payload_json, emptyCatalog());
  return json({ ok: true, ...payload, cloud_updated_at: stringValue(row.updated_at) });
}

async function getSnapshot(db: D1Database, table: string, fallback: JsonObject): Promise<Response> {
  const row = await db.prepare(`SELECT payload_json, updated_at FROM ${table} WHERE name = ?`).bind("default").first<D1Row>();
  if (!row) return json(fallback);
  const payload = parseJsonObject(row.payload_json, fallback);
  return json({ ok: true, ...payload, cloud_updated_at: stringValue(row.updated_at) });
}

async function getDieselSnapshot(db: D1Database): Promise<Response> {
  const snapshot = await loadSnapshot(db, "portal_snapshot", "default", emptyPortalSnapshot());
  const diesel = objectFrom(snapshot.payload.diesel);
  return json({ ok: true, ...diesel, cloud_updated_at: snapshot.updatedAt });
}

async function saveSnapshot(request: Request, db: D1Database, table: string, name: string): Promise<Response> {
  const payload = await readJson(request);
  const updatedAt = nowIso();
  await db
    .prepare(`INSERT INTO ${table} (name, payload_json, updated_at) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET payload_json = excluded.payload_json, updated_at = excluded.updated_at`)
    .bind(name, JSON.stringify(payload), updatedAt)
    .run();
  return json({ ok: true, updated_at: updatedAt });
}

async function saveFilterSnapshot(request: Request, db: D1Database): Promise<Response> {
  const payload = await readJson(request);
  const rows = Array.isArray(payload.inventory) ? payload.inventory : [];
  const updatedAt = nowIso();
  await db
    .prepare("INSERT INTO filter_inventory_snapshot (name, payload_json, updated_at) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET payload_json = excluded.payload_json, updated_at = excluded.updated_at")
    .bind("default", JSON.stringify({ ok: true, generated_at: updatedAt, inventory: rows }), updatedAt)
    .run();
  return json({ ok: true, imported: rows.length, updated_at: updatedAt });
}

async function syncMobile(request: Request, env: Env): Promise<Response> {
  const payload = await readJson(request);
  const records = Array.isArray(payload.records) ? payload.records : null;
  if (!records) throw new HttpError(400, "El paquete no contiene lista de capturas.");

  const sourceDevice = stringValue(payload.device);
  const userName = stringValue(payload.user);
  const results: JsonObject[] = [];
  let created = 0;
  let skipped = 0;
  let errors = 0;

  for (const rawRecord of records) {
    try {
      if (!isObject(rawRecord)) throw new Error("Captura invalida.");
      const mobileId = stringValue(rawRecord.mobile_id).trim();
      if (!mobileId) throw new Error("Captura sin mobile_id.");

      const existing = await env.DB.prepare("SELECT id, payload_json, source_device, user_name, desktop_imported_at FROM mobile_capture WHERE mobile_id = ?")
        .bind(mobileId)
        .first<D1Row>();
      if (existing) {
        const existingPayload = parseJsonObject(existing.payload_json);
        for (const hose of mobileHoseChangeRows(existingPayload, stringValue(existing.source_device), stringValue(existing.user_name))) {
          await upsertHoseChange(env.DB, hose);
        }
        skipped += 1;
        results.push({
          mobile_id: mobileId,
          capture_id: Number(existing.id || 0),
          created: false,
          stored: true,
          desktop_imported: Boolean(existing.desktop_imported_at)
        });
        continue;
      }

      const photos = Array.isArray(rawRecord.photos) ? rawRecord.photos : [];
      const storedRecord = { ...rawRecord, photos: [] };
      const receivedAt = nowIso();
      const insert = await env.DB
        .prepare(
          "INSERT INTO mobile_capture (mobile_id, source_device, user_name, equipment_code, component_name, work_date, created_at, received_at, payload_json) " +
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
        )
        .bind(
          mobileId,
          sourceDevice,
          stringValue(rawRecord.user_name || userName),
          stringValue(rawRecord.equipment_code),
          stringValue(rawRecord.component_name),
          stringValue(rawRecord.work_date),
          stringValue(rawRecord.created_at),
          receivedAt,
          JSON.stringify(storedRecord)
        )
        .run();
      const captureId = Number(insert.meta.last_row_id || 0);

      let hoseChanges = 0;
      for (const hose of mobileHoseChangeRows(storedRecord, sourceDevice, stringValue(rawRecord.user_name || userName))) {
        await upsertHoseChange(env.DB, hose);
        hoseChanges += 1;
      }

      let evidenceCount = 0;
      for (const rawPhoto of photos) {
        if (!isObject(rawPhoto)) continue;
        const dataUrl = stringValue(rawPhoto.data);
        if (!dataUrl) continue;
        const fileName = safeFileName(stringValue(rawPhoto.name) || `${mobileId}.jpg`);
        const parsed = dataUrlToBytes(dataUrl, stringValue(rawPhoto.mime_type) || "image/jpeg");
        let key = "";
        if (env.EVIDENCIAS) {
          key = `evidencias/${safeSegment(mobileId)}/${Date.now()}-${evidenceCount + 1}-${fileName}`;
          await env.EVIDENCIAS.put(key, parsed.bytes, { httpMetadata: { contentType: parsed.mimeType } });
        }
        await env.DB
          .prepare("INSERT INTO mobile_photo (capture_id, mobile_id, file_name, mime_type, captured_at, r2_key, data_url, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)")
          .bind(captureId, mobileId, fileName, parsed.mimeType, stringValue(rawPhoto.captured_at), key, env.EVIDENCIAS ? "" : dataUrl, receivedAt)
          .run();
        evidenceCount += 1;
      }

      created += 1;
      results.push({
        mobile_id: mobileId,
        capture_id: captureId,
        created: true,
        stored: true,
        desktop_imported: false,
        evidence: evidenceCount,
        hose_changes: hoseChanges
      });
    } catch (error) {
      errors += 1;
      results.push({
        mobile_id: isObject(rawRecord) ? stringValue(rawRecord.mobile_id) : "",
        created: false,
        error: error instanceof Error ? error.message : "Error al guardar captura."
      });
    }
  }

  return json({
    ok: errors === 0,
    created,
    skipped,
    errors,
    stored: created + skipped,
    captures: await captureCounts(env.DB),
    results
  });
}

async function desktopPending(url: URL, env: Env): Promise<Response> {
  const limit = Math.min(Math.max(Number(url.searchParams.get("limit") || 200), 1), 1000);
  const includeImported = ["1", "true", "si", "yes"].includes((url.searchParams.get("include_imported") || "").toLowerCase());
  let sql = "SELECT * FROM mobile_capture";
  if (!includeImported) sql += " WHERE desktop_imported_at IS NULL OR desktop_imported_at = ''";
  sql += " ORDER BY id ASC LIMIT ?";
  const rows = (await env.DB.prepare(sql).bind(limit).all<D1Row>()).results || [];
  const records: JsonObject[] = [];

  for (const row of rows) {
    const photos = (await env.DB.prepare("SELECT file_name, mime_type, captured_at, r2_key, data_url FROM mobile_photo WHERE capture_id = ? ORDER BY id ASC")
      .bind(Number(row.id || 0))
      .all<D1Row>()).results || [];
    const photoPayload: JsonObject[] = [];
    for (const photo of photos) {
      const data = stringValue(photo.data_url) || await photoDataUrl(env.EVIDENCIAS, stringValue(photo.r2_key), stringValue(photo.mime_type));
      photoPayload.push({
        name: stringValue(photo.file_name),
        mime_type: stringValue(photo.mime_type) || "image/jpeg",
        captured_at: stringValue(photo.captured_at),
        data
      });
    }
    records.push({
      id: Number(row.id || 0),
      mobile_id: stringValue(row.mobile_id),
      source_device: stringValue(row.source_device),
      user_name: stringValue(row.user_name),
      payload: parseJsonObject(row.payload_json),
      photos: photoPayload,
      received_at: stringValue(row.received_at)
    });
  }

  return json({ ok: true, records, count: records.length, captures: await captureCounts(env.DB) });
}

async function desktopAck(request: Request, env: Env): Promise<Response> {
  const payload = await readJson(request);
  const ids = Array.isArray(payload.ids) ? payload.ids.map((item) => Number(item)).filter((item) => Number.isInteger(item) && item > 0) : [];
  if (!ids.length) return json({ ok: true, updated: 0, captures: await captureCounts(env.DB) });
  const importedAt = nowIso();
  let updated = 0;
  for (const id of ids) {
    const result = await env.DB.prepare("UPDATE mobile_capture SET desktop_imported_at = ? WHERE id = ?").bind(importedAt, id).run();
    updated += result.meta.changes || 0;
  }
  return json({ ok: true, updated, captures: await captureCounts(env.DB) });
}

async function mobileStatus(request: Request, env: Env): Promise<Response> {
  const payload = await readJson(request);
  const ids = Array.isArray(payload.mobile_ids) ? payload.mobile_ids.map((item) => stringValue(item).trim()).filter(Boolean) : [];
  const statuses: JsonObject[] = [];
  for (const mobileId of ids) {
    const row = await env.DB.prepare("SELECT id, desktop_imported_at FROM mobile_capture WHERE mobile_id = ?").bind(mobileId).first<D1Row>();
    statuses.push({
      mobile_id: mobileId,
      uploaded: Boolean(row),
      capture_id: Number(row?.id || 0) || undefined,
      desktop_imported: Boolean(row?.desktop_imported_at),
      desktop_imported_at: stringValue(row?.desktop_imported_at)
    });
  }
  return json({ ok: true, statuses });
}

function dataUrlToBytes(dataUrl: string, fallbackMime: string): { bytes: Uint8Array; mimeType: string } {
  const match = /^data:([^;,]+)?(;base64)?,(.*)$/i.exec(dataUrl);
  if (!match) {
    return { bytes: new TextEncoder().encode(dataUrl), mimeType: fallbackMime || "application/octet-stream" };
  }
  const mimeType = match[1] || fallbackMime || "application/octet-stream";
  const raw = match[2] ? atob(match[3]) : decodeURIComponent(match[3]);
  const bytes = new Uint8Array(raw.length);
  for (let index = 0; index < raw.length; index += 1) bytes[index] = raw.charCodeAt(index);
  return { bytes, mimeType };
}

async function photoDataUrl(bucket: R2Bucket | undefined, key: string, mimeType: string): Promise<string> {
  if (!bucket || !key) return "";
  const object = await bucket.get(key);
  if (!object) return "";
  const buffer = await object.arrayBuffer();
  return `data:${mimeType || object.httpMetadata?.contentType || "image/jpeg"};base64,${arrayBufferToBase64(buffer)}`;
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  }
  return btoa(binary);
}

function safeFileName(value: string): string {
  const cleaned = value.replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, "_").slice(0, 180);
  return cleaned || "evidencia.jpg";
}

function safeSegment(value: string): string {
  return value.replace(/[^A-Za-z0-9_.-]+/g, "_").slice(0, 120) || "sin_id";
}

function mobileHoseChangeRows(record: JsonObject, sourceDevice = "", userName = ""): JsonObject[] {
  if (normalizeText(record.kind) !== "MANGUERA") return [];
  const hoses = Array.isArray(record.hoses) ? record.hoses : [];
  const mobileId = stringValue(record.mobile_id).trim();
  const equipment = normalizeText(record.equipment_code || record.equipment);
  const changeDate = isoDateOrEmpty(record.work_date) || todayIso();
  const system = normalizeText(record.system || "HIDRAULICO");
  const location = stringValue(record.location).trim();
  const workType = stringValue(record.work_type).trim();
  const failureReason = stringValue(record.failure_reason).trim();
  const meter = Math.max(numberValue(record.meter), 0);
  const downtimeHours = Math.max(numberValue(record.downtime_hours), 0);
  const details = stringValue(record.details || record.observations).trim();
  const supervisor = stringValue(record.supervisor).trim();
  const mechanic = normalizeText(record.mechanic || record.technician || userName);
  const folio = stringValue(record.capture_folio).trim();
  const rows: JsonObject[] = [];

  hoses.forEach((item, itemIndex) => {
    if (!isObject(item)) return;
    const connectionType = stringValue(item.connection_type).trim();
    const connectionNumber = stringValue(item.connection_number).trim();
    const length = Math.max(numberValue(item.length), 0);
    const layers = stringValue(item.hose_layers).trim();
    if (!connectionNumber && length <= 0) return;
    const noteParts = [
      connectionType ? `Conexion: ${connectionType}` : "",
      layers ? `Malla: ${layers}` : "",
      workType ? `Tipo: ${workType}` : "",
      failureReason ? `Motivo: ${failureReason}` : "",
      meter > 0 ? `Horometro consulta: ${meter}` : "",
      downtimeHours > 0 ? `Horas paro: ${downtimeHours}` : "",
      supervisor ? `Supervisor: ${supervisor}` : "",
      sourceDevice ? `Dispositivo: ${sourceDevice}` : "",
      folio ? `Folio: ${folio}` : "",
      details
    ].filter(Boolean);
    rows.push({
      change_date: changeDate,
      equipment,
      system,
      part_type: "MANGUERA",
      diameter: connectionNumber,
      length_m: length,
      quantity: 1,
      unit_cost: 0,
      estimated_life_days: 30,
      estimated_weekly_qty: 0,
      failure_reason: [failureReason, location].filter(Boolean).join(" - "),
      technician: mechanic,
      notes: noteParts.join(" | "),
      source: "mobile",
      external_id: mobileId ? `mobile:${mobileId}:${itemIndex + 1}` : ""
    });
  });
  return rows;
}

async function getHoseChanges(url: URL, db: D1Database): Promise<Response> {
  const start = isoDateOrEmpty(url.searchParams.get("start")) || monthStartIso();
  const end = isoDateOrEmpty(url.searchParams.get("end")) || todayIso();
  const equipment = normalizeText(url.searchParams.get("equipment"));
  let sql = "SELECT * FROM hose_change WHERE deleted = 0 AND change_date >= ? AND change_date <= ?";
  const params: (string | number)[] = [start, end];
  if (equipment && equipment !== "TODOS" && equipment !== "TODOS LOS EQUIPOS") {
    sql += " AND equipment = ?";
    params.push(equipment);
  }
  sql += " ORDER BY change_date DESC, id DESC";
  const rows = (await db.prepare(sql).bind(...params).all<D1Row>()).results || [];
  return json(hoseReport(rows, start, end));
}

async function saveHoseChange(request: Request, db: D1Database): Promise<Response> {
  const payload = await readJson(request);
  const changeDate = isoDateOrEmpty(payload.change_date) || todayIso();
  const equipment = normalizeText(payload.equipment);
  if (!equipment) throw new HttpError(400, "Equipo requerido.");
  const row = await upsertHoseChange(db, { ...payload, change_date: changeDate, equipment, source: "web" });
  const report = await getHoseRowsForDates(db, changeDate, changeDate);
  return json({ ok: true, record: row, hoses: hoseReport(report, changeDate, changeDate) });
}

async function deleteHoseChange(request: Request, db: D1Database): Promise<Response> {
  const payload = await readJson(request);
  const id = Number(payload.id || 0);
  if (!Number.isInteger(id) || id <= 0) throw new HttpError(400, "Selecciona un cambio para eliminar.");
  await db.prepare("UPDATE hose_change SET deleted = 1, updated_at = ? WHERE id = ?").bind(nowIso(), id).run();
  return json({ ok: true });
}

async function getHoseRowsForDates(db: D1Database, start: string, end: string): Promise<D1Row[]> {
  return (await db
    .prepare("SELECT * FROM hose_change WHERE deleted = 0 AND change_date >= ? AND change_date <= ? ORDER BY change_date DESC, id DESC")
    .bind(start, end)
    .all<D1Row>()).results || [];
}

async function upsertHoseChange(db: D1Database, payload: JsonObject): Promise<JsonObject> {
  const externalId = stringValue(payload.external_id).trim();
  const existing = externalId
    ? await db.prepare("SELECT id FROM hose_change WHERE external_id = ?").bind(externalId).first<D1Row>()
    : payload.id
      ? await db.prepare("SELECT id FROM hose_change WHERE id = ?").bind(Number(payload.id)).first<D1Row>()
      : null;
  const rowPayload = {
    external_id: externalId,
    change_date: isoDateOrEmpty(payload.change_date) || todayIso(),
    equipment: normalizeText(payload.equipment),
    system: normalizeText(payload.system),
    part_type: normalizeText(payload.part_type || "MANGUERA") || "MANGUERA",
    diameter: normalizeText(payload.diameter).slice(0, 80),
    length_m: Math.max(numberValue(payload.length_m), 0),
    quantity: Math.max(numberValue(payload.quantity, 1), 0),
    unit_cost: Math.max(numberValue(payload.unit_cost), 0),
    estimated_life_days: Math.max(numberValue(payload.estimated_life_days, 30), 0),
    estimated_weekly_qty: Math.max(numberValue(payload.estimated_weekly_qty), 0),
    failure_reason: normalizeText(payload.failure_reason).slice(0, 220),
    technician: normalizeText(payload.technician).slice(0, 180),
    notes: stringValue(payload.notes).trim(),
    source: stringValue(payload.source || "web").trim().slice(0, 80),
    updated_at: nowIso()
  };
  if (existing) {
    await db
      .prepare(
        "UPDATE hose_change SET external_id = ?, change_date = ?, equipment = ?, system = ?, part_type = ?, diameter = ?, length_m = ?, quantity = ?, unit_cost = ?, estimated_life_days = ?, estimated_weekly_qty = ?, failure_reason = ?, technician = ?, notes = ?, source = ?, deleted = 0, updated_at = ? WHERE id = ?"
      )
      .bind(
        rowPayload.external_id,
        rowPayload.change_date,
        rowPayload.equipment,
        rowPayload.system,
        rowPayload.part_type,
        rowPayload.diameter,
        rowPayload.length_m,
        rowPayload.quantity,
        rowPayload.unit_cost,
        rowPayload.estimated_life_days,
        rowPayload.estimated_weekly_qty,
        rowPayload.failure_reason,
        rowPayload.technician,
        rowPayload.notes,
        rowPayload.source,
        rowPayload.updated_at,
        Number(existing.id)
      )
      .run();
    return { id: Number(existing.id), ...rowPayload };
  }

  const createdAt = nowIso();
  const insert = await db
    .prepare(
      "INSERT INTO hose_change (external_id, change_date, equipment, system, part_type, diameter, length_m, quantity, unit_cost, estimated_life_days, estimated_weekly_qty, failure_reason, technician, notes, source, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    )
    .bind(
      rowPayload.external_id,
      rowPayload.change_date,
      rowPayload.equipment,
      rowPayload.system,
      rowPayload.part_type,
      rowPayload.diameter,
      rowPayload.length_m,
      rowPayload.quantity,
      rowPayload.unit_cost,
      rowPayload.estimated_life_days,
      rowPayload.estimated_weekly_qty,
      rowPayload.failure_reason,
      rowPayload.technician,
      rowPayload.notes,
      rowPayload.source,
      createdAt,
      rowPayload.updated_at
    )
    .run();
  return { id: Number(insert.meta.last_row_id || 0), ...rowPayload };
}

function hoseReport(rows: D1Row[], start: string, end: string): JsonObject {
  const periodDays = Math.max(daysBetween(start, end) + 1, 1);
  const records: JsonObject[] = [];
  const summary = new Map<string, JsonObject & { systemsSet: Set<string> }>();
  for (const row of rows) {
    const payload = hosePayload(row);
    const qty = Math.max(numberValue(payload.quantity), 0);
    const unitCost = Math.max(numberValue(payload.unit_cost), 0);
    const lifeDays = Math.max(numberValue(payload.estimated_life_days), 0);
    let weekly = Math.max(numberValue(payload.estimated_weekly_qty), 0);
    if (weekly <= 0 && lifeDays > 0) weekly = (qty * 7) / lifeDays;
    const estimatedQty = Math.max((weekly * periodDays) / 7, 0);
    const totalCost = qty * unitCost;
    payload.estimated_qty_period = estimatedQty;
    payload.total_cost = totalCost;
    records.push(payload);
    const key = `${stringValue(payload.equipment) || "SIN EQUIPO"}\n${stringValue(payload.part_type) || "MANGUERA"}`;
    const bucket = summary.get(key) || {
      equipment: stringValue(payload.equipment) || "SIN EQUIPO",
      part_type: stringValue(payload.part_type) || "MANGUERA",
      systems: "",
      systemsSet: new Set<string>(),
      changes: 0,
      quantity: 0,
      estimated_qty: 0,
      variance: 0,
      total_cost: 0
    };
    bucket.changes = numberValue(bucket.changes) + 1;
    bucket.quantity = numberValue(bucket.quantity) + qty;
    bucket.estimated_qty = numberValue(bucket.estimated_qty) + estimatedQty;
    bucket.total_cost = numberValue(bucket.total_cost) + totalCost;
    if (payload.system) bucket.systemsSet.add(stringValue(payload.system));
    summary.set(key, bucket);
  }
  const summaryRows = Array.from(summary.values()).map((bucket) => {
    const quantity = numberValue(bucket.quantity);
    const estimatedQty = numberValue(bucket.estimated_qty);
    return {
      equipment: bucket.equipment,
      part_type: bucket.part_type,
      systems: Array.from(bucket.systemsSet).sort().join(", "),
      changes: bucket.changes,
      quantity,
      estimated_qty: estimatedQty,
      variance: quantity - estimatedQty,
      total_cost: bucket.total_cost
    };
  });
  summaryRows.sort((left, right) => numberValue(right.quantity) - numberValue(left.quantity) || stringValue(left.equipment).localeCompare(stringValue(right.equipment)));
  const realQty = summaryRows.reduce((sum, row) => sum + numberValue(row.quantity), 0);
  const estimatedQty = summaryRows.reduce((sum, row) => sum + numberValue(row.estimated_qty), 0);
  return {
    ok: true,
    start,
    end,
    period_days: periodDays,
    records,
    summary: summaryRows,
    totals: {
      changes: records.length,
      real_qty: realQty,
      estimated_qty: estimatedQty,
      variance: realQty - estimatedQty,
      total_cost: summaryRows.reduce((sum, row) => sum + numberValue(row.total_cost), 0),
      critical_equipment: summaryRows[0]?.equipment || "S/D"
    }
  };
}

function hosePayload(row: D1Row): JsonObject {
  return {
    id: Number(row.id || 0),
    source: stringValue(row.source || "web"),
    external_id: stringValue(row.external_id),
    change_date: stringValue(row.change_date),
    equipment: stringValue(row.equipment),
    system: stringValue(row.system),
    part_type: stringValue(row.part_type),
    diameter: stringValue(row.diameter),
    length_m: numberValue(row.length_m),
    quantity: numberValue(row.quantity),
    unit_cost: numberValue(row.unit_cost),
    estimated_life_days: numberValue(row.estimated_life_days),
    estimated_weekly_qty: numberValue(row.estimated_weekly_qty),
    failure_reason: stringValue(row.failure_reason),
    technician: stringValue(row.technician),
    notes: stringValue(row.notes),
    updated_at: stringValue(row.updated_at)
  };
}

function monthStartIso(): string {
  const now = new Date();
  return `${now.getUTCFullYear()}-${String(now.getUTCMonth() + 1).padStart(2, "0")}-01`;
}

function daysBetween(start: string, end: string): number {
  const left = Date.parse(`${start}T00:00:00Z`);
  const right = Date.parse(`${end}T00:00:00Z`);
  if (!Number.isFinite(left) || !Number.isFinite(right)) return 0;
  return Math.round(Math.abs(right - left) / 86400000);
}
