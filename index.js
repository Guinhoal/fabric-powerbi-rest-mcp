import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { Connection, Request as TdsRequest } from "tedious";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  isInitializeRequest
} from "@modelcontextprotocol/sdk/types.js";
import { execSync } from "node:child_process";
import { appendFileSync, mkdirSync } from "node:fs";
import { dirname } from "node:path";
import net from "node:net";
import express from "express";
import crypto from "node:crypto";

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const LOG_FILE =
  process.env.MCP_LOG_FILE ??
  (process.platform === "win32"
    ? "C:\\MCPServers\\fabric-powerbi-rest-mcp\\mcp-debug.log"
    : "/var/log/mcp/mcp-debug.log");

const HTTP_MODE = process.env.MCP_TRANSPORT === "http";
const PORT = parseInt(process.env.PORT ?? "8000", 10);

const POWERBI_RESOURCE = "https://analysis.windows.net/powerbi/api";
const FABRIC_RESOURCE = "https://api.fabric.microsoft.com";
const PBI_BASE = "https://api.powerbi.com/v1.0/myorg";
const FABRIC_BASE = "https://api.fabric.microsoft.com/v1";

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------

function log(message) {
  try {
    mkdirSync(dirname(LOG_FILE), { recursive: true });
    appendFileSync(LOG_FILE, `[${new Date().toISOString()}] ${message}\n`, "utf8");
  } catch {
    // Keep stdout reserved for JSON-RPC in stdio mode.
  }
}

function errorResponse(toolName, error) {
  const details = {
    ok: false,
    tool: toolName,
    errorMessage: error?.message ?? String(error),
    stack: error?.stack ?? null
  };
  log(`ERROR in ${toolName}: ${JSON.stringify(details)}`);
  return { content: [{ type: "text", text: JSON.stringify(details, null, 2) }] };
}

function successResponse(payload) {
  return { content: [{ type: "text", text: JSON.stringify(payload, null, 2) }] };
}

// ---------------------------------------------------------------------------
// Auth: Service Principal (Docker) or az CLI (local dev)
// ---------------------------------------------------------------------------

const tokenCache = new Map();

async function getAzureToken(resource) {
  const cached = tokenCache.get(resource);
  if (cached && cached.expiresAt > Date.now()) return cached.token;

  log(`Obtaining token for resource: ${resource}`);

  if (process.env.AZURE_CLIENT_ID) {
    const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

    if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
      throw new Error("Defina AZURE_TENANT_ID, AZURE_CLIENT_ID e AZURE_CLIENT_SECRET.");
    }

    const params = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: AZURE_CLIENT_ID,
      client_secret: AZURE_CLIENT_SECRET,
      scope: `${resource}/.default`
    });

    const res = await fetch(
      `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params
      }
    );

    const data = await res.json();

    if (!res.ok) {
      throw new Error(`Entra ID error: ${data.error_description ?? JSON.stringify(data)}`);
    }

    const token = data.access_token;
    const expiresAt = Date.now() + (data.expires_in - 60) * 1000;
    tokenCache.set(resource, { token, expiresAt });
    log(`Token OK (SP). Expires in ~${data.expires_in}s`);
    return token;
  }

  const token = execSync(
    `az account get-access-token --resource "${resource}" --query accessToken -o tsv`,
    { encoding: "utf8", windowsHide: true, shell: true, env: { ...process.env } }
  ).trim();

  if (!token) throw new Error("Token vazio retornado pelo Azure CLI.");
  tokenCache.set(resource, { token, expiresAt: Date.now() + 50 * 60 * 1000 });
  log(`Token OK (az CLI). Length: ${token.length}`);
  return token;
}

// ---------------------------------------------------------------------------
// HTTP helpers
// ---------------------------------------------------------------------------

async function getJson(url, token) {
  log(`GET ${url}`);
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }
  });
  const text = await response.text();
  if (!response.ok) throw new Error(`HTTP ${response.status} em ${url}: ${text}`);
  return text ? JSON.parse(text) : {};
}

async function pbiGet(path) {
  return getJson(`${PBI_BASE}${path}`, await getAzureToken(POWERBI_RESOURCE));
}

async function pbiGetValue(path) {
  const data = await pbiGet(path);
  return Array.isArray(data.value) ? data.value : [];
}

// ---------------------------------------------------------------------------
// Workspaces
// ---------------------------------------------------------------------------

async function listFabricWorkspaces() {
  const token = await getAzureToken(FABRIC_RESOURCE);
  let url = `${FABRIC_BASE}/workspaces`;
  const all = [];

  while (url) {
    const data = await getJson(url, token);
    if (Array.isArray(data.value)) all.push(...data.value);
    url = data.continuationUri || null;
  }

  return all.map((w) => ({
    id: w.id,
    displayName: w.displayName,
    description: w.description ?? null,
    type: w.type ?? null,
    capacityId: w.capacityId ?? null
  }));
}

async function listPowerBiWorkspaces() {
  const token = await getAzureToken(POWERBI_RESOURCE);
  const all = [];
  let skip = 0;

  while (true) {
    const data = await getJson(`${PBI_BASE}/groups?$top=5000&$skip=${skip}`, token);
    const batch = Array.isArray(data.value) ? data.value : [];
    all.push(...batch);
    if (batch.length < 5000) break;
    skip += 5000;
  }

  return all.map((w) => ({
    id: w.id,
    name: w.name,
    type: w.type ?? null,
    isReadOnly: w.isReadOnly ?? null,
    isOnDedicatedCapacity: w.isOnDedicatedCapacity ?? null,
    capacityId: w.capacityId ?? null,
    state: w.state ?? null
  }));
}

async function resolveWorkspaceId(nameOrId) {
  const guidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (guidRe.test(nameOrId)) return nameOrId;

  const workspaces = await listPowerBiWorkspaces();
  const target = nameOrId.trim().toLowerCase();
  const exact = workspaces.find((w) => (w.name ?? "").toLowerCase() === target);
  if (exact) return exact.id;

  const partial = workspaces.filter((w) => (w.name ?? "").toLowerCase().includes(target));
  if (partial.length === 1) return partial[0].id;
  if (partial.length > 1) {
    throw new Error(
      `"${nameOrId}" casou com ${partial.length} workspaces: ` +
        partial.map((w) => `"${w.name}" (${w.id})`).join(", ")
    );
  }

  throw new Error(`Workspace nao encontrado: "${nameOrId}"`);
}

// ---------------------------------------------------------------------------
// Datasets
// ---------------------------------------------------------------------------

async function listDatasets(groupId) {
  const rows = await pbiGetValue(`/groups/${groupId}/datasets`);
  return rows.map((d) => ({
    id: d.id,
    name: d.name,
    configuredBy: d.configuredBy ?? null,
    isRefreshable: d.isRefreshable ?? null,
    isOnPremGatewayRequired: d.isOnPremGatewayRequired ?? null,
    targetStorageMode: d.targetStorageMode ?? null,
    createdDate: d.createdDate ?? null,
    webUrl: d.webUrl ?? null
  }));
}

async function getDataset(groupId, datasetId) {
  return pbiGet(`/groups/${groupId}/datasets/${datasetId}`);
}

async function getRefreshHistory(groupId, datasetId, top) {
  const q = top ? `?$top=${encodeURIComponent(top)}` : "";
  return pbiGetValue(`/groups/${groupId}/datasets/${datasetId}/refreshes${q}`);
}

async function getRefreshExecutionDetails(groupId, datasetId, refreshId) {
  return pbiGet(`/groups/${groupId}/datasets/${datasetId}/refreshes/${refreshId}`);
}

async function getRefreshSchedule(groupId, datasetId) {
  return pbiGet(`/groups/${groupId}/datasets/${datasetId}/refreshSchedule`);
}

async function getDatasetDatasources(groupId, datasetId) {
  return pbiGetValue(`/groups/${groupId}/datasets/${datasetId}/datasources`);
}

async function resolveDatasetId(groupId, nameOrId) {
  const guidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (guidRe.test(nameOrId)) return nameOrId;

  const datasets = await listDatasets(groupId);
  const target = nameOrId.trim().toLowerCase();
  const exact = datasets.find((d) => (d.name ?? "").toLowerCase() === target);
  if (exact) return exact.id;

  const partial = datasets.filter((d) => (d.name ?? "").toLowerCase().includes(target));
  if (partial.length === 1) return partial[0].id;
  if (partial.length > 1) {
    throw new Error(
      `"${nameOrId}" casou com ${partial.length} datasets: ` +
        partial.map((d) => `"${d.name}" (${d.id})`).join(", ")
    );
  }

  throw new Error(`Dataset nao encontrado: "${nameOrId}" no workspace ${groupId}`);
}

async function findDatasetLastRefresh(workspaceNameOrId, datasetNameOrId) {
  const groupId = await resolveWorkspaceId(workspaceNameOrId);
  const datasetId = await resolveDatasetId(groupId, datasetNameOrId);
  const history = await getRefreshHistory(groupId, datasetId, 5);
  let meta = null;
  try {
    meta = await getDataset(groupId, datasetId);
  } catch {
    // Keep refresh output even when metadata is unavailable.
  }

  return {
    groupId,
    datasetId,
    datasetName: meta?.name ?? null,
    isRefreshable: meta?.isRefreshable ?? null,
    lastRefresh: history[0] ?? null,
    recentRefreshes: history
  };
}

// ---------------------------------------------------------------------------
// Reports / dashboards / dataflows
// ---------------------------------------------------------------------------

async function listReports(groupId) {
  const rows = await pbiGetValue(`/groups/${groupId}/reports`);
  return rows.map((r) => ({
    id: r.id,
    name: r.name,
    reportType: r.reportType ?? null,
    datasetId: r.datasetId ?? null,
    webUrl: r.webUrl ?? null,
    embedUrl: r.embedUrl ?? null
  }));
}

async function listDashboards(groupId) {
  const rows = await pbiGetValue(`/groups/${groupId}/dashboards`);
  return rows.map((d) => ({
    id: d.id,
    displayName: d.displayName ?? d.name ?? null,
    isReadOnly: d.isReadOnly ?? null,
    embedUrl: d.embedUrl ?? null
  }));
}

async function listDataflows(groupId) {
  const rows = await pbiGetValue(`/groups/${groupId}/dataflows`);
  return rows.map((d) => ({
    objectId: d.objectId ?? d.id ?? null,
    name: d.name ?? null,
    description: d.description ?? null,
    configuredBy: d.configuredBy ?? null,
    modelUrl: d.modelUrl ?? null
  }));
}

async function getDataflowDatasources(groupId, dataflowId) {
  return pbiGetValue(`/groups/${groupId}/dataflows/${dataflowId}/datasources`);
}

async function getWorkspaceInventory(workspaceNameOrId) {
  const groupId = await resolveWorkspaceId(workspaceNameOrId);
  const settle = async (fn) => {
    try {
      return { ok: true, items: await fn(groupId) };
    } catch (e) {
      return { ok: false, error: e.message };
    }
  };

  const [datasets, reports, dashboards, dataflows] = await Promise.all([
    settle(listDatasets),
    settle(listReports),
    settle(listDashboards),
    settle(listDataflows)
  ]);

  return { groupId, datasets, reports, dashboards, dataflows };
}

// ---------------------------------------------------------------------------
// SQL Endpoint (Fabric Lakehouse / Warehouse)
// ---------------------------------------------------------------------------

function normalizeSqlEndpoint(sqlEndpoint) {
  return String(sqlEndpoint)
    .trim()
    .replace(/^tcp:/i, "")
    .replace(/,1433$/i, "");
}

function buildSafeQuery(query, maxRows) {
  const trimmed = String(query).trim().replace(/;+\s*$/, "");
  if (!trimmed) throw new Error("Query vazia.");

  const boundedMaxRows = Math.max(1, Math.min(Number(maxRows) || 100, 1000));
  const isSelect = /^SELECT\s/i.test(trimmed);

  return {
    isSelect,
    maxRows: boundedMaxRows,
    query: isSelect ? `SELECT TOP ${boundedMaxRows} * FROM (${trimmed}) AS _mcp_result` : trimmed
  };
}

function getServicePrincipalSqlAuth() {
  const { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET } = process.env;

  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID || !AZURE_CLIENT_SECRET) {
    throw new Error("Defina AZURE_TENANT_ID, AZURE_CLIENT_ID e AZURE_CLIENT_SECRET.");
  }

  return {
    type: "azure-active-directory-service-principal-secret",
    options: {
      tenantId: AZURE_TENANT_ID,
      clientId: AZURE_CLIENT_ID,
      clientSecret: AZURE_CLIENT_SECRET
    }
  };
}

async function executeSqlQuery(sqlEndpoint, database, query, maxRows = 100) {
  const server = normalizeSqlEndpoint(sqlEndpoint);
  const safe = buildSafeQuery(query, maxRows);

  log(`[SQL] Conectando em ${server}, db=${database}`);
  log("[SQL] Auth type: azure-active-directory-service-principal-secret");
  log(`[SQL] Query: ${safe.query.substring(0, 500)}`);

  return new Promise((resolve, reject) => {
    const config = {
      server,
      authentication: getServicePrincipalSqlAuth(),
      options: {
        database,
        port: 1433,
        encrypt: true,
        trustServerCertificate: false,
        connectTimeout: 30000,
        requestTimeout: 60000,
        rowCollectionOnDone: false
      }
    };

    const connection = new Connection(config);
    const rows = [];
    let columns = [];
    let settled = false;

    function finishOk(payload) {
      if (settled) return;
      settled = true;
      try {
        connection.close();
      } catch {
        // Ignore close errors.
      }
      resolve(payload);
    }

    function finishError(err, phase = "fatal") {
      if (settled) return;
      settled = true;
      try {
        connection.close();
      } catch {
        // Ignore close errors.
      }
      log(`[SQL] Erro ${phase}: ${err?.message ?? String(err)}`);
      reject(err);
    }

    connection.on("debug", (message) => {
      if (process.env.MCP_SQL_DEBUG === "1") log(`[SQL DEBUG] ${message}`);
    });

    connection.on("error", (err) => finishError(err, "fatal"));

    connection.on("errorMessage", (msg) => {
      log(`[SQL] Server error ${msg.number ?? ""}: ${msg.message}`);
    });

    connection.on("infoMessage", (msg) => {
      log(`[SQL] Server info ${msg.number ?? ""}: ${msg.message}`);
    });

    connection.on("connect", (err) => {
      if (err) return finishError(err, "de conexao");

      log("[SQL] Conectado. Executando query...");

      const request = new TdsRequest(safe.query, (requestErr, rowCount) => {
        if (requestErr) return finishError(requestErr, "na query");

        log(`[SQL] Query concluida. ${rowCount} linhas reportadas, ${rows.length} linhas coletadas.`);
        return finishOk({
          rows,
          rowCount,
          columns: columns.map((c) => c.colName)
        });
      });

      request.on("columnMetadata", (cols) => {
        columns = cols;
      });

      request.on("row", (cols) => {
        const row = {};
        cols.forEach((col) => {
          row[col.metadata.colName] = col.value;
        });
        rows.push(row);
      });

      request.on("error", (requestErr) => {
        finishError(requestErr, "do request");
      });

      connection.execSql(request);
    });

    connection.connect();
  });
}

async function testSqlTcpConnection(sqlEndpoint) {
  const server = normalizeSqlEndpoint(sqlEndpoint);

  return new Promise((resolve) => {
    const startedAt = Date.now();
    const socket = net.createConnection({ host: server, port: 1433, timeout: 10000 });

    socket.on("connect", () => {
      const elapsedMs = Date.now() - startedAt;
      socket.destroy();
      resolve({ ok: true, server, port: 1433, elapsedMs });
    });

    socket.on("timeout", () => {
      socket.destroy();
      resolve({ ok: false, server, port: 1433, error: "TCP timeout" });
    });

    socket.on("error", (err) => {
      resolve({ ok: false, server, port: 1433, error: err.message, code: err.code ?? null });
    });
  });
}

async function listLakehouseSqlEndpoints(workspaceNameOrId) {
  const workspaceId = await resolveWorkspaceId(workspaceNameOrId);
  const token = await getAzureToken(FABRIC_RESOURCE);
  const data = await getJson(`${FABRIC_BASE}/workspaces/${workspaceId}/lakehouses`, token);
  const items = Array.isArray(data.value) ? data.value : [];

  return {
    workspaceId,
    total: items.length,
    lakehouses: items.map((lh) => ({
      id: lh.id,
      displayName: lh.displayName,
      sql_endpoint: lh.properties?.sqlEndpointProperties?.connectionString ?? null,
      database: lh.displayName,
      provisioningStatus: lh.properties?.sqlEndpointProperties?.provisioningStatus ?? null
    }))
  };
}

// ---------------------------------------------------------------------------
// Tool definitions
// ---------------------------------------------------------------------------

const TOOL_DEFINITIONS = [
  {
    name: "list_fabric_workspaces",
    description: "Lista todos os workspaces do Microsoft Fabric acessiveis via Fabric REST API.",
    inputSchema: { type: "object", properties: {} }
  },
  {
    name: "list_powerbi_workspaces",
    description: "Lista todos os workspaces do Power BI acessiveis via Power BI REST API.",
    inputSchema: { type: "object", properties: {} }
  },
  {
    name: "compare_fabric_and_powerbi_workspaces",
    description: "Lista workspaces do Fabric e do Power BI e mostra diferencas entre as duas APIs.",
    inputSchema: { type: "object", properties: {} }
  },
  {
    name: "list_datasets",
    description: "Lista os datasets/modelos semanticos de um workspace do Power BI.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string", description: "Nome ou id do workspace." }
      },
      required: ["workspace"]
    }
  },
  {
    name: "get_dataset",
    description: "Retorna metadados completos de um dataset especifico.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataset: { type: "string", description: "Nome ou id do dataset." }
      },
      required: ["workspace", "dataset"]
    }
  },
  {
    name: "get_dataset_refresh_history",
    description: "Retorna o historico de atualizacoes de um dataset.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataset: { type: "string" },
        top: { type: "integer" }
      },
      required: ["workspace", "dataset"]
    }
  },
  {
    name: "get_refresh_execution_details",
    description: "Detalhes de execucao de um refresh especifico.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataset: { type: "string" },
        refreshId: { type: "string" }
      },
      required: ["workspace", "dataset", "refreshId"]
    }
  },
  {
    name: "get_dataset_refresh_schedule",
    description: "Retorna o agendamento de atualizacao configurado para um dataset.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataset: { type: "string" }
      },
      required: ["workspace", "dataset"]
    }
  },
  {
    name: "get_dataset_datasources",
    description: "Lista as fontes de dados usadas por um dataset.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataset: { type: "string" }
      },
      required: ["workspace", "dataset"]
    }
  },
  {
    name: "find_dataset_last_refresh",
    description: "Resolve nomes para ids e retorna a ultima atualizacao e as 5 mais recentes.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataset: { type: "string" }
      },
      required: ["workspace", "dataset"]
    }
  },
  {
    name: "list_reports",
    description: "Lista os relatorios de um workspace com datasetId associado.",
    inputSchema: {
      type: "object",
      properties: { workspace: { type: "string" } },
      required: ["workspace"]
    }
  },
  {
    name: "list_dashboards",
    description: "Lista os dashboards de um workspace do Power BI.",
    inputSchema: {
      type: "object",
      properties: { workspace: { type: "string" } },
      required: ["workspace"]
    }
  },
  {
    name: "list_dataflows",
    description: "Lista os dataflows de um workspace.",
    inputSchema: {
      type: "object",
      properties: { workspace: { type: "string" } },
      required: ["workspace"]
    }
  },
  {
    name: "get_dataflow_datasources",
    description: "Lista as fontes de dados de um dataflow especifico.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: { type: "string" },
        dataflowId: { type: "string" }
      },
      required: ["workspace", "dataflowId"]
    }
  },
  {
    name: "get_workspace_inventory",
    description: "Inventario consolidado de um workspace: datasets, reports, dashboards e dataflows.",
    inputSchema: {
      type: "object",
      properties: { workspace: { type: "string" } },
      required: ["workspace"]
    }
  },
  {
    name: "execute_sql_query",
    description: "Executa uma query T-SQL no SQL Endpoint de um Lakehouse ou Warehouse do Microsoft Fabric.",
    inputSchema: {
      type: "object",
      properties: {
        sql_endpoint: {
          type: "string",
          description: "FQDN do SQL Endpoint. Ex: abc.datawarehouse.fabric.microsoft.com"
        },
        database: {
          type: "string",
          description: "Nome do Lakehouse ou Warehouse."
        },
        query: {
          type: "string",
          description: "Query T-SQL a executar. Prefira SELECT."
        },
        max_rows: {
          type: "integer",
          description: "Limite de linhas retornadas para SELECT. Default 100, maximo 1000.",
          default: 100
        }
      },
      required: ["sql_endpoint", "database", "query"]
    }
  },
  {
    name: "test_sql_tcp_connection",
    description: "Testa conectividade TCP do container ate o SQL Endpoint na porta 1433.",
    inputSchema: {
      type: "object",
      properties: {
        sql_endpoint: {
          type: "string",
          description: "FQDN do SQL Endpoint."
        }
      },
      required: ["sql_endpoint"]
    }
  },
  {
    name: "list_lakehouse_sql_endpoints",
    description: "Lista Lakehouses de um workspace do Fabric e retorna o SQL Endpoint de cada um.",
    inputSchema: {
      type: "object",
      properties: {
        workspace: {
          type: "string",
          description: "Nome ou id do workspace."
        }
      },
      required: ["workspace"]
    }
  }
];

// ---------------------------------------------------------------------------
// Tool dispatcher
// ---------------------------------------------------------------------------

async function dispatchTool(toolName, args) {
  switch (toolName) {
    case "list_fabric_workspaces": {
      const workspaces = await listFabricWorkspaces();
      return successResponse({
        ok: true,
        total: workspaces.length,
        source: "Fabric REST API /v1/workspaces",
        workspaces
      });
    }
    case "list_powerbi_workspaces": {
      const workspaces = await listPowerBiWorkspaces();
      return successResponse({
        ok: true,
        total: workspaces.length,
        source: "Power BI REST API /v1.0/myorg/groups",
        workspaces
      });
    }
    case "compare_fabric_and_powerbi_workspaces": {
      const [fab, pbi] = await Promise.all([listFabricWorkspaces(), listPowerBiWorkspaces()]);
      const fabIds = new Set(fab.map((w) => w.id));
      const pbiIds = new Set(pbi.map((w) => w.id));
      return successResponse({
        ok: true,
        fabricTotal: fab.length,
        powerBiTotal: pbi.length,
        inBoth: pbi.filter((w) => fabIds.has(w.id)),
        onlyInFabric: fab.filter((w) => !pbiIds.has(w.id)),
        onlyInPowerBi: pbi.filter((w) => !fabIds.has(w.id))
      });
    }
    case "list_datasets": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasets = await listDatasets(groupId);
      return successResponse({ ok: true, groupId, total: datasets.length, datasets });
    }
    case "get_dataset": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasetId = await resolveDatasetId(groupId, args.dataset);
      return successResponse({ ok: true, groupId, dataset: await getDataset(groupId, datasetId) });
    }
    case "get_dataset_refresh_history": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasetId = await resolveDatasetId(groupId, args.dataset);
      const refreshes = await getRefreshHistory(groupId, datasetId, args.top);
      return successResponse({ ok: true, groupId, datasetId, total: refreshes.length, refreshes });
    }
    case "get_refresh_execution_details": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasetId = await resolveDatasetId(groupId, args.dataset);
      const details = await getRefreshExecutionDetails(groupId, datasetId, args.refreshId);
      return successResponse({ ok: true, groupId, datasetId, details });
    }
    case "get_dataset_refresh_schedule": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasetId = await resolveDatasetId(groupId, args.dataset);
      const schedule = await getRefreshSchedule(groupId, datasetId);
      return successResponse({ ok: true, groupId, datasetId, schedule });
    }
    case "get_dataset_datasources": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasetId = await resolveDatasetId(groupId, args.dataset);
      const datasources = await getDatasetDatasources(groupId, datasetId);
      return successResponse({ ok: true, groupId, datasetId, total: datasources.length, datasources });
    }
    case "find_dataset_last_refresh": {
      return successResponse({ ok: true, ...(await findDatasetLastRefresh(args.workspace, args.dataset)) });
    }
    case "list_reports": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const reports = await listReports(groupId);
      return successResponse({ ok: true, groupId, total: reports.length, reports });
    }
    case "list_dashboards": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const dashboards = await listDashboards(groupId);
      return successResponse({ ok: true, groupId, total: dashboards.length, dashboards });
    }
    case "list_dataflows": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const dataflows = await listDataflows(groupId);
      return successResponse({ ok: true, groupId, total: dataflows.length, dataflows });
    }
    case "get_dataflow_datasources": {
      const groupId = await resolveWorkspaceId(args.workspace);
      const datasources = await getDataflowDatasources(groupId, args.dataflowId);
      return successResponse({
        ok: true,
        groupId,
        dataflowId: args.dataflowId,
        total: datasources.length,
        datasources
      });
    }
    case "get_workspace_inventory": {
      return successResponse({ ok: true, ...(await getWorkspaceInventory(args.workspace)) });
    }
    case "execute_sql_query": {
      const result = await executeSqlQuery(
        args.sql_endpoint,
        args.database,
        args.query,
        args.max_rows ?? 100
      );
      return successResponse({
        ok: true,
        sql_endpoint: args.sql_endpoint,
        database: args.database,
        rowCount: result.rowCount,
        columns: result.columns,
        rows: result.rows
      });
    }
    case "test_sql_tcp_connection": {
      return successResponse(await testSqlTcpConnection(args.sql_endpoint));
    }
    case "list_lakehouse_sql_endpoints": {
      const result = await listLakehouseSqlEndpoints(args.workspace);
      return successResponse({ ok: true, ...result });
    }
    default:
      return errorResponse(toolName, new Error(`Ferramenta desconhecida: ${toolName}`));
  }
}

// ---------------------------------------------------------------------------
// Server factory
// ---------------------------------------------------------------------------

function createMcpServer() {
  const server = new Server(
    { name: "fabric-powerbi-rest-mcp", version: "3.1.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: TOOL_DEFINITIONS
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const toolName = request.params.name;
    const args = request.params.arguments ?? {};
    try {
      return await dispatchTool(toolName, args);
    } catch (error) {
      return errorResponse(toolName, error);
    }
  });

  return server;
}

// ---------------------------------------------------------------------------
// Startup
// ---------------------------------------------------------------------------

if (HTTP_MODE) {
  const app = express();

  const authMiddleware = (req, res, next) => {
    const key = process.env.MCP_API_KEY || process.env.MCP_BEARER_TOKEN;

    if (!key) {
      log("[AUTH] MCP_API_KEY/MCP_BEARER_TOKEN nao definido. Servidor aberto.");
      return next();
    }

    const authorization = req.headers.authorization || "";
    const apiKey = req.headers["x-api-key"] || "";
    const valid = authorization === `Bearer ${key}` || apiKey === key;

    if (!valid) {
      log(
        `[AUTH] Rejeitado: ${req.ip} -> ${req.path}. Header Authorization: ${
          authorization ? "presente" : "ausente"
        }, x-api-key: ${apiKey ? "presente" : "ausente"}`
      );

      return res.status(401).json({ error: "Unauthorized" });
    }

    return next();
  };

  const sessions = new Map();

  app.get("/sse", authMiddleware, async (req, res) => {
    log(`[SSE] Nova sessao de ${req.ip}`);
    const transport = new SSEServerTransport("/message", res);
    const mcpServer = createMcpServer();
    sessions.set(transport.sessionId, { transport, mcpServer });

    res.on("close", () => {
      sessions.delete(transport.sessionId);
      log(`[SSE] Sessao encerrada: ${transport.sessionId}`);
    });

    await mcpServer.connect(transport);
  });

  app.post("/message", authMiddleware, async (req, res) => {
    const session = sessions.get(req.query.sessionId);
    if (!session) return res.status(404).json({ error: "Session not found" });
    return session.transport.handlePostMessage(req, res);
  });

  app.get("/health", (req, res) =>
    res.json({ ok: true, version: "3.1.0", sessions: sessions.size })
  );

  const streamableSessions = new Map();
  const jsonParser = express.json({ limit: "10mb" });

  app.post("/mcp", authMiddleware, jsonParser, async (req, res) => {
    try {
      const sessionId = req.headers["mcp-session-id"];
      let transport;

      if (sessionId && streamableSessions.has(sessionId)) {
        transport = streamableSessions.get(sessionId);
      } else if (!sessionId && isInitializeRequest(req.body)) {
        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => crypto.randomUUID(),
          onsessioninitialized: (newSessionId) => {
            streamableSessions.set(newSessionId, transport);
            log(`[MCP] Streamable session initialized: ${newSessionId}`);
          }
        });

        transport.onclose = () => {
          if (transport.sessionId) {
            streamableSessions.delete(transport.sessionId);
            log(`[MCP] Streamable session closed: ${transport.sessionId}`);
          }
        };

        const mcpServer = createMcpServer();
        await mcpServer.connect(transport);
      } else {
        return res.status(400).json({
          jsonrpc: "2.0",
          error: {
            code: -32000,
            message: "Bad Request: invalid or missing MCP session."
          },
          id: null
        });
      }

      return transport.handleRequest(req, res, req.body);
    } catch (error) {
      log(`[MCP] POST /mcp error: ${error.stack || error.message}`);

      if (!res.headersSent) {
        return res.status(500).json({ error: error.message });
      }
    }
  });

  async function handleStreamableSessionRequest(req, res) {
    try {
      const sessionId = req.headers["mcp-session-id"];

      if (!sessionId || !streamableSessions.has(sessionId)) {
        return res.status(400).json({ error: "Invalid or missing MCP session ID" });
      }

      const transport = streamableSessions.get(sessionId);
      return transport.handleRequest(req, res);
    } catch (error) {
      log(`[MCP] Session request error: ${error.stack || error.message}`);

      if (!res.headersSent) {
        return res.status(500).json({ error: error.message });
      }
    }
  }

  app.get("/mcp", authMiddleware, handleStreamableSessionRequest);
  app.delete("/mcp", authMiddleware, handleStreamableSessionRequest);

  app.listen(PORT, "0.0.0.0", () => {
    log(`[HTTP] MCP server escutando na porta ${PORT}`);
    log(`[HTTP] SSE endpoint: http://0.0.0.0:${PORT}/sse`);
    log(
      `[HTTP] Auth: ${
        process.env.MCP_API_KEY || process.env.MCP_BEARER_TOKEN
          ? "Bearer token ativo"
          : "aberto (dev)"
      }`
    );
  });
} else {
  log("Iniciando MCP stdio server (v3.1.0)...");
  const transport = new StdioServerTransport();
  const mcpServer = createMcpServer();
  await mcpServer.connect(transport);
  log("MCP stdio server conectado.");
}
