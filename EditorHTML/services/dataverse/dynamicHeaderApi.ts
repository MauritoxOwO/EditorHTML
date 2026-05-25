export async function fetchDynamicDocumentHeaderHtml(
  baseUrl: string,
  endpointUrl: string,
  entityId: string
): Promise<string> {
  if (!baseUrl || !entityId) return "";

  const url = buildDynamicHeaderUrl(baseUrl, endpointUrl);
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "text/html, application/json, */*",
      "Content-Type": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
    credentials: "same-origin",
    body: JSON.stringify({
      guidVersionAnuncio: entityId,
    }),
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`No se pudo cargar la cabecera dinamica (HTTP ${response.status}). ${body}`);
  }

  const contentType = response.headers.get("content-type") ?? "";
  if (/application\/json/i.test(contentType)) {
    return extractHtmlFromJson(await response.json());
  }

  return (await response.text()).trim();
}

function buildDynamicHeaderUrl(baseUrl: string, endpointUrl: string): string {
  const endpoint = endpointUrl.trim() || "/api/data/v9.2/ays_GenerarCabezaAnuncio";
  if (/^https?:\/\//i.test(endpoint)) return endpoint;
  if (endpoint.startsWith("/")) return `${baseUrl.replace(/\/$/, "")}${endpoint}`;
  return `${baseUrl.replace(/\/$/, "")}/${endpoint}`;
}

function extractHtmlFromJson(payload: unknown): string {
  if (typeof payload === "string") return payload.trim();
  if (!payload || typeof payload !== "object") return "";

  const row = payload as Record<string, unknown>;
  const responseObjectHtml = extractResponseObjectHtml(row.responseObjects);
  if (responseObjectHtml) return responseObjectHtml;

  const direct = row.html ?? row.Html ?? row.HTML ?? row.value ?? row.Value;
  if (direct !== undefined) return String(direct).trim();

  const firstStringValue = Object.values(row).find((value) => typeof value === "string");
  return firstStringValue ? String(firstStringValue).trim() : "";
}

function extractResponseObjectHtml(responseObjects: unknown): string {
  if (!Array.isArray(responseObjects)) return "";

  const preferred = extractResponseBody(responseObjects[1]);
  if (preferred) return preferred;

  return responseObjects.map(extractResponseBody).find(Boolean) ?? "";
}

function extractResponseBody(value: unknown): string {
  if (typeof value === "string") return value.trim();
  if (!value || typeof value !== "object") return "";

  const row = value as Record<string, unknown>;
  const body = row.responseBody ?? row.ResponseBody ?? row.value ?? row.Value;
  return body === undefined ? "" : String(body).trim();
}
