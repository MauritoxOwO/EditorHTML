interface DynamicHeaderResponse {
  responseObjects?: Array<{
    responseBody?: string;
  }>;
}

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

  const payload = (await response.json()) as DynamicHeaderResponse;
  return payload.responseObjects?.[1]?.responseBody?.trim() ?? "";
}

function buildDynamicHeaderUrl(baseUrl: string, endpointUrl: string): string {
  const endpoint = endpointUrl.trim() || "/api/data/v9.2/ays_GenerarCabezaAnuncio";
  if (/^https?:\/\//i.test(endpoint)) return endpoint;
  if (endpoint.startsWith("/")) return `${baseUrl.replace(/\/$/, "")}${endpoint}`;
  return `${baseUrl.replace(/\/$/, "")}/${endpoint}`;
}
