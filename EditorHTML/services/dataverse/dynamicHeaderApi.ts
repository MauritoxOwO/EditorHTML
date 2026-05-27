interface DynamicHeaderResponse {
  cabeceraHTML?: string;
}

const DYNAMIC_HEADER_ENDPOINT = "/api/data/v9.2/ays_GenerarCabeceraAnuncio";

export async function fetchDynamicDocumentHeaderHtml(
  baseUrl: string,
  entityId: string
): Promise<string> {
  if (!baseUrl || !entityId) return "";

  const url = `${baseUrl.replace(/\/$/, "")}${DYNAMIC_HEADER_ENDPOINT}`;
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
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
  return payload.cabeceraHTML?.trim() ?? "";
}
