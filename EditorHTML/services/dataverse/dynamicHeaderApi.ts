export async function fetchDynamicDocumentHeaderHtml(
  endpointUrl: string,
  entityId: string
): Promise<string> {
  if (!endpointUrl || !entityId) return "";

  const url = buildDynamicHeaderUrl(endpointUrl, entityId);
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "text/html, application/json, */*",
    },
    credentials: "same-origin",
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`No se pudo cargar la cabecera dinamica (HTTP ${response.status}). ${body}`);
  }

  const contentType = response.headers.get("content-type") ?? "";
  if (/application\/json/i.test(contentType)) {
    const payload = (await response.json()) as { html?: unknown; value?: unknown };
    return String(payload.html ?? payload.value ?? "").trim();
  }

  return (await response.text()).trim();
}

function buildDynamicHeaderUrl(endpointUrl: string, entityId: string): string {
  const encodedId = encodeURIComponent(entityId);
  if (endpointUrl.includes("{id}") || endpointUrl.includes("{guid}")) {
    return endpointUrl.replace(/\{(?:id|guid)\}/g, encodedId);
  }

  const separator = endpointUrl.includes("?") ? "&" : "?";
  return `${endpointUrl}${separator}id=${encodedId}`;
}
