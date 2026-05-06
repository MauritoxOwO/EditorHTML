export async function fetchHtmlFromFileField(
  baseUrl: string,
  entityName: string,
  entityId: string,
  fieldName: string
): Promise<string> {
  const url = `${baseUrl}/api/data/v9.2/${entityName}(${entityId})/${fieldName}/$value`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "text/html, application/octet-stream, */*"
    },
    credentials: "same-origin",
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    console.log(`Error in url: ${url}`);
    throw new Error(
      `No se pudo cargar el contenido (HTTP ${response.status}). ${body}`
    );

  }

  const buffer = await response.arrayBuffer();
  return new TextDecoder("utf-8").decode(buffer);
}

export async function saveHtmlToFileField(
  baseUrl: string,
  entityName: string,
  entityId: string,
  fieldName: string,
  htmlContent: string
): Promise<void> {
  console.log(baseUrl);
  const url = `${baseUrl}/api/data/v9.2/${entityName}(${entityId})/${fieldName}`;

  const blob = new Blob([htmlContent], { type: "text/html; charset=utf-8" });

  const response = await fetch(url, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/octet-stream",
      "x-ms-file-name": "content.html",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
    credentials: "same-origin",
    body: blob,
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(
      `No se pudo guardar (HTTP ${response.status}). ${body}`
    );
  }
}