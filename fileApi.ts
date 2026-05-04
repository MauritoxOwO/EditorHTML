/**
 * fileApi.ts
 * Operaciones OData contra campos de tipo File en Dataverse.
 */

/**
 * Lee el contenido HTML de un campo de archivo (file field) via OData $value.
 */
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
      Accept: "text/html, application/octet-stream, */*",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
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

  return response.text();
}

/**
 * Guarda el HTML en un campo de archivo (file field) via OData PATCH.
 *
 * Los file fields de Dataverse aceptan un PATCH directo con el binario
 * cuando el tamaño es < 128 MB. Para archivos pequeños (HTML) es suficiente.
 */
export async function saveHtmlToFileField(
  baseUrl: string,
  entityName: string,
  entityId: string,
  fieldName: string,
  htmlContent: string
): Promise<void> {
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
