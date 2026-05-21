export interface HtmlFileContent {
  html: string;
  fileName?: string;
}

export async function fetchHtmlFromFileField(
  baseUrl: string,
  entityName: string,
  entityId: string,
  fieldName: string
): Promise<HtmlFileContent> {
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
  return {
    html: new TextDecoder("utf-8").decode(buffer),
    fileName: getFileNameFromHeaders(response.headers),
  };
}

export async function saveHtmlToFileField(
  baseUrl: string,
  entityName: string,
  entityId: string,
  fieldName: string,
  htmlContent: string,
  fileName = "content.html"
): Promise<void> {
  const url = `${baseUrl}/api/data/v9.2/${entityName}(${entityId})/${fieldName}`;

  const blob = new Blob([htmlContent], { type: "text/html; charset=utf-8" });

  const response = await fetch(url, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/octet-stream",
      "x-ms-file-name": fileName,
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

function getFileNameFromHeaders(headers: Headers): string | undefined {
  const explicitName = headers.get("x-ms-file-name")?.trim();
  if (explicitName) return explicitName;

  const disposition = headers.get("content-disposition");
  if (!disposition) return undefined;

  const encodedMatch = /filename\*\s*=\s*(?:UTF-8'')?([^;]+)/i.exec(disposition);
  if (encodedMatch?.[1]) return decodeHeaderFileName(encodedMatch[1]);

  const plainMatch = /filename\s*=\s*("?)([^";]+)\1/i.exec(disposition);
  if (plainMatch?.[2]) return plainMatch[2].trim();

  return undefined;
}

function decodeHeaderFileName(value: string): string {
  const cleanValue = value.trim().replace(/^"|"$/g, "");
  try {
    return decodeURIComponent(cleanValue);
  } catch {
    return cleanValue;
  }
}
