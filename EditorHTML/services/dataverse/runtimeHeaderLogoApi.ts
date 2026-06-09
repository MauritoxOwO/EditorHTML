export interface RuntimeHeaderLogoTableConfig {
  entitySetName?: string;
  imageField?: string;
  recordId?: string;
  idField?: string;
  nameField?: string;
  nameValue?: string;
}

export async function fetchRuntimeHeaderLogoSrc(
  baseUrl: string,
  config: RuntimeHeaderLogoTableConfig
): Promise<string> {
  if (!baseUrl || !config.entitySetName || !config.imageField) return "";

  const recordId =
    cleanGuid(config.recordId) ?? (await fetchRuntimeHeaderLogoRecordId(baseUrl, config));
  if (!recordId) return "";

  const url =
    `${baseUrl.replace(/\/$/, "")}/api/data/v9.2/` +
    `${config.entitySetName}(${recordId})/${config.imageField}/$value?size=full`;

  const blob = await fetchRuntimeHeaderLogoBlob(`${url}?size=full`).catch(() =>
    fetchRuntimeHeaderLogoBlob(url)
  );

  return blob.size > 0 ? URL.createObjectURL(blob) : "";
}

async function fetchRuntimeHeaderLogoBlob(url: string): Promise<Blob> {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "image/*, application/octet-stream, */*",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
    credentials: "same-origin",
  });

  if (response.status === 204 && url.endsWith("?size=full")) {
    return fetchRuntimeHeaderLogoBlob(url.slice(0, -"?size=full".length));
  }

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`No se pudo cargar el logo BOCM (HTTP ${response.status}). ${body}`);
  }

  return response.blob();
}

async function fetchRuntimeHeaderLogoRecordId(
  baseUrl: string,
  config: RuntimeHeaderLogoTableConfig
): Promise<string> {
  if (!config.entitySetName || !config.idField || !config.nameField || !config.nameValue) {
    return "";
  }

  const select = encodeURIComponent(`${config.idField},${config.nameField}`);
  const filter = encodeURIComponent(
    `${config.nameField} eq '${escapeODataString(config.nameValue)}'`
  );
  const orderBy = encodeURIComponent(config.nameField);
  const url =
    `${baseUrl.replace(/\/$/, "")}/api/data/v9.2/${config.entitySetName}` +
    `?$select=${select}&$filter=${filter}&$orderby=${orderBy} asc&$top=1`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
    credentials: "same-origin",
  });

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(`No se pudo localizar el logo BOCM (HTTP ${response.status}). ${body}`);
  }

  const payload = (await response.json()) as { value?: Array<Record<string, unknown>> };
  const row = payload.value?.[0];
  return cleanGuid(String(row?.[config.idField] ?? "")) ?? "";
}

function cleanGuid(value: string | undefined): string | null {
  const cleaned = value?.replace(/[{}]/g, "").trim();
  if (!cleaned) return null;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(cleaned)
    ? cleaned
    : null;
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
