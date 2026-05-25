export interface DynamicDocumentHeaderConfig {
  entitySetName: string;
  titleField?: string;
  subtitleField?: string;
  subtitle2Field?: string;
}

export interface DynamicDocumentHeaderValues {
  title: string;
  subtitle: string;
  subtitle2: string;
}

export async function fetchDynamicDocumentHeader(
  baseUrl: string,
  entityId: string,
  config: DynamicDocumentHeaderConfig
): Promise<DynamicDocumentHeaderValues> {
  const fieldMap = [
    { key: "title" as const, field: config.titleField },
    { key: "subtitle" as const, field: config.subtitleField },
    { key: "subtitle2" as const, field: config.subtitle2Field },
  ].filter((entry): entry is { key: keyof DynamicDocumentHeaderValues; field: string } =>
    Boolean(entry.field)
  );

  if (!baseUrl || !entityId || !config.entitySetName || fieldMap.length === 0) {
    return makeEmptyHeaderValues();
  }

  const select = Array.from(new Set(fieldMap.map((entry) => entry.field))).join(",");
  const url =
    `${baseUrl}/api/data/v9.2/${config.entitySetName}(${entityId})` +
    `?$select=${encodeURIComponent(select)}`;

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
    throw new Error(`No se pudo cargar la cabecera dinamica (HTTP ${response.status}). ${body}`);
  }

  const row = (await response.json()) as Record<string, unknown>;
  const values = makeEmptyHeaderValues();
  fieldMap.forEach((entry) => {
    values[entry.key] = String(row[entry.field] ?? "").trim();
  });
  return values;
}

function makeEmptyHeaderValues(): DynamicDocumentHeaderValues {
  return {
    title: "",
    subtitle: "",
    subtitle2: "",
  };
}
