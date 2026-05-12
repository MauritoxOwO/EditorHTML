export interface ParagraphStyleDefinition {
  label: string;
  className: string;
  cssText: string;
}

export interface ParagraphStyleTableConfig {
  entitySetName: string;
  labelField: string;
  classField: string;
  cssField: string;
}

export async function fetchParagraphStyles(
  baseUrl: string,
  config: ParagraphStyleTableConfig
): Promise<ParagraphStyleDefinition[]> {
  const select = [config.labelField, config.classField, config.cssField].join(",");
  const url =
    `${baseUrl}/api/data/v9.2/${config.entitySetName}` +
    `?$select=${encodeURIComponent(select)}` +
    `&$orderby=${encodeURIComponent(config.labelField)} asc`;

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
    throw new Error(`No se pudieron cargar estilos (HTTP ${response.status}). ${body}`);
  }

  const payload = (await response.json()) as { value?: Array<Record<string, unknown>> };
  return (payload.value ?? [])
    .map((row) => ({
      label: String(row[config.labelField] ?? row[config.classField] ?? "").trim(),
      className: String(row[config.classField] ?? "").trim().replace(/^\./, ""),
      cssText: String(row[config.cssField] ?? "").trim(),
    }))
    .filter((style) => style.label && style.className);
}
