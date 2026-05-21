export interface ParagraphStyleDefinition {
  label: string;
  className: string;
  cssText: string;
}

export interface ParagraphFontFaceDefinition {
  label: string;
  cssText: string;
}

export interface ParagraphStyleCatalog {
  styles: ParagraphStyleDefinition[];
  fonts: ParagraphFontFaceDefinition[];
}

export interface ParagraphStyleTableConfig {
  entitySetName: string;
  classField: string;
  cssField: string;
  typeField?: string;
  styleTypeValue?: string;
  fontTypeValue?: string;
}

type StyleCatalogRowKind = "style" | "font";

const EMPTY_STYLE_CATALOG: ParagraphStyleCatalog = {
  styles: [],
  fonts: [],
};

export async function fetchParagraphStyles(
  baseUrl: string,
  config: ParagraphStyleTableConfig
): Promise<ParagraphStyleDefinition[]> {
  return (await fetchParagraphStyleCatalog(baseUrl, config)).styles;
}

export async function fetchParagraphStyleCatalog(
  baseUrl: string,
  config: ParagraphStyleTableConfig
): Promise<ParagraphStyleCatalog> {
  const selectFields = [config.classField, config.cssField];
  if (config.typeField) selectFields.push(config.typeField);

  const select = Array.from(new Set(selectFields)).join(",");
  const url =
    `${baseUrl}/api/data/v9.2/${config.entitySetName}` +
    `?$select=${encodeURIComponent(select)}` +
    `&$orderby=${encodeURIComponent(config.classField)} asc`;

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
  return mapRowsToStyleCatalog(payload.value ?? [], config);
}

export function mapRowsToStyleCatalog(
  rows: Array<Record<string, unknown>>,
  config: ParagraphStyleTableConfig
): ParagraphStyleCatalog {
  if (rows.length === 0) return EMPTY_STYLE_CATALOG;

  const catalog: ParagraphStyleCatalog = {
    styles: [],
    fonts: [],
  };

  rows.forEach((row) => {
    const label = String(row[config.classField] ?? "").trim();
    const cssText = String(row[config.cssField] ?? "").trim();
    if (!label || !cssText) return;

    const kind = getStyleCatalogRowKind(row, config, cssText);
    if (kind === "font") {
      catalog.fonts.push({
        label,
        cssText,
      });
      return;
    }

    catalog.styles.push({
      label,
      className: label.replace(/^\./, ""),
      cssText,
    });
  });

  return catalog;
}

function getStyleCatalogRowKind(
  row: Record<string, unknown>,
  config: ParagraphStyleTableConfig,
  cssText: string
): StyleCatalogRowKind {
  const configuredKind = getConfiguredStyleCatalogRowKind(row, config);
  if (configuredKind) return configuredKind;

  return /@font-face\b/i.test(cssText) ? "font" : "style";
}

function getConfiguredStyleCatalogRowKind(
  row: Record<string, unknown>,
  config: ParagraphStyleTableConfig
): StyleCatalogRowKind | null {
  if (!config.typeField || (!config.styleTypeValue && !config.fontTypeValue)) {
    return null;
  }

  const rawValue = row[config.typeField];
  const value = rawValue === null || rawValue === undefined ? "" : String(rawValue).trim();
  if (!value) return null;

  if (config.fontTypeValue !== undefined && value === String(config.fontTypeValue).trim()) {
    return "font";
  }
  if (config.styleTypeValue !== undefined && value === String(config.styleTypeValue).trim()) {
    return "style";
  }

  return null;
}
