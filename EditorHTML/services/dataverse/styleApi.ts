export interface ParagraphStyleDefinition {
  label: string;
  className: string;
  cssText: string;
  showInDropdown?: boolean;
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
  stateField?: string;
  dropdownField?: string;
  documentTypeField?: string;
  documentTypeDropdownValue?: string;
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
  if (config.stateField) selectFields.push(config.stateField);
  if (config.dropdownField) selectFields.push(config.dropdownField);
  if (config.documentTypeField) selectFields.push(config.documentTypeField);
  if (config.typeField) selectFields.push(config.typeField);

  const select = Array.from(new Set(selectFields)).join(",");
  const filter = config.stateField
    ? `&$filter=${encodeURIComponent(`${config.stateField} eq 0`)}`
    : "";
  const url =
    `${baseUrl}/api/data/v9.2/${config.entitySetName}` +
    `?$select=${encodeURIComponent(select)}` +
    filter +
    `&$orderby=${encodeURIComponent(config.classField)} asc`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
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
    if (!isActiveRow(row, config)) return;

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
      showInDropdown: shouldShowInDropdown(row, config),
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
  if (!config.typeField) {
    return null;
  }

  const values = getChoiceValues(row, config.typeField);
  if (values.length === 0) return null;

  if (
    config.fontTypeValue !== undefined &&
    values.includes(normalizeChoiceValue(config.fontTypeValue))
  ) {
    return "font";
  }
  if (
    config.styleTypeValue !== undefined &&
    values.includes(normalizeChoiceValue(config.styleTypeValue))
  ) {
    return "style";
  }

  if (values.some(isFontChoiceLabel)) return "font";
  if (values.some(isStyleChoiceLabel)) return "style";

  return null;
}

function getChoiceValues(row: Record<string, unknown>, typeField: string): string[] {
  return [
    row[typeField],
    row[`${typeField}@OData.Community.Display.V1.FormattedValue`],
  ]
    .map(normalizeChoiceValue)
    .filter(Boolean);
}

function normalizeChoiceValue(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim().toLowerCase();
}

function isActiveRow(row: Record<string, unknown>, config: ParagraphStyleTableConfig): boolean {
  if (!config.stateField) return true;
  const state = row[config.stateField];
  return state === undefined || normalizeChoiceValue(state) === "0";
}

function shouldShowInDropdown(
  row: Record<string, unknown>,
  config: ParagraphStyleTableConfig
): boolean {
  const visible = config.dropdownField ? readBoolean(row[config.dropdownField]) ?? true : true;
  return visible && isDropdownDocumentType(row, config);
}

function isDropdownDocumentType(
  row: Record<string, unknown>,
  config: ParagraphStyleTableConfig
): boolean {
  if (!config.documentTypeField || !config.documentTypeDropdownValue) return true;
  return getChoiceValues(row, config.documentTypeField).includes(
    normalizeChoiceValue(config.documentTypeDropdownValue)
  );
}

function readBoolean(value: unknown): boolean | null {
  if (typeof value === "boolean") return value;
  if (value === null || value === undefined) return null;

  const normalized = normalizeChoiceValue(value);
  if (["true", "1", "yes", "si", "sí"].includes(normalized)) return true;
  if (["false", "0", "no"].includes(normalized)) return false;
  return null;
}

function isFontChoiceLabel(value: string): boolean {
  return ["font", "fonts", "font-face", "fontface", "fuente", "fuentes"].includes(value);
}

function isStyleChoiceLabel(value: string): boolean {
  return ["style", "styles", "css", "estilo", "estilos"].includes(value);
}
