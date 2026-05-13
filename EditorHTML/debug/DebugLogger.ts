export type HweDebugEvent = {
  data?: unknown;
  durationMs?: number;
  name: string;
  time: number;
};

const DEBUG_STORAGE_KEY = "hwe.debug";
const LAST_REPORT_STORAGE_KEY = "hwe.debug.lastReport";
const MAX_EVENTS = 500;
const HEARTBEAT_INTERVAL_MS = 1000;
const MAIN_THREAD_STALL_MS = 3000;

const events: HweDebugEvent[] = [];
let installed = false;

type HweDebugWindow = Window & {
  __HWE_DEBUG__?: boolean;
  __hweDebug?: {
    clear: () => void;
    enabled: boolean;
    events: HweDebugEvent[];
    lastReport: () => unknown;
    report: () => unknown;
  };
};

export function isHweDebugEnabled(): boolean {
  const debugWindow = window as HweDebugWindow;
  if (debugWindow.__HWE_DEBUG__ === true) return true;

  try {
    const params = new URLSearchParams(window.location.search);
    return (
      window.localStorage.getItem(DEBUG_STORAGE_KEY) === "1" ||
      params.get("hweDebug") === "1" ||
      params.get("hwe-debug") === "1"
    );
  } catch {
    return false;
  }
}

export function hweDebugLog(name: string, data?: unknown): void {
  const event: HweDebugEvent = {
    name,
    time: Math.round(performance.now() * 100) / 100,
  };
  if (data !== undefined) event.data = data;
  events.push(event);
  if (events.length > MAX_EVENTS) events.splice(0, events.length - MAX_EVENTS);
  if (isHweDebugEnabled()) console.debug("[HtmlWordEditor]", name, data ?? "");
}

export function hweDebugStart(name: string, data?: unknown): (endData?: unknown) => void {
  const start = performance.now();
  hweDebugLog(`${name}:start`, data);
  return (endData?: unknown) => {
    const durationMs = Math.round((performance.now() - start) * 100) / 100;
    hweDebugLog(`${name}:end`, { ...(asRecord(endData) ?? {}), durationMs });
  };
}

export function installHweDebugGlobals(rootProvider: () => HTMLElement | null): void {
  if (installed) return;
  installed = true;

  const debugWindow = window as HweDebugWindow;
  debugWindow.__hweDebug = {
    clear: () => {
      events.length = 0;
    },
    get enabled() {
      return isHweDebugEnabled();
    },
    events,
    lastReport: () => readLastReport(),
    report: () => saveHweDebugReport(rootProvider()),
  };
}

export function getHweDebugReport(root: HTMLElement | null): unknown {
  return buildReport(root);
}

export function saveHweDebugReport(root: HTMLElement | null): unknown {
  const report = buildReport(root);
  try {
    window.localStorage.setItem(LAST_REPORT_STORAGE_KEY, JSON.stringify(report));
  } catch {
    // Ignore storage quota or privacy-mode failures; the visible copy button still works.
  }
  return report;
}

export function startHweDebugHeartbeat(rootProvider: () => HTMLElement | null): () => void {
  let lastBeat = performance.now();
  const intervalId = window.setInterval(() => {
    const now = performance.now();
    const drift = now - lastBeat - HEARTBEAT_INTERVAL_MS;
    lastBeat = now;
    if (drift <= MAIN_THREAD_STALL_MS) return;

    hweDebugLog("mainThreadStall", {
      driftMs: Math.round(drift * 100) / 100,
      thresholdMs: MAIN_THREAD_STALL_MS,
    });
    saveHweDebugReport(rootProvider());
  }, HEARTBEAT_INTERVAL_MS);

  return () => window.clearInterval(intervalId);
}

function buildReport(root: HTMLElement | null): unknown {
  const scope = root ?? document.body;
  const pages = Array.from(scope.querySelectorAll<HTMLElement>(".hwe-page"));
  const images = Array.from(scope.querySelectorAll<HTMLImageElement>("img"));
  const tables = Array.from(scope.querySelectorAll<HTMLTableElement>("table"));

  return {
    enabled: isHweDebugEnabled(),
    eventCount: events.length,
    generatedAt: new Date().toISOString(),
    viewport: {
      height: window.innerHeight,
      width: window.innerWidth,
    },
    root: root ? getRect(root) : null,
    workspace: getRect(scope.querySelector<HTMLElement>(".hwe-workspace")),
    pages: pages.map((page, index) => getPageReport(page, index)),
    images: images.map((image, index) => getImageReport(image, index)),
    ocrStates: countBy(
      images.map((image) => image.getAttribute("data-hwe-ocr-state") ?? "none")
    ),
    tables: tables.map((table, index) => getTableReport(table, index)),
    recentEvents: events.slice(-150),
  };
}

function getPageReport(page: HTMLElement, index: number): unknown {
  const inner = page.querySelector<HTMLElement>(".hwe-page-inner");
  const children = inner ? getMeaningfulChildElements(inner) : [];
  const contentBottom = children.reduce<number | null>((bottom, child) => {
    const rect = child.getBoundingClientRect();
    return bottom === null ? rect.bottom : Math.max(bottom, rect.bottom);
  }, null);
  const limitBottom = inner ? getContentLimitBottom(inner) : null;
  const scrollOverflows = inner ? inner.scrollHeight > inner.clientHeight + 1 : false;

  return {
    index,
    childCount: children.length,
    contentBottom,
    firstChild: describeElement(children[0]),
    inner: getRect(inner),
    lastChild: describeElement(children[children.length - 1]),
    limitBottom,
    scroll: inner
      ? {
          clientHeight: inner.clientHeight,
          scrollHeight: inner.scrollHeight,
        }
      : null,
    overflows:
      contentBottom !== null && limitBottom !== null
        ? contentBottom > limitBottom + 1 || scrollOverflows
        : scrollOverflows,
    page: getRect(page),
    userBlankCount: inner?.querySelectorAll("[data-hwe-user-blank='true']").length ?? 0,
  };
}

function getImageReport(image: HTMLImageElement, index: number): unknown {
  const src = image.getAttribute("src") ?? "";
  return {
    index,
    alt: image.getAttribute("alt") ?? "",
    complete: image.complete,
    dataUrlBytes: src.startsWith("data:image/") ? estimateDataUrlBytes(src) : 0,
    declaredHeight: image.getAttribute("height") ?? "",
    declaredWidth: image.getAttribute("width") ?? "",
    hasDetachedOriginal: image.hasAttribute("data-hwe-large-image-id"),
    isDataUrl: src.startsWith("data:image/"),
    isPlaceholder: image.getAttribute("data-hwe-large-image-placeholder") === "true",
    naturalHeight: image.naturalHeight,
    naturalWidth: image.naturalWidth,
    ocrState: image.getAttribute("data-hwe-ocr-state") ?? "",
    rect: getRect(image),
  };
}

function getTableReport(table: HTMLTableElement, index: number): unknown {
  const invalidBlockAncestor = table.closest("p, li, span");
  const cellAncestor = table.closest("td, th");
  return {
    index,
    className: table.getAttribute("class") ?? "",
    flowId: table.getAttribute("data-hwe-table-flow-id") ?? "",
    fragment: table.getAttribute("data-hwe-table-fragment") === "true",
    invalidBlockAncestor:
      invalidBlockAncestor && !cellAncestor
        ? {
            className: invalidBlockAncestor.getAttribute("class") ?? "",
            tagName: invalidBlockAncestor.tagName,
            text: (invalidBlockAncestor.textContent ?? "").replace(/\s+/g, " ").trim().slice(0, 80),
          }
        : null,
    rect: getRect(table),
    rows: table.querySelectorAll("tr").length,
  };
}

function getContentLimitBottom(inner: HTMLElement): number {
  const styles = getComputedStyle(inner);
  const paddingBottom = parseFloat(styles.paddingBottom) || 0;
  return inner.getBoundingClientRect().bottom - Math.min(paddingBottom, 12);
}

function getMeaningfulChildElements(inner: HTMLElement): HTMLElement[] {
  return Array.from(inner.children).filter((child) => {
    if (child.hasAttribute("data-hwe-user-blank")) return true;
    if (child.matches("style, meta, link, script")) return false;
    if (child.querySelector("img, table, tr, td, th, video, canvas, svg")) return true;
    return (child.textContent ?? "").replace(/\u00a0/g, " ").trim().length > 0;
  }) as HTMLElement[];
}

function getRect(element: HTMLElement | null): unknown {
  if (!element) return null;
  const rect = element.getBoundingClientRect();
  return {
    bottom: Math.round(rect.bottom * 100) / 100,
    height: Math.round(rect.height * 100) / 100,
    left: Math.round(rect.left * 100) / 100,
    top: Math.round(rect.top * 100) / 100,
    width: Math.round(rect.width * 100) / 100,
  };
}

function describeElement(element: Element | undefined): unknown {
  if (!element) return null;
  return {
    blank: element.getAttribute("data-hwe-user-blank") === "true",
    className: element.getAttribute("class") ?? "",
    tagName: element.tagName,
    text: (element.textContent ?? "").replace(/\s+/g, " ").trim().slice(0, 80),
  };
}

function estimateDataUrlBytes(src: string): number {
  const commaIndex = src.indexOf(",");
  const payload = commaIndex >= 0 ? src.slice(commaIndex + 1) : src;
  if (/^data:[^;,]+;base64,/i.test(src)) {
    return Math.floor(payload.length * 0.75);
  }

  return payload.length;
}

function countBy(values: string[]): Record<string, number> {
  return values.reduce<Record<string, number>>((acc, value) => {
    acc[value] = (acc[value] ?? 0) + 1;
    return acc;
  }, {});
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function readLastReport(): unknown {
  try {
    const stored = window.localStorage.getItem(LAST_REPORT_STORAGE_KEY);
    return stored ? JSON.parse(stored) : null;
  } catch {
    return null;
  }
}