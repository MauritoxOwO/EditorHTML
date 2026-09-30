export interface EditorHistoryControllerOptions<ViewState = undefined> {
  maxStates?: number;
  recordDelayMs?: number;
  collectSnapshot: () => string;
  collectViewState?: () => ViewState;
  restoreSnapshot: (snapshot: string, viewState?: ViewState) => Promise<void> | void;
  onRestored?: () => void;
  onAvailabilityChanged?: (canUndo: boolean, canRedo: boolean) => void;
}

interface HistoryEntry<ViewState> {
  html: string;
  viewState?: ViewState;
}

export class EditorHistoryController<ViewState = undefined> {
  private readonly maxStates: number;
  private readonly recordDelayMs: number;
  private snapshots: HistoryEntry<ViewState>[] = [];
  private index = -1;
  private recordTimer: number | undefined;
  private isRestoring = false;

  constructor(private readonly options: EditorHistoryControllerOptions<ViewState>) {
    this.maxStates = Math.max(2, options.maxStates ?? 10);
    this.recordDelayMs = options.recordDelayMs ?? 650;
  }

  get canUndo(): boolean {
    return this.index > 0;
  }

  get canRedo(): boolean {
    return this.index >= 0 && this.index < this.snapshots.length - 1;
  }

  handleShortcut(event: KeyboardEvent): boolean {
    if (event.altKey) return false;

    const key = event.key.toLowerCase();
    const isUndo = (event.ctrlKey || event.metaKey) && key === "z" && !event.shiftKey;
    const isRedo = event.ctrlKey && !event.metaKey && key === "y" && !event.shiftKey;
    if (!isUndo && !isRedo) return false;

    event.preventDefault();
    event.stopPropagation();
    if (isRedo) {
      this.redo();
    } else {
      this.undo();
    }
    return true;
  }

  reset(): void {
    this.clearPendingRecord();
    const snapshot = this.options.collectSnapshot();
    this.snapshots = [{ html: snapshot, viewState: this.options.collectViewState?.() }];
    this.index = 0;
    this.notifyAvailabilityChanged();
  }

  scheduleRecord(): void {
    if (this.isRestoring) return;
    this.clearPendingRecord();
    this.recordTimer = window.setTimeout(() => {
      this.recordTimer = undefined;
      this.recordNow();
    }, this.recordDelayMs);
  }

  // Recordar el inicio de una edición sin crear otro paso ni romper la agrupación al escribir.
  rememberViewState(): void {
    if (this.isRestoring || this.recordTimer !== undefined) return;
    const current = this.snapshots[this.index];
    if (current) current.viewState = this.options.collectViewState?.();
  }

  flushPendingRecord(): void {
    if (this.recordTimer === undefined) return;
    this.clearPendingRecord();
    this.recordNow();
  }

  recordNow(): boolean {
    if (this.isRestoring) return false;
    this.clearPendingRecord();
    return this.record(this.options.collectSnapshot());
  }

  undo(): void {
    if (this.isRestoring) return;

    this.flushPendingRecord();
    const snapshot = this.popUndoSnapshot();
    if (snapshot) void this.restore(snapshot);
  }

  redo(): void {
    if (this.isRestoring) return;

    this.flushPendingRecord();
    if (!this.canRedo) return;

    this.index += 1;
    this.notifyAvailabilityChanged();
    const snapshot = this.snapshots[this.index];
    if (snapshot) void this.restore(snapshot);
  }

  destroy(): void {
    this.clearPendingRecord();
  }

  private record(snapshot: string): boolean {
    if (!snapshot) return false;

    const viewState = this.options.collectViewState?.();
    if (this.snapshots[this.index]?.html === snapshot) {
      this.snapshots[this.index].viewState = viewState;
      return false;
    }

    if (this.index < this.snapshots.length - 1) {
      this.snapshots = this.snapshots.slice(0, this.index + 1);
    }

    this.snapshots.push({ html: snapshot, viewState });
    if (this.snapshots.length > this.maxStates) {
      this.snapshots.shift();
    }
    this.index = this.snapshots.length - 1;
    this.notifyAvailabilityChanged();
    return true;
  }

  private popUndoSnapshot(): HistoryEntry<ViewState> | null {
    if (!this.canUndo) return null;
    this.index -= 1;
    this.notifyAvailabilityChanged();
    return this.snapshots[this.index] ?? null;
  }

  private async restore(snapshot: HistoryEntry<ViewState>): Promise<void> {
    if (this.isRestoring) return;

    this.clearPendingRecord();
    this.isRestoring = true;
    try {
      await this.options.restoreSnapshot(snapshot.html, snapshot.viewState);
      this.options.onRestored?.();
    } finally {
      this.isRestoring = false;
    }
  }

  private clearPendingRecord(): void {
    if (this.recordTimer === undefined) return;
    window.clearTimeout(this.recordTimer);
    this.recordTimer = undefined;
  }

  private notifyAvailabilityChanged(): void {
    this.options.onAvailabilityChanged?.(this.canUndo, this.canRedo);
  }
}
