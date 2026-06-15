export interface EditorHistoryControllerOptions {
  maxStates?: number;
  recordDelayMs?: number;
  collectSnapshot: () => string;
  restoreSnapshot: (snapshot: string) => Promise<void> | void;
  onRestored?: () => void;
}

export class EditorHistoryController {
  private readonly maxStates: number;
  private readonly recordDelayMs: number;
  private snapshots: string[] = [];
  private index = -1;
  private recordTimer: number | undefined;
  private isRestoring = false;

  constructor(private readonly options: EditorHistoryControllerOptions) {
    this.maxStates = Math.max(2, options.maxStates ?? 10);
    this.recordDelayMs = options.recordDelayMs ?? 650;
  }

  get canUndo(): boolean {
    return this.index > 0;
  }

  handleShortcut(event: KeyboardEvent): boolean {
    const isModifierPressed = event.ctrlKey || event.metaKey;
    if (!isModifierPressed || event.altKey) return false;

    const isUndo = event.key.toLowerCase() === "z" && !event.shiftKey;
    if (!isUndo) return false;

    event.preventDefault();
    event.stopPropagation();
    this.undo();
    return true;
  }

  reset(): void {
    this.clearPendingRecord();
    const snapshot = this.options.collectSnapshot();
    this.snapshots = [snapshot];
    this.index = 0;
  }

  scheduleRecord(): void {
    if (this.isRestoring) return;
    this.clearPendingRecord();
    this.recordTimer = window.setTimeout(() => {
      this.recordTimer = undefined;
      this.recordNow();
    }, this.recordDelayMs);
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

  destroy(): void {
    this.clearPendingRecord();
  }

  private record(snapshot: string): boolean {
    if (!snapshot) return false;

    if (this.snapshots[this.index] === snapshot) {
      return false;
    }

    if (this.index < this.snapshots.length - 1) {
      this.snapshots = this.snapshots.slice(0, this.index + 1);
    }

    this.snapshots.push(snapshot);
    if (this.snapshots.length > this.maxStates) {
      this.snapshots.shift();
    }
    this.index = this.snapshots.length - 1;
    return true;
  }

  private popUndoSnapshot(): string | null {
    if (!this.canUndo) return null;
    this.index -= 1;
    return this.snapshots[this.index] ?? null;
  }

  private async restore(snapshot: string): Promise<void> {
    if (this.isRestoring) return;

    this.clearPendingRecord();
    this.isRestoring = true;
    try {
      await this.options.restoreSnapshot(snapshot);
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
}
