export class EditorHistoryController {
  private readonly maxStates: number;
  private snapshots: string[] = [];
  private index = -1;
  private suspended = false;

  constructor(maxStates = 10) {
    this.maxStates = Math.max(2, maxStates);
  }

  get canUndo(): boolean {
    return this.index > 0;
  }

  get canRedo(): boolean {
    return this.index >= 0 && this.index < this.snapshots.length - 1;
  }

  reset(snapshot: string): void {
    this.snapshots = [snapshot];
    this.index = 0;
  }

  record(snapshot: string): boolean {
    if (this.suspended || !snapshot) return false;

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

  undo(): string | null {
    if (!this.canUndo) return null;
    this.index -= 1;
    return this.snapshots[this.index] ?? null;
  }

  redo(): string | null {
    if (!this.canRedo) return null;
    this.index += 1;
    return this.snapshots[this.index] ?? null;
  }

  async runSuspended<T>(callback: () => T | Promise<T>): Promise<T> {
    this.suspended = true;
    try {
      return await callback();
    } finally {
      this.suspended = false;
    }
  }
}
