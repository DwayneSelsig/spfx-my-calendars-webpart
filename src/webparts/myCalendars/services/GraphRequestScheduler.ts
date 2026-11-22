export const DEFAULT_GRAPH_CONCURRENCY = 10;
export const MIN_GRAPH_CONCURRENCY = 2;
export const MAX_GRAPH_CONCURRENCY = 20;

export function normalizeGraphConcurrency(value: unknown): number {
  return typeof value === 'number' && Number.isInteger(value)
    ? Math.max(MIN_GRAPH_CONCURRENCY, Math.min(MAX_GRAPH_CONCURRENCY, value))
    : DEFAULT_GRAPH_CONCURRENCY;
}

export class GraphRequestCancelled extends Error {
  public constructor() {
    super('Graph request was cancelled before execution.');
    Object.setPrototypeOf(this, GraphRequestCancelled.prototype);
  }
}

export class GraphRequestScope {
  private readonly listeners = new Set<() => void>();
  private priority = false;
  private cancelled = false;

  public get high(): boolean { return this.priority; }
  public get aborted(): boolean { return this.cancelled; }

  public setHigh(high: boolean): void {
    if (this.cancelled || this.priority === high) return;
    this.priority = high;
    this.listeners.forEach(listener => listener());
  }

  public cancel(): void {
    if (this.cancelled) return;
    this.cancelled = true;
    this.listeners.forEach(listener => listener());
  }

  public subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }
}

interface QueuedRequest {
  high: boolean;
  start: (high: boolean) => void;
  cancel: () => void;
  unsubscribe?: () => void;
}

export class GraphRequestScheduler {
  private readonly normal: QueuedRequest[] = [];
  private readonly high: QueuedRequest[] = [];
  private activeNormal = 0;
  private activeHigh = 0;
  private limit: number;
  private draining = false;

  public constructor(limit?: number) { this.limit = normalizeGraphConcurrency(limit); }

  public get counts(): { limit: number; activeNormal: number; activeHigh: number; queuedNormal: number; queuedHigh: number } {
    return { limit: this.limit, activeNormal: this.activeNormal, activeHigh: this.activeHigh,
      queuedNormal: this.normal.length, queuedHigh: this.high.length };
  }

  public setLimit(limit: number): void {
    this.limit = normalizeGraphConcurrency(limit);
    this.drain();
  }

  public run<T>(operation: () => Promise<T> | T, priority: boolean | GraphRequestScope = false): Promise<T> {
    const scope = typeof priority === 'boolean' ? undefined : priority;
    return new Promise<T>((resolve, reject) => {
      if (scope?.aborted) { reject(new GraphRequestCancelled()); return; }
      const request: QueuedRequest = {
        high: scope ? scope.high : priority === true,
        start: high => {
          const execute = async (): Promise<void> => {
            try { resolve(await operation()); }
            catch (error) { reject(error); }
            finally {
              if (high) this.activeHigh--;
              else this.activeNormal--;
              this.drain();
            }
          };
          execute().catch(reject);
        },
        cancel: () => reject(new GraphRequestCancelled())
      };
      this.queue(request.high).push(request);
      if (scope) request.unsubscribe = scope.subscribe(() => {
        const previousQueue = this.queue(request.high);
        const index = previousQueue.indexOf(request);
        if (index < 0) return;
        previousQueue.splice(index, 1);
        if (scope.aborted) {
          request.unsubscribe?.();
          request.cancel();
        } else {
          request.high = scope.high;
          this.queue(request.high).push(request);
        }
        this.drain();
      });
      this.drain();
    });
  }

  private queue(high: boolean): QueuedRequest[] { return high ? this.high : this.normal; }

  private next(): QueuedRequest | undefined {
    const total = this.activeNormal + this.activeHigh;
    if (total > this.limit) return undefined;
    if (total === this.limit) {
      return this.limit < MAX_GRAPH_CONCURRENCY && this.activeHigh === 0 ? this.high.shift() : undefined;
    }
    if (this.high.length && this.activeHigh === 0) return this.high.shift();
    const minNormal = this.limit > 10 ? 2 : 1;
    if (this.normal.length && this.activeNormal < minNormal) return this.normal.shift();
    if (this.high.length) return this.high.shift();
    if (this.limit === MAX_GRAPH_CONCURRENCY && this.activeNormal === MAX_GRAPH_CONCURRENCY - 1 && this.activeHigh === 0) {
      return undefined;
    }
    return this.normal.shift();
  }

  private drain(): void {
    if (this.draining) return;
    this.draining = true;
    try {
      let request = this.next();
      while (request) {
        request.unsubscribe?.();
        if (request.high) this.activeHigh++;
        else this.activeNormal++;
        request.start(request.high);
        request = this.next();
      }
    } finally { this.draining = false; }
  }
}
