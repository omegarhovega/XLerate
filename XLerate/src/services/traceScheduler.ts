export type TraceTaskKind = "navigation" | "expansion" | "initial" | "prefetch";

type QueuedTask = {
  kind: TraceTaskKind;
  key: string;
  generation: number;
  work: () => Promise<unknown>;
  resolve: (value: unknown | undefined) => void;
  reject: (reason: unknown) => void;
};

const PRIORITY: Record<TraceTaskKind, number> = {
  navigation: 0,
  expansion: 1,
  initial: 2,
  prefetch: 3,
};

export class TraceRequestScheduler {
  private queue: QueuedTask[] = [];
  private running = false;
  private generation = 0;

  enqueue<T>(kind: TraceTaskKind, key: string, work: () => Promise<T>): Promise<T | undefined> {
    return new Promise<T | undefined>((resolve, reject) => {
      this.queue.push({
        kind,
        key,
        generation: this.generation,
        work,
        resolve: (value) => resolve(value as T | undefined),
        reject,
      });
      this.sortQueue();
      void this.drain();
    });
  }

  enqueueNavigation<T>(key: string, work: () => Promise<T>): Promise<T | undefined> {
    const retained: QueuedTask[] = [];
    for (const task of this.queue) {
      if (task.kind === "navigation") task.resolve(undefined);
      else retained.push(task);
    }
    this.queue = retained;
    return this.enqueue("navigation", key, work);
  }

  cancelAll(): void {
    this.generation += 1;
    for (const task of this.queue) task.resolve(undefined);
    this.queue = [];
  }

  get queuedTaskCount(): number {
    return this.queue.length;
  }

  private sortQueue(): void {
    this.queue.sort((left, right) => PRIORITY[left.kind] - PRIORITY[right.kind]);
  }

  private async drain(): Promise<void> {
    if (this.running) return;
    this.running = true;
    try {
      while (this.queue.length > 0) {
        const task = this.queue.shift()!;
        if (task.generation !== this.generation) {
          task.resolve(undefined);
          continue;
        }
        try {
          const result = await task.work();
          task.resolve(task.generation === this.generation ? result : undefined);
        } catch (error) {
          if (task.generation === this.generation) task.reject(error);
          else task.resolve(undefined);
        }
      }
    } finally {
      this.running = false;
    }
  }
}
