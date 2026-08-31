import { describe, expect, it } from "vitest";
import { TraceRequestScheduler } from "../src/services/traceScheduler";

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((res) => {
    resolve = res;
  });
  return { promise, resolve };
}

describe("TraceRequestScheduler", () => {
  it("runs only one task at a time and prioritizes expansion over prefetch", async () => {
    const scheduler = new TraceRequestScheduler();
    const blocker = deferred<void>();
    const order: string[] = [];

    const initial = scheduler.enqueue("initial", "initial", async () => {
      order.push("initial-start");
      await blocker.promise;
      order.push("initial-end");
    });
    const prefetch = scheduler.enqueue("prefetch", "p1", async () => {
      order.push("prefetch");
    });
    const expansion = scheduler.enqueue("expansion", "e1", async () => {
      order.push("expansion");
    });

    blocker.resolve();
    await Promise.all([initial, prefetch, expansion]);
    expect(order).toEqual(["initial-start", "initial-end", "expansion", "prefetch"]);
  });

  it("coalesces queued navigation so only the newest address runs", async () => {
    const scheduler = new TraceRequestScheduler();
    const blocker = deferred<void>();
    const navigated: string[] = [];
    const running = scheduler.enqueue("expansion", "block", async () => {
      await blocker.promise;
    });

    const pending = Array.from({ length: 30 }, (_, index) => {
      const address = `A${index + 1}`;
      return scheduler.enqueueNavigation(address, async () => {
        navigated.push(address);
      });
    });
    blocker.resolve();

    await Promise.all([running, ...pending]);
    expect(navigated).toEqual(["A30"]);
  });

  it("cancels queued work and ignores the running result", async () => {
    const scheduler = new TraceRequestScheduler();
    const blocker = deferred<number>();
    const running = scheduler.enqueue("initial", "initial", async () => blocker.promise);
    const queued = scheduler.enqueue("prefetch", "prefetch", async () => 2);
    scheduler.cancelAll();
    blocker.resolve(1);

    await expect(running).resolves.toBeUndefined();
    await expect(queued).resolves.toBeUndefined();
  });
});
