import type {
  DirectNeighborResult,
  TraceCellData,
  TraceNodeData,
  TraceRangeData,
  TraceRangePageResult,
  TraceTarget,
} from "../core/traceGraph";
import { getRangePageBounds } from "../core/tracePolicy";
import type { TraceDirection } from "../core/traceUtils";
import type { TraceLoadRequest, TracePort } from "./tracePort";

export class TracePortFake implements TracePort {
  readonly directNeighborRequests: string[] = [];
  readonly rangePageRequests: Array<{ key: string; page: number }> = [];
  readonly selections: TraceTarget[] = [];

  private readonly neighbors = new Map<string, TraceNodeData[]>();
  private readonly rangeCells = new Map<string, TraceCellData[]>();
  private pausePromise: Promise<void> | null = null;
  private resumePause: (() => void) | null = null;

  constructor(private readonly root: TraceCellData) {}

  setNeighbors(key: string, nodes: TraceNodeData[]): void {
    this.neighbors.set(
      key,
      nodes.map((node) => ({ ...node }))
    );
  }

  setRangeCells(key: string, nodes: TraceCellData[]): void {
    this.rangeCells.set(
      key,
      nodes.map((node) => ({ ...node }))
    );
  }

  pauseDirectNeighborRequests(): void {
    if (this.pausePromise) return;
    this.pausePromise = new Promise<void>((resolve) => {
      this.resumePause = resolve;
    });
  }

  resumeDirectNeighborRequests(): void {
    this.resumePause?.();
    this.pausePromise = null;
    this.resumePause = null;
  }

  async getRoot(): Promise<TraceCellData> {
    return { ...this.root };
  }

  async getDirectNeighbors(
    node: TraceCellData,
    _direction: TraceDirection,
    request: TraceLoadRequest
  ): Promise<DirectNeighborResult> {
    this.directNeighborRequests.push(node.key);
    if (this.pausePromise) await this.pausePromise;
    const all = this.neighbors.get(node.key) ?? [];
    const nodes = all.slice(0, request.remainingNodes).map((item) => ({ ...item }));
    return {
      nodes,
      truncated: nodes.length < all.length,
      materializedCellCount: nodes.filter((item) => item.kind === "cell").length,
    };
  }

  async getRangePage(
    node: TraceRangeData,
    page: number,
    request: TraceLoadRequest
  ): Promise<TraceRangePageResult> {
    this.rangePageRequests.push({ key: node.key, page });
    const all = this.rangeCells.get(node.key) ?? [];
    const bounds = getRangePageBounds(all.length, page, request.policy);
    const count = Math.min(bounds.count, request.remainingNodes);
    return {
      nodes: all.slice(bounds.start, bounds.start + count).map((item) => ({ ...item })),
      page,
      hasMore: bounds.hasMore || count < bounds.count,
      truncated: count < bounds.count,
    };
  }

  async select(target: TraceTarget): Promise<void> {
    this.selections.push({ ...target });
  }
}
