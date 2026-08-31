import type {
  DirectNeighborResult,
  TraceCellData,
  TraceRangeData,
  TraceRangePageResult,
  TraceTarget,
} from "../core/traceGraph";
import type { TraceLoadingPolicy } from "../core/tracePolicy";
import type { TraceDirection } from "../core/traceUtils";

export type TraceLoadRequest = {
  remainingNodes: number;
  policy: TraceLoadingPolicy;
};

export interface TracePort {
  getRoot(): Promise<TraceCellData>;
  getDirectNeighbors(
    node: TraceCellData,
    direction: TraceDirection,
    request: TraceLoadRequest
  ): Promise<DirectNeighborResult>;
  getRangePage(
    node: TraceRangeData,
    page: number,
    request: TraceLoadRequest
  ): Promise<TraceRangePageResult>;
  select(target: TraceTarget): Promise<void>;
}
