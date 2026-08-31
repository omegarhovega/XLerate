import type { TraceNode } from "./traceGraph";
import { MAX_TRACE_SAFETY_LIMIT, sanitizeTraceSafetyLimit } from "./traceUtils";

export type TraceLoadingPolicy = {
  rangeSummaryThreshold: number;
  rangePageSize: number;
  visibleNodeLimit: number;
  hardNodeLimit: number;
  prefetchMaxCandidates: number;
  prefetchMaxCells: number;
};

export const DEFAULT_TRACE_RANGE_SUMMARY_THRESHOLD = 50;
export const DEFAULT_TRACE_RANGE_PAGE_SIZE = 100;
export const DEFAULT_TRACE_VISIBLE_NODE_LIMIT = 500;
export const DEFAULT_TRACE_PREFETCH_MAX_CANDIDATES = 20;
export const DEFAULT_TRACE_PREFETCH_MAX_CELLS = 50;

export function buildTraceLoadingPolicy(
  input: Partial<TraceLoadingPolicy> & { safetyLimit?: number } = {}
): TraceLoadingPolicy {
  const hardNodeLimit = Math.min(
    MAX_TRACE_SAFETY_LIMIT,
    sanitizeTraceSafetyLimit(input.safetyLimit ?? input.hardNodeLimit)
  );

  return {
    rangeSummaryThreshold: input.rangeSummaryThreshold ?? DEFAULT_TRACE_RANGE_SUMMARY_THRESHOLD,
    rangePageSize: input.rangePageSize ?? DEFAULT_TRACE_RANGE_PAGE_SIZE,
    visibleNodeLimit: Math.min(
      input.visibleNodeLimit ?? DEFAULT_TRACE_VISIBLE_NODE_LIMIT,
      hardNodeLimit
    ),
    hardNodeLimit,
    prefetchMaxCandidates: input.prefetchMaxCandidates ?? DEFAULT_TRACE_PREFETCH_MAX_CANDIDATES,
    prefetchMaxCells: input.prefetchMaxCells ?? DEFAULT_TRACE_PREFETCH_MAX_CELLS,
  };
}

export function shouldSummarizeTraceRange(cellCount: number, policy: TraceLoadingPolicy): boolean {
  return cellCount > policy.rangeSummaryThreshold;
}

export function getRangePageBounds(
  cellCount: number,
  page: number,
  policy: TraceLoadingPolicy
): { start: number; count: number; hasMore: boolean } {
  const normalizedPage = Number.isInteger(page) && page >= 0 ? page : 0;
  const start = Math.min(cellCount, normalizedPage * policy.rangePageSize);
  const count = Math.min(policy.rangePageSize, Math.max(0, cellCount - start));
  return { start, count, hasMore: start + count < cellCount };
}

export function shouldPrefetchTraceNodes(nodes: TraceNode[], policy: TraceLoadingPolicy): boolean {
  if (nodes.length === 0 || nodes.length > policy.prefetchMaxCandidates) {
    return false;
  }

  let estimatedCells = 0;
  for (const node of nodes) {
    if (node.kind !== "cell" || node.loadState !== "unloaded") {
      return false;
    }
    estimatedCells += 1;
  }
  return estimatedCells <= policy.prefetchMaxCells;
}
