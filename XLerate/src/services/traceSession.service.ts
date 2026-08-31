import type { TracePort } from "../adapters/tracePort";
import type {
  TraceCellData,
  TraceCellNode,
  TraceNode,
  TraceNodeData,
  TraceNodeLoadState,
  TraceRangeData,
  TraceRangeNode,
  TraceReferenceNode,
  TraceTarget,
} from "../core/traceGraph";
import type { TraceLoadingPolicy } from "../core/tracePolicy";
import { shouldPrefetchTraceNodes } from "../core/tracePolicy";
import type { TraceDirection } from "../core/traceUtils";

export type TraceExpansionResult = {
  parentId: string;
  parentState: TraceNodeLoadState;
  children: TraceNode[];
  truncated: boolean;
  cached: boolean;
  totalNodeCount: number;
};

export type TraceOpenResult = {
  root: TraceCellNode;
  children: TraceNode[];
  truncated: boolean;
  totalNodeCount: number;
};

export type TraceRangePageLoadResult = {
  rangeId: string;
  parentState: TraceNodeLoadState;
  page: number;
  children: TraceNode[];
  hasMore: boolean;
  truncated: boolean;
  totalNodeCount: number;
};

function cancelledError(): Error {
  return new Error("Trace session was cancelled");
}

export class TraceSessionService {
  private readonly nodesById = new Map<string, TraceNode>();
  private readonly canonicalIdByKey = new Map<string, string>();
  private readonly childCache = new Map<string, TraceExpansionResult>();
  private readonly pageCache = new Map<string, TraceRangePageLoadResult>();
  private readonly inFlight = new Map<string, Promise<TraceExpansionResult>>();
  private rootTarget: TraceTarget | null = null;
  private referenceSequence = 0;
  private cancelled = false;

  constructor(
    private readonly port: TracePort,
    private readonly direction: TraceDirection,
    readonly policy: TraceLoadingPolicy,
    private readonly maxDepth = 20
  ) {}

  cancel(): void {
    this.cancelled = true;
  }

  isCancelled(): boolean {
    return this.cancelled;
  }

  getNode(nodeId: string): TraceNode | undefined {
    return this.nodesById.get(nodeId);
  }

  getPrefetchCandidates(nodes: TraceNode[]): TraceCellNode[] {
    const candidates = nodes.filter(
      (node): node is TraceCellNode =>
        node.kind === "cell" && node.loadState === "unloaded" && node.level < this.maxDepth
    );
    return shouldPrefetchTraceNodes(candidates, this.policy) ? candidates : [];
  }

  async open(): Promise<TraceOpenResult> {
    this.assertActive();
    const rootData = await this.port.getRoot();
    this.assertActive();
    this.rootTarget = { address: rootData.address };
    const root = this.placeCanonicalNode(rootData, null, 0) as TraceCellNode;
    const expansion = await this.expand(root.id);
    const updatedRoot = this.nodesById.get(root.id) as TraceCellNode;
    return {
      root: updatedRoot,
      children: expansion.children,
      truncated: expansion.truncated,
      totalNodeCount: this.nodesById.size,
    };
  }

  async expand(nodeId: string): Promise<TraceExpansionResult> {
    this.assertActive();
    const cached = this.childCache.get(nodeId);
    if (cached) return { ...cached, cached: true, children: [...cached.children] };

    const running = this.inFlight.get(nodeId);
    if (running) {
      const result = await running;
      return { ...result, cached: true, children: [...result.children] };
    }

    const promise = this.loadChildren(nodeId);
    this.inFlight.set(nodeId, promise);
    try {
      return await promise;
    } finally {
      this.inFlight.delete(nodeId);
    }
  }

  /**
   * Load a small branch into the session cache without changing dialog state.
   * A result that exceeds the prefetch budget is deliberately discarded so an
   * explicit expansion can retry with the full remaining trace budget.
   */
  async prefetch(nodeId: string, maxNewNodes: number): Promise<number> {
    this.assertActive();
    if (maxNewNodes <= 0 || this.childCache.has(nodeId)) return 0;
    const node = this.nodesById.get(nodeId);
    if (!node || node.kind !== "cell" || node.loadState !== "unloaded") return 0;
    if (node.level >= this.maxDepth) return 0;
    if (this.direction === "precedents" && node.formula.length === 0) return 0;

    const remainingNodes = Math.min(
      maxNewNodes,
      Math.max(0, this.policy.hardNodeLimit - this.nodesById.size)
    );
    if (remainingNodes === 0) return 0;

    const result = await this.port.getDirectNeighbors(this.toCellData(node), this.direction, {
      remainingNodes,
      policy: this.policy,
    });
    this.assertActive();
    if (result.truncated) return result.nodes.length;

    const children = result.nodes.map((child) => this.placeNode(child, node.id, node.level + 1));
    const parentState: TraceNodeLoadState = children.length === 0 ? "leaf" : "loaded";
    this.updateNodeState(nodeId, parentState);
    this.childCache.set(nodeId, {
      parentId: nodeId,
      parentState,
      children,
      truncated: false,
      cached: false,
      totalNodeCount: this.nodesById.size,
    });
    return children.length;
  }

  async loadRangePage(nodeId: string, page: number): Promise<TraceRangePageLoadResult> {
    this.assertActive();
    const cacheKey = `${nodeId}:${page}`;
    const cached = this.pageCache.get(cacheKey);
    if (cached) return { ...cached, children: [...cached.children] };

    const node = this.nodesById.get(nodeId);
    if (!node || node.kind !== "range") {
      throw new Error("Trace range node was not found");
    }
    if (node.level >= this.maxDepth) {
      this.updateNodeState(nodeId, "leaf");
      return {
        rangeId: nodeId,
        parentState: "leaf",
        page,
        children: [],
        hasMore: false,
        truncated: false,
        totalNodeCount: this.nodesById.size,
      };
    }

    const remainingNodes = Math.max(0, this.policy.hardNodeLimit - this.nodesById.size);
    const result = await this.port.getRangePage(this.toRangeData(node), page, {
      remainingNodes,
      policy: this.policy,
    });
    this.assertActive();
    const children = result.nodes.map((child) => this.placeNode(child, node.id, node.level + 1));
    const hasMore = result.hasMore && !result.truncated;
    const parentState: TraceNodeLoadState = "loaded";
    this.updateNodeState(nodeId, parentState);
    const out: TraceRangePageLoadResult = {
      rangeId: nodeId,
      parentState,
      page,
      children,
      hasMore,
      truncated: result.truncated,
      totalNodeCount: this.nodesById.size,
    };
    this.pageCache.set(cacheKey, out);
    return { ...out, children: [...children] };
  }

  async navigate(target: TraceTarget): Promise<void> {
    this.assertActive();
    await this.port.select(target);
    this.assertActive();
  }

  async restoreRootSelection(): Promise<void> {
    this.assertActive();
    if (!this.rootTarget) return;
    await this.port.select(this.rootTarget);
    this.assertActive();
  }

  private async loadChildren(nodeId: string): Promise<TraceExpansionResult> {
    const node = this.nodesById.get(nodeId);
    if (!node || node.kind === "reference") {
      return this.finishWithoutChildren(nodeId, "leaf", false);
    }
    if (node.kind === "range") {
      throw new Error("Range nodes must be expanded through loadRangePage");
    }
    if (node.level >= this.maxDepth) {
      return this.finishWithoutChildren(nodeId, "leaf", false);
    }
    if (this.direction === "precedents" && node.formula.length === 0) {
      return this.finishWithoutChildren(nodeId, "leaf", false);
    }

    const remainingNodes = Math.max(0, this.policy.hardNodeLimit - this.nodesById.size);
    if (remainingNodes === 0) {
      return this.finishWithoutChildren(nodeId, "leaf", true);
    }

    this.updateNodeState(nodeId, "loading");
    const result = await this.port.getDirectNeighbors(this.toCellData(node), this.direction, {
      remainingNodes,
      policy: this.policy,
    });
    this.assertActive();

    const children = result.nodes.map((child) => this.placeNode(child, node.id, node.level + 1));
    const parentState: TraceNodeLoadState = children.length === 0 ? "leaf" : "loaded";
    this.updateNodeState(nodeId, parentState);
    const out: TraceExpansionResult = {
      parentId: nodeId,
      parentState,
      children,
      truncated: result.truncated,
      cached: false,
      totalNodeCount: this.nodesById.size,
    };
    this.childCache.set(nodeId, out);
    return { ...out, children: [...children] };
  }

  private finishWithoutChildren(
    nodeId: string,
    parentState: TraceNodeLoadState,
    truncated: boolean
  ): TraceExpansionResult {
    this.updateNodeState(nodeId, parentState);
    const out: TraceExpansionResult = {
      parentId: nodeId,
      parentState,
      children: [],
      truncated,
      cached: false,
      totalNodeCount: this.nodesById.size,
    };
    this.childCache.set(nodeId, out);
    return out;
  }

  private placeCanonicalNode(
    data: TraceNodeData,
    parentId: string | null,
    level: number
  ): TraceNode {
    const loadState = this.initialLoadState(data);
    const node: TraceCellNode | TraceRangeNode = {
      ...data,
      id: data.key,
      parentId,
      level,
      loadState,
    };
    this.canonicalIdByKey.set(data.key, node.id);
    this.nodesById.set(node.id, node);
    return node;
  }

  private placeNode(data: TraceNodeData, parentId: string, level: number): TraceNode {
    const canonicalId = this.canonicalIdByKey.get(data.key);
    if (!canonicalId) return this.placeCanonicalNode(data, parentId, level);

    const canonical = this.nodesById.get(canonicalId);
    const reference: TraceReferenceNode = {
      kind: "reference",
      id: `${data.key}::ref:${++this.referenceSequence}`,
      key: data.key,
      targetId: canonicalId,
      parentId,
      level,
      worksheetName: data.worksheetName,
      address: data.address,
      value: canonical?.kind === "cell" || canonical?.kind === "reference" ? canonical.value : "",
      formula:
        canonical?.kind === "cell" || canonical?.kind === "reference" ? canonical.formula : "",
      loadState: "leaf",
    };
    this.nodesById.set(reference.id, reference);
    return reference;
  }

  private initialLoadState(data: TraceNodeData): TraceNodeLoadState {
    if (data.kind === "cell" && this.direction === "precedents" && data.formula.length === 0) {
      return "leaf";
    }
    return "unloaded";
  }

  private updateNodeState(nodeId: string, loadState: TraceNodeLoadState): void {
    const node = this.nodesById.get(nodeId);
    if (node) this.nodesById.set(nodeId, { ...node, loadState });
  }

  private toCellData(node: TraceCellNode): TraceCellData {
    return {
      kind: "cell",
      key: node.key,
      worksheetName: node.worksheetName,
      rowIndex: node.rowIndex,
      columnIndex: node.columnIndex,
      address: node.address,
      value: node.value,
      formula: node.formula,
    };
  }

  private toRangeData(node: TraceRangeNode): TraceRangeData {
    return {
      kind: "range",
      key: node.key,
      worksheetName: node.worksheetName,
      rowIndex: node.rowIndex,
      columnIndex: node.columnIndex,
      rowCount: node.rowCount,
      columnCount: node.columnCount,
      cellCount: node.cellCount,
      address: node.address,
    };
  }

  private assertActive(): void {
    if (this.cancelled) throw cancelledError();
  }
}
