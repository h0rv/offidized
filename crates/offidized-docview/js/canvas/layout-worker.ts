import { layoutDocument } from "./layout-engine.ts";
import type {
  LayoutConfig,
  LayoutResult,
  MeasuredBlock,
} from "./layout-engine.ts";
import type { SectionModel } from "../types.ts";

const LAYOUT_WORKER_PROTOCOL_NAME = "offidized-docview.layout" as const;
const LAYOUT_WORKER_PROTOCOL_VERSION = 1 as const;

interface LayoutJobRequest {
  blocks: MeasuredBlock[];
  sections: SectionModel[];
  docVersion: number;
  config?: Partial<LayoutConfig>;
}

interface LayoutWorkerRequest {
  kind: "layout.request";
  protocol: typeof LAYOUT_WORKER_PROTOCOL_NAME;
  version: number;
  requestId: number;
  payload: LayoutJobRequest;
}

type LayoutWorkerResponse =
  | {
      kind: "layout.result";
      protocol: typeof LAYOUT_WORKER_PROTOCOL_NAME;
      version: typeof LAYOUT_WORKER_PROTOCOL_VERSION;
      requestId: number;
      result: LayoutResult;
    }
  | {
      kind: "layout.error";
      protocol: typeof LAYOUT_WORKER_PROTOCOL_NAME;
      version: typeof LAYOUT_WORKER_PROTOCOL_VERSION;
      requestId: number;
      error: string;
    };

function isRequestMessage(value: unknown): value is LayoutWorkerRequest {
  if (!value || typeof value !== "object") return false;
  const msg = value as Record<string, unknown>;
  return (
    msg.kind === "layout.request" &&
    msg.protocol === LAYOUT_WORKER_PROTOCOL_NAME &&
    typeof msg.version === "number" &&
    typeof msg.requestId === "number" &&
    !!msg.payload &&
    typeof msg.payload === "object"
  );
}

const workerScope = self as unknown as {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  postMessage: (message: unknown) => void;
};

workerScope.onmessage = (event: MessageEvent<unknown>): void => {
  const raw = event.data;
  if (!isRequestMessage(raw)) {
    return;
  }

  const req = raw;

  if (req.version !== LAYOUT_WORKER_PROTOCOL_VERSION) {
    const badVersion: LayoutWorkerResponse = {
      kind: "layout.error",
      protocol: LAYOUT_WORKER_PROTOCOL_NAME,
      version: LAYOUT_WORKER_PROTOCOL_VERSION,
      requestId: req.requestId,
      error: `protocol version mismatch: got ${req.version}, expected ${LAYOUT_WORKER_PROTOCOL_VERSION}`,
    };
    workerScope.postMessage(badVersion);
    return;
  }

  try {
    const result = layoutDocument(
      req.payload.blocks,
      req.payload.sections,
      req.payload.docVersion,
      req.payload.config,
    );

    const response: LayoutWorkerResponse = {
      kind: "layout.result",
      protocol: LAYOUT_WORKER_PROTOCOL_NAME,
      version: LAYOUT_WORKER_PROTOCOL_VERSION,
      requestId: req.requestId,
      result,
    };
    workerScope.postMessage(response);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    const response: LayoutWorkerResponse = {
      kind: "layout.error",
      protocol: LAYOUT_WORKER_PROTOCOL_NAME,
      version: LAYOUT_WORKER_PROTOCOL_VERSION,
      requestId: req.requestId,
      error: message || "layout failed",
    };
    workerScope.postMessage(response);
  }
};

export {};
